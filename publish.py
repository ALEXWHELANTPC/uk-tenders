"""
Pushes collected notices into the Tender Pro client portal (Supabase) and
scores them against each client's matching profile.

Runs only when SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are set. Without
them the collector just builds the workbook as before.

Tables (see the portal's supabase/migrations/012_notice_matching.sql):
  notices          one row per notice, upserted on key
  client_profiles  what each client is looking for
  notice_matches   a notice that scored above a client's threshold
"""

from __future__ import annotations

import logging
import os
import re
from datetime import datetime, timedelta, timezone

import requests

from match import score

log = logging.getLogger("collector.publish")

NOTICE_FIELDS = [
    "source", "notice_kind", "notice_form", "published", "title", "buyer", "region", "value_gbp",
    "cpv_codes", "cpv_description", "deadline", "contract_start", "contract_end", "suppliers",
    "supplier_details", "procedure", "sme_suitable", "link", "notice_id", "ocid", "description",
]
CHUNK = 400
KEEP_DAYS = 183


class Portal:
    def __init__(self, url: str, key: str):
        self.base = url.rstrip("/") + "/rest/v1"
        self.s = requests.Session()
        self.s.headers.update({
            "apikey": key,
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
        })

    def upsert_notices(self, rows: list[dict]) -> dict[str, str]:
        """Upsert rows on key. Returns {key: notice uuid}."""
        ids: dict[str, str] = {}
        for i in range(0, len(rows), CHUNK):
            batch = []
            for r in rows[i:i + CHUNK]:
                row = {k: (r.get(k) if r.get(k) not in ("",) else None) for k in NOTICE_FIELDS}
                row["key"] = r["_key"]
                batch.append(row)
            resp = self.s.post(
                f"{self.base}/notices?on_conflict=key&select=id,key",
                json=batch,
                headers={"Prefer": "resolution=merge-duplicates,return=representation"},
                timeout=120,
            )
            if resp.status_code >= 300:
                log.error("notices upsert failed (%s): %s", resp.status_code, resp.text[:300])
                continue
            for item in resp.json():
                ids[item["key"]] = item["id"]
        return ids

    def profiles(self) -> list[dict]:
        resp = self.s.get(f"{self.base}/client_profiles?enabled=eq.true&select=*", timeout=60)
        if resp.status_code != 200:
            log.error("could not read client profiles (%s): %s", resp.status_code, resp.text[:300])
            return []
        return resp.json()

    def existing_match_keys(self) -> set[tuple[str, str]]:
        """(company_id, tender key) for every match already recorded, so the same
        tender published twice (two portals, or a corrected notice) is not
        offered to a client twice."""
        out: set[tuple[str, str]] = set()
        offset = 0
        while True:
            resp = self.s.get(
                f"{self.base}/notice_matches?select=company_id,notices(title,buyer)"
                f"&offset={offset}&limit=1000",
                timeout=60,
            )
            if resp.status_code != 200:
                log.warning("could not read existing matches (%s)", resp.status_code)
                return out
            rows = resp.json()
            for m in rows:
                n = m.get("notices") or {}
                out.add((m["company_id"], tender_key(n)))
            if len(rows) < 1000:
                return out
            offset += 1000

    def insert_matches(self, rows: list[dict]) -> int:
        done = 0
        for i in range(0, len(rows), CHUNK):
            resp = self.s.post(
                f"{self.base}/notice_matches?on_conflict=company_id,notice_id",
                json=rows[i:i + CHUNK],
                headers={"Prefer": "resolution=ignore-duplicates,return=minimal"},
                timeout=120,
            )
            if resp.status_code >= 300:
                log.error("notice_matches insert failed (%s): %s", resp.status_code, resp.text[:300])
            else:
                done += len(rows[i:i + CHUNK])
        return done

    def prune(self) -> None:
        cutoff = (datetime.now(timezone.utc) - timedelta(days=KEEP_DAYS)).strftime("%Y-%m-%dT%H:%M:%S")
        resp = self.s.delete(f"{self.base}/notices?published=lt.{cutoff}", headers={"Prefer": "return=minimal"}, timeout=120)
        if resp.status_code >= 300:
            log.warning("prune failed (%s): %s", resp.status_code, resp.text[:200])


def tender_key(row: dict) -> str:
    """Same tender on two portals, or re-issued, shares title and buyer."""
    t = re.sub(r"[^a-z0-9]", "", (row.get("title") or "").lower())[:60]
    b = re.sub(r"[^a-z0-9]", "", (row.get("buyer") or "").lower().replace("&", " and "))[:20]
    return f"{t}|{b}"


def profile_from_row(row: dict) -> dict:
    """DB row -> the dict shape match.score expects."""
    return {
        "regions": row.get("regions") or [],
        "allow_blank_region": bool(row.get("allow_blank_region", True)),
        "value_min": float(row["value_min"]) if row.get("value_min") is not None else 0,
        "value_max": float(row["value_max"]) if row.get("value_max") is not None else 1e12,
        "kinds": row.get("kinds") or ["Tender", "Pipeline"],
        "cpv": {str(k): int(v) for k, v in (row.get("cpv") or {}).items()},
        "keywords": {str(k): int(v) for k, v in (row.get("keywords") or {}).items()},
        "negatives": row.get("negatives") or [],
        "threshold": int(row.get("threshold") or 7),
    }


def publish(touched: list[dict]) -> dict:
    """
    touched: normalised rows that were new or changed this run.
    Returns a summary dict for the log.
    """
    url = os.environ.get("SUPABASE_URL", "").strip()
    key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "").strip()
    if not url or not key:
        log.info("Portal publishing is off (no SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY)")
        return {"enabled": False}

    portal = Portal(url, key)
    summary = {"enabled": True, "notices": 0, "profiles": 0, "matches": 0}
    if not touched:
        log.info("Portal: nothing new to publish")
        portal.prune()
        return summary

    ids = portal.upsert_notices(touched)
    summary["notices"] = len(ids)
    log.info("Portal: %s notices upserted", len(ids))

    profiles = portal.profiles()
    summary["profiles"] = len(profiles)
    seen = portal.existing_match_keys() if profiles else set()
    matches: list[dict] = []
    for prow in profiles:
        profile = profile_from_row(prow)
        best: dict[str, tuple[int, dict, dict]] = {}
        for r in touched:
            nid = ids.get(r["_key"])
            if not nid:
                continue
            s, reasons = score(r, profile)
            if s < profile["threshold"]:
                continue
            tk = tender_key(r)
            if (prow["company_id"], tk) in seen:
                continue
            match = {
                "company_id": prow["company_id"],
                "notice_id": nid,
                "score": int(s),
                "reasons": reasons,
            }
            # Keep one notice per tender: the higher score, then the one with a deadline.
            cur = best.get(tk)
            rank = (int(s), 1 if r.get("deadline") else 0, r.get("published") or "")
            if cur is None or rank > cur[0]:
                best[tk] = (rank, match, r)
        for _rank, match, _r in best.values():
            matches.append(match)
    if matches:
        summary["matches"] = portal.insert_matches(matches)
    log.info("Portal: %s profiles scored, %s matches written", len(profiles), summary["matches"])
    portal.prune()
    return summary
