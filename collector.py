#!/usr/bin/env python3
"""
UK tender collector.

Pulls every notice published on the UK public procurement portals that offer
an open feed (Find a Tender, Contracts Finder, Public Contracts Scotland,
Sell2Wales), normalises each one to a flat row, keeps a running history in
monthly JSONL files under data/, and writes UK_Tenders.xlsx.

Run with no arguments. Settings live in settings.json next to this file.
"""

from __future__ import annotations

import json
import logging
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterable

import requests

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
HISTORY = DATA / "notices"
STATE_FILE = DATA / "state.json"

log = logging.getLogger("collector")

UA = "uk-tenders-collector/1.0 (open data collector; GitHub Actions)"
TIMEOUT = 180
RETRIES = 3
# Stop asking the feeds for more after this long and save what we have. The
# next run carries on from where this one stopped. Keeps well inside the
# job's own timeout even when a feed is answering slowly.
TIME_BUDGET = int(os.environ.get("TIME_BUDGET_SECONDS", "2100"))

# ---------------------------------------------------------------------------
# Row shape. Keep this list in step with workbook.py.
# ---------------------------------------------------------------------------
COLUMNS = [
    "source",          # portal label
    "notice_kind",     # Tender / Award / Pipeline / Update / Other
    "notice_form",     # UK4, F02, contractNotice ...
    "published",       # ISO date-time
    "title",
    "buyer",
    "region",
    "value_gbp",
    "cpv_codes",
    "cpv_description",
    "deadline",        # ISO date-time, tenders only
    "contract_start",
    "contract_end",
    "suppliers",
    "procedure",
    "sme_suitable",
    "link",
    "notice_id",
    "ocid",
    "description",
]

NUTS_NAMES = {
    "UKC": "North East England",
    "UKD": "North West England",
    "UKE": "Yorkshire and the Humber",
    "UKF": "East Midlands",
    "UKG": "West Midlands",
    "UKH": "East of England",
    "UKI": "London",
    "UKJ": "South East England",
    "UKK": "South West England",
    "UKL": "Wales",
    "UKM": "Scotland",
    "UKN": "Northern Ireland",
    "UK": "United Kingdom",
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def utcnow() -> datetime:
    return datetime.now(timezone.utc).replace(microsecond=0)


def iso(dt: datetime) -> str:
    return dt.strftime("%Y-%m-%dT%H:%M:%S")


def parse_dt(value: Any) -> datetime | None:
    if not value or not isinstance(value, str):
        return None
    v = value.strip()
    if v.endswith("Z"):
        v = v[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(v)
    except ValueError:
        try:
            dt = datetime.fromisoformat(v[:19])
        except ValueError:
            return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def clean(text: Any, limit: int | None = None) -> str:
    if text is None:
        return ""
    s = str(text)
    s = re.sub(r"\s+", " ", s).strip()
    if limit and len(s) > limit:
        s = s[: limit - 1].rstrip() + "…"
    return s


def first(*values: Any) -> Any:
    for v in values:
        if v not in (None, "", [], {}):
            return v
    return None


def region_name(code: str | None) -> str:
    if not code:
        return ""
    c = code.strip().upper()
    if not c.startswith("UK"):
        return code
    for n in (3, 2):
        if c[:n] in NUTS_NAMES:
            return NUTS_NAMES[c[:n]]
    return code


def get(session: requests.Session, url: str, params: dict | None = None) -> requests.Response | None:
    """GET with retries. Returns None when the server keeps failing."""
    for attempt in range(1, RETRIES + 1):
        try:
            r = session.get(url, params=params, timeout=TIMEOUT)
            if r.status_code == 200:
                return r
            log.warning("HTTP %s from %s (attempt %s)", r.status_code, r.url, attempt)
            if 400 <= r.status_code < 500 and r.status_code != 429:
                return None
        except requests.RequestException as exc:
            log.warning("Request failed: %s (attempt %s)", exc, attempt)
        time.sleep(5 * attempt)
    return None


# ---------------------------------------------------------------------------
# OCDS release -> flat row
# ---------------------------------------------------------------------------
def kind_from_tags(tags: Iterable[str], form: str) -> str:
    tags = set(tags or [])
    f = (form or "").upper()
    if f in ("UK1", "UK2", "UK3") or "planning" in tags:
        return "Pipeline"
    if tags & {"award", "contract"} or f in ("UK6", "UK7"):
        return "Award"
    if "tender" in tags or f == "UK4":
        return "Tender"
    if tags & {"tenderUpdate", "tenderAmendment", "tenderCancellation", "awardUpdate",
               "awardCancellation", "contractUpdate", "contractAmendment"}:
        return "Update"
    if "contractTermination" in tags:
        return "Contract end"
    return "Other"


def buyer_region(release: dict) -> str:
    tender = release.get("tender") or {}
    # Delivery region first: it says where the work is.
    for item in tender.get("items") or []:
        for addr in item.get("deliveryAddresses") or []:
            if addr.get("region"):
                return region_name(addr["region"])
    buyer_id = (release.get("buyer") or {}).get("id")
    for party in release.get("parties") or []:
        roles = party.get("roles") or []
        if "buyer" in roles or party.get("id") == buyer_id:
            addr = party.get("address") or {}
            if addr.get("region"):
                return region_name(addr["region"])
            if addr.get("locality"):
                return clean(addr["locality"])
    return ""


def cpv_codes(release: dict) -> tuple[str, str]:
    tender = release.get("tender") or {}
    codes: list[str] = []
    descs: list[str] = []

    def add(cl: dict | None):
        if not cl or cl.get("scheme", "CPV").upper() != "CPV":
            return
        cid = str(cl.get("id") or "").strip()
        if cid and cid not in codes:
            codes.append(cid)
            if cl.get("description"):
                descs.append(clean(cl["description"]))

    add(tender.get("classification"))
    for item in tender.get("items") or []:
        add(item.get("classification"))
        for cl in item.get("additionalClassifications") or []:
            add(cl)
    for cl in tender.get("additionalClassifications") or []:
        add(cl)
    return "; ".join(codes), "; ".join(dict.fromkeys(descs))


def value_gbp(release: dict) -> float | None:
    tender = release.get("tender") or {}
    candidates = [tender.get("value")]
    for a in release.get("awards") or []:
        candidates.append(a.get("value"))
    for c in release.get("contracts") or []:
        candidates.append(c.get("value"))
    for lot in tender.get("lots") or []:
        candidates.append(lot.get("value"))
    for v in candidates:
        if not v:
            continue
        amt = first(v.get("amountGross"), v.get("amount"), v.get("maximumValue"))
        if amt is None:
            continue
        try:
            return float(amt)
        except (TypeError, ValueError):
            continue
    return None


def contract_period(release: dict) -> tuple[str, str]:
    tender = release.get("tender") or {}
    periods = [tender.get("contractPeriod")]
    for lot in tender.get("lots") or []:
        periods.append(lot.get("contractPeriod"))
    for a in release.get("awards") or []:
        periods.append(a.get("contractPeriod"))
    for c in release.get("contracts") or []:
        periods.append(c.get("period"))
    for p in periods:
        if p and (p.get("startDate") or p.get("endDate")):
            return (p.get("startDate") or "")[:10], (p.get("endDate") or "")[:10]
    return "", ""


def suppliers(release: dict) -> str:
    names: list[str] = []
    for a in release.get("awards") or []:
        for s in a.get("suppliers") or []:
            n = clean(s.get("name"))
            if n and n not in names:
                names.append(n)
    if not names:
        for p in release.get("parties") or []:
            if "supplier" in (p.get("roles") or []):
                n = clean(p.get("name"))
                if n and n not in names:
                    names.append(n)
    return "; ".join(names)


def notice_link(release: dict, source_key: str, settings: dict) -> str:
    tender = release.get("tender") or {}
    docs = list(tender.get("documents") or [])
    for a in release.get("awards") or []:
        docs.extend(a.get("documents") or [])
    for d in docs:
        url = d.get("url") or ""
        if "/Notice/" in url or "search_view.aspx" in url:
            return url
    rid = str(release.get("id") or "")
    if source_key == "find_a_tender" and rid:
        return f"https://www.find-tender.service.gov.uk/Notice/{rid}"
    if source_key == "contracts_finder" and rid:
        # CF release ids look like <notice guid>-<number>
        guid = rid.rsplit("-", 1)[0] if re.search(r"-\d+$", rid) else rid
        return f"https://www.contractsfinder.service.gov.uk/Notice/{guid}"
    tmpl = settings["sources"][source_key].get("notice_link")
    if tmpl:
        for d in docs:
            if d.get("id"):
                return tmpl.format(id=d["id"])
        m = re.search(r"([A-Z]{3}\d{6,})", rid)
        if m:
            return tmpl.format(id=m.group(1))
    for d in docs:
        if d.get("url"):
            return d["url"]
    return ""


def notice_form(release: dict) -> str:
    tender = release.get("tender") or {}
    for d in tender.get("documents") or []:
        if d.get("noticeType"):
            return str(d["noticeType"])
    for d in tender.get("documents") or []:
        if d.get("documentType") and "Notice" in d["documentType"]:
            return str(d["documentType"])
    for a in release.get("awards") or []:
        for d in a.get("documents") or []:
            if d.get("documentType"):
                return str(d["documentType"])
    return ""


def normalise(release: dict, source_key: str, settings: dict) -> dict | None:
    tender = release.get("tender") or {}
    rid = str(release.get("id") or "").strip()
    ocid = str(release.get("ocid") or "").strip()
    if not rid and not ocid:
        return None
    published = parse_dt(release.get("date"))
    if not published:
        return None
    title = clean(first(tender.get("title"), release.get("title"), (release.get("planning") or {}).get("rationale")), 300)
    if not title:
        title = clean(release.get("description"), 300) or "(untitled)"
    buyer = clean((release.get("buyer") or {}).get("name"))
    if not buyer:
        for p in release.get("parties") or []:
            if "buyer" in (p.get("roles") or []):
                buyer = clean(p.get("name"))
                break
    form = notice_form(release)
    codes, cpv_desc = cpv_codes(release)
    start, end = contract_period(release)
    deadline = parse_dt((tender.get("tenderPeriod") or {}).get("endDate"))
    sme = (tender.get("suitability") or {}).get("sme")
    for lot in tender.get("lots") or []:
        if sme is None:
            sme = (lot.get("suitability") or {}).get("sme")
    desc_limit = settings["workbook"].get("description_chars", 1500)
    description = clean(first(tender.get("description"), release.get("description")), desc_limit)

    return {
        "source": settings["sources"][source_key]["label"],
        "notice_kind": kind_from_tags(release.get("tag") or [], form),
        "notice_form": form,
        "published": iso(published),
        "title": title,
        "buyer": buyer,
        "region": buyer_region(release),
        "value_gbp": value_gbp(release),
        "cpv_codes": codes,
        "cpv_description": cpv_desc,
        "deadline": iso(deadline) if deadline else "",
        "contract_start": start,
        "contract_end": end,
        "suppliers": suppliers(release),
        "procedure": clean(first(tender.get("procurementMethodDetails"), tender.get("procurementMethod")), 200),
        "sme_suitable": "" if sme is None else ("Yes" if sme else "No"),
        "link": notice_link(release, source_key, settings),
        "notice_id": rid or ocid,
        "ocid": ocid,
        "description": description,
        "_key": f"{source_key}|{rid or ocid}",
    }


# ---------------------------------------------------------------------------
# Fetchers. Each yields OCDS releases.
# ---------------------------------------------------------------------------
def out_of_time(deadline: float) -> bool:
    return time.monotonic() > deadline


def fetch_cursor_feed(session: requests.Session, url: str, params: dict, stats: dict, deadline: float) -> Iterable[dict]:
    """Find a Tender and Contracts Finder: first page by params, then links.next."""
    next_url: str | None = url
    first_page = True
    pages = 0
    while next_url:
        if out_of_time(deadline):
            log.warning("Time budget used up after %s pages; will carry on next run", pages)
            stats["partial"] = True
            return
        r = get(session, next_url, params if first_page else None)
        first_page = False
        if r is None:
            stats["errors"] += 1
            return
        try:
            body = r.json()
        except ValueError:
            log.warning("Non-JSON reply from %s", r.url)
            stats["errors"] += 1
            return
        releases = body.get("releases") or []
        pages += 1
        stats["pages"] += 1
        if pages % 10 == 0:
            log.info("  page %s, %s releases so far", pages, stats["releases"])
        for rel in releases:
            yield rel
        next_url = (body.get("links") or {}).get("next")
        if not releases:
            break
        if pages > 400:
            log.warning("Stopping after 400 pages from %s", url)
            break


def fetch_windowed(session: requests.Session, url: str, from_param: str, to_param: str,
                   since: datetime, until: datetime, window_hours: int, page_size: int,
                   stats: dict, deadline: float) -> Iterable[dict]:
    """
    Walk a cursor feed in time windows, oldest first. Both feeds return newest
    first, so paging the whole range in one go would leave the oldest notices
    unfetched if the run stopped early. Working window by window means a run
    that stops early has still fully covered everything up to the last window
    it finished, and stats["covered_to"] says where to resume.
    """
    start = since
    window = timedelta(hours=window_hours)
    while start < until:
        end = min(start + window, until)
        if out_of_time(deadline):
            log.warning("Time budget used up; covered up to %s, will carry on next run", iso(start))
            stats["partial"] = True
            return
        params = {from_param: iso(start - timedelta(minutes=1)), to_param: iso(end), "limit": page_size}
        before = stats["errors"]
        for rel in fetch_cursor_feed(session, url, params, stats, deadline):
            yield rel
        if stats["partial"] or stats["errors"] > before:
            stats["partial"] = True
            return
        stats["covered_to"] = iso(end)
        start = end


def fetch_monthly_feed(session: requests.Session, url: str, months: list[str], notice_types: list[int], stats: dict, deadline: float) -> Iterable[dict]:
    """Public Contracts Scotland and Sell2Wales: one call per month and notice type."""
    for month in months:
        for nt in notice_types:
            if out_of_time(deadline):
                log.warning("Time budget used up; will carry on next run")
                stats["partial"] = True
                return
            r = get(session, url, {"dateFrom": month, "noticeType": nt, "outputType": 0})
            if r is None:
                stats["errors"] += 1
                continue
            try:
                body = r.json()
            except ValueError:
                log.warning("Non-JSON reply for %s type %s", month, nt)
                stats["errors"] += 1
                continue
            stats["pages"] += 1
            for rel in body.get("releases") or []:
                yield rel


def months_covering(since: datetime, until: datetime) -> list[str]:
    out: list[str] = []
    y, m = since.year, since.month
    while (y, m) <= (until.year, until.month):
        out.append(f"{m:02d}-{y}")
        m += 1
        if m == 13:
            y, m = y + 1, 1
    return out


# ---------------------------------------------------------------------------
# Storage: monthly JSONL files keyed by _key
# ---------------------------------------------------------------------------
def month_of(row: dict) -> str:
    return row["published"][:7]


def load_month(month: str) -> dict[str, dict]:
    path = HISTORY / f"{month}.jsonl"
    rows: dict[str, dict] = {}
    if path.exists():
        with path.open(encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if line:
                    row = json.loads(line)
                    rows[row["_key"]] = row
    return rows


def save_month(month: str, rows: dict[str, dict]) -> None:
    HISTORY.mkdir(parents=True, exist_ok=True)
    path = HISTORY / f"{month}.jsonl"
    with path.open("w", encoding="utf-8") as fh:
        for key in sorted(rows):
            fh.write(json.dumps(rows[key], ensure_ascii=False, separators=(",", ":")) + "\n")


def load_all(keep_months: int) -> list[dict]:
    rows: list[dict] = []
    if not HISTORY.exists():
        return rows
    files = sorted(HISTORY.glob("*.jsonl"))[-keep_months:]
    for path in files:
        with path.open(encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if line:
                    rows.append(json.loads(line))
    return rows


def load_state() -> dict:
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    return {"sources": {}}


def save_state(state: dict) -> None:
    DATA.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(json.dumps(state, indent=2), encoding="utf-8")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def run() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", stream=sys.stdout)
    settings = json.loads((ROOT / "settings.json").read_text(encoding="utf-8"))
    state = load_state()
    now = utcnow()
    force = os.environ.get("FORCE_ALL") == "1"

    session = requests.Session()
    session.headers.update({"User-Agent": UA, "Accept": "application/json"})

    changed_months: dict[str, dict[str, dict]] = {}
    totals: dict[str, dict] = {}
    deadline = time.monotonic() + TIME_BUDGET

    for key, src in settings["sources"].items():
        if not src.get("enabled", True):
            continue
        st = state["sources"].setdefault(key, {})
        last_ok = parse_dt(st.get("last_success"))
        every = src.get("run_every_hours", 1)
        if last_ok and not force and not st.get("last_partial") and (now - last_ok) < timedelta(hours=every) - timedelta(minutes=5):
            log.info("%s: ran %s ago, skipping this hour", src["label"], now - last_ok)
            continue

        since = parse_dt(st.get("covered_to"))
        if since:
            since = since - timedelta(minutes=settings.get("overlap_minutes", 90))
        else:
            since = now - timedelta(days=settings.get("first_run_days_back", 14))

        if out_of_time(deadline):
            log.warning("%s: no time left this run, skipping", src["label"])
            continue

        stats = {"pages": 0, "errors": 0, "releases": 0, "new": 0, "updated": 0, "partial": False}
        log.info("%s: collecting from %s", src["label"], iso(since))

        window_hours = src.get("window_hours", 24)
        if key == "find_a_tender":
            releases = fetch_windowed(session, src["url"], "updatedFrom", "updatedTo", since, now, window_hours, src.get("page_size", 100), stats, deadline)
        elif key == "contracts_finder":
            releases = fetch_windowed(session, src["url"], "publishedFrom", "publishedTo", since, now, window_hours, src.get("page_size", 100), stats, deadline)
        else:
            releases = fetch_monthly_feed(session, src["url"], months_covering(since, now), src["notice_types"], stats, deadline)

        for rel in releases:
            stats["releases"] += 1
            row = normalise(rel, key, settings)
            if not row:
                continue
            month = month_of(row)
            bucket = changed_months.get(month)
            if bucket is None:
                bucket = changed_months[month] = load_month(month)
            old = bucket.get(row["_key"])
            if old is None:
                stats["new"] += 1
            elif old != row:
                stats["updated"] += 1
            bucket[row["_key"]] = row

        totals[src["label"]] = stats
        log.info("%s: %s", src["label"], json.dumps(stats))

        complete = stats["errors"] == 0 and not stats["partial"]
        if complete:
            st["last_success"] = iso(now)
            st["covered_to"] = iso(now)
            st.pop("last_partial", None)
        elif stats["releases"] > 0 or stats.get("covered_to"):
            # Got some of the way. Only claim coverage up to the end of the
            # last window fully read, so the next run picks up the rest.
            st["last_success"] = iso(now)
            st["last_partial"] = iso(now)
            if stats.get("covered_to"):
                st["covered_to"] = stats["covered_to"]
        else:
            st["last_failure"] = iso(now)
            log.error("%s: nothing collected", src["label"])

    for month, rows in changed_months.items():
        save_month(month, rows)

    # Trim months beyond the keep window
    keep = settings.get("keep_months", 18)
    if HISTORY.exists():
        for old in sorted(HISTORY.glob("*.jsonl"))[:-keep]:
            old.unlink()

    state["last_run"] = iso(now)
    state["last_totals"] = totals
    save_state(state)

    from workbook import build_workbook  # local module

    all_rows = load_all(keep)
    build_workbook(all_rows, ROOT / settings["workbook"]["file"], settings, state)
    log.info("Workbook written with %s notices in history", len(all_rows))
    return 0


if __name__ == "__main__":
    sys.exit(run())
