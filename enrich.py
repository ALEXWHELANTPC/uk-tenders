"""
Companies House lookup for award winners.

Gives each winning supplier its registered office address (the legally
recognised postal address for a company), company number, status and SIC
codes. Runs only when the COMPANIES_HOUSE_API_KEY environment variable is
set; without it the Award winners tab falls back to addresses printed on
the notices.

Results are cached in data/companies_house.json so each company is looked
up once and refreshed every 90 days. Companies House allows 600 requests
per five minutes; we stay well under that with a per-run cap.
"""

from __future__ import annotations

import json
import logging
import os
import re
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests

log = logging.getLogger("collector.enrich")

API = "https://api.company-information.service.gov.uk"
CACHE = Path(__file__).resolve().parent / "data" / "companies_house.json"
MAX_LOOKUPS_PER_RUN = int(os.environ.get("CH_MAX_LOOKUPS", "400"))
REFRESH_DAYS = 90

_STOP = {"ltd", "limited", "plc", "llp", "llc", "inc", "the", "co", "company", "uk", "cic", "t/a", "ta"}


def norm(name: str) -> str:
    words = re.sub(r"[^a-z0-9 ]", " ", (name or "").lower()).split()
    return " ".join(w for w in words if w not in _STOP)


def load_cache() -> dict:
    if CACHE.exists():
        return json.loads(CACHE.read_text(encoding="utf-8"))
    return {}


def save_cache(cache: dict) -> None:
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    CACHE.write_text(json.dumps(cache, indent=0, ensure_ascii=False, sort_keys=True), encoding="utf-8")


def _profile(session: requests.Session, number: str) -> dict | None:
    r = session.get(f"{API}/company/{number}", timeout=30)
    if r.status_code == 200:
        return r.json()
    if r.status_code == 429:
        log.warning("Companies House rate limit reached; pausing")
        time.sleep(60)
    return None


def _search(session: requests.Session, name: str) -> dict | None:
    r = session.get(f"{API}/search/companies", params={"q": name, "items_per_page": 5}, timeout=30)
    if r.status_code == 429:
        log.warning("Companies House rate limit reached; pausing")
        time.sleep(60)
        return None
    if r.status_code != 200:
        return None
    target = norm(name)
    for item in r.json().get("items") or []:
        # Only accept a close name match; a loose match would post letters to the wrong firm.
        if norm(item.get("title", "")) == target and item.get("company_status") in ("active", None):
            return item
    for item in r.json().get("items") or []:
        if norm(item.get("title", "")) == target:
            return item
    return None


def _record(profile: dict) -> dict:
    return {
        "company_name": profile.get("company_name"),
        "company_number": profile.get("company_number"),
        "company_status": profile.get("company_status"),
        "date_of_creation": profile.get("date_of_creation"),
        "sic_codes": profile.get("sic_codes") or [],
        "registered_office": profile.get("registered_office_address") or {},
        "looked_up": datetime.now(timezone.utc).strftime("%Y-%m-%d"),
    }


def enrich_winners(winners: list[dict]) -> dict:
    """
    winners: list of {"name": str, "company_number": str} from the award rows.
    Returns {key: record} where key is "coh:<number>" or "name:<normalised name>".
    """
    key = os.environ.get("COMPANIES_HOUSE_API_KEY", "").strip()
    cache = load_cache()
    if not key:
        log.info("No Companies House key set; Award winners will use notice addresses only")
        return cache

    session = requests.Session()
    session.auth = (key, "")
    session.headers.update({"Accept": "application/json"})
    cutoff = (datetime.now(timezone.utc) - timedelta(days=REFRESH_DAYS)).strftime("%Y-%m-%d")

    done = 0
    for w in winners:
        if done >= MAX_LOOKUPS_PER_RUN:
            log.info("Companies House: reached %s lookups this run, the rest next time", done)
            break
        number = (w.get("company_number") or "").strip().upper()
        k = f"coh:{number}" if number else f"name:{norm(w.get('name', ''))}"
        if not k.split(":", 1)[1]:
            continue
        existing = cache.get(k)
        if existing and existing.get("looked_up", "") >= cutoff:
            continue
        try:
            profile = None
            if number:
                profile = _profile(session, number)
            else:
                hit = _search(session, w["name"])
                if hit and hit.get("company_number"):
                    profile = _profile(session, hit["company_number"])
            done += 1
            if profile:
                rec = _record(profile)
                cache[k] = rec
                if rec.get("company_number"):
                    cache[f"coh:{rec['company_number']}"] = rec
            else:
                cache[k] = {"looked_up": datetime.now(timezone.utc).strftime("%Y-%m-%d"), "not_found": True}
            time.sleep(0.6)  # ~100 requests a minute, comfortably inside the limit
        except requests.RequestException as exc:
            log.warning("Companies House lookup failed for %s: %s", w.get("name"), exc)
    log.info("Companies House: %s lookups this run, %s companies cached", done, len(cache))
    save_cache(cache)
    return cache
