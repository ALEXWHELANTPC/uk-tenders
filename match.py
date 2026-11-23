"""
Trial matcher: scores collected notices against client profiles.

Run:  python match.py data/notices/2026-11.jsonl review.xlsx
"""

from __future__ import annotations

import json
import re
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

PROFILES = {
    "Space Solutions (Scotland)": {
        "regions": ["Scotland"],
        "allow_blank_region": True,
        "value_min": 50_000,
        "value_max": 25_000_000,
        "kinds": ["Tender", "Pipeline"],
        # CPV families (prefix match) and specific codes
        "cpv": {
            "45": 2, "454": 3, "4542": 4, "4533": 4, "4526": 3, "4544": 3, "4531": 3,
            "39": 3, "391": 4,
            "71": 2, "712": 4, "7122": 4, "7124": 4, "713": 2,
            "507": 3, "7033": 3, "7999": 3, "98392": 5, "6312": 4, "6310": 3, "909": 1,
        },
        "keywords": {
            "fit-out": 4, "fit out": 4, "fitout": 4, "refurbishment": 4, "refurb": 3, "interior": 3,
            "office": 3, "workplace": 4, "relocation": 4, "removal": 4, "removals": 4, "storage": 2,
            "move management": 5, "moves": 2, "furniture": 4, "joinery": 3, "joiner": 3,
            "plumbing": 3, "plumber": 3, "heating": 2, "window": 3, "windows": 3, "door replacement": 4,
            "doors": 2, "roofing": 3, "roof": 2, "bathroom": 3, "kitchen": 2, "toilet": 2,
            "minor works": 4, "architect": 3, "architectural": 3, "design team": 4, "design services": 3,
            "professional services": 2, "facilities management": 4, "building services": 2,
            "estates": 2, "refit": 3, "partition": 4, "flooring": 3, "floor covering": 4,
            "school": 1, "learning estate": 2, "modernisation": 3, "upgrade": 1, "alteration": 3,
        },
        "negatives": ["software", "catering", "agency", "vehicle", "highway", "road", "bridge", "it services",
                      "telecom", "legal", "insurance", "audit", "training", "recruitment", "medical", "waste collection", "equipment", "lift", "lifts", "fire doors", "fire door"],
    },
    "JH Contracts": {
        "regions": ["East Midlands", "West Midlands"],
        "allow_blank_region": True,
        "value_min": 0,
        "value_max": 3_000_000,
        "kinds": ["Tender", "Pipeline"],
        "cpv": {
            "45": 2, "454": 3, "4545": 4, "45453": 4, "4533": 4, "45331": 4, "453311": 5,
            "4526": 3, "4544": 3, "45443": 4, "4532": 3, "45321": 4, "45342": 4, "45233": 2,
            "507": 3, "5072": 4, "50531": 5, "7033": 3, "909": 2, "9091": 3, "4511": 2, "4521": 2, "4531": 3,
        },
        "keywords": {
            "repairs": 4, "repair": 3, "maintenance": 3, "reactive": 4, "responsive": 4, "void": 5, "voids": 5,
            "kitchen": 4, "bathroom": 4, "handyman": 5, "gas": 3, "boiler": 4, "heating": 3, "servicing": 3,
            "cladding": 5, "facade": 3, "render": 3, "damp": 4, "fencing": 4, "paving": 3, "groundworks": 3,
            "insulation": 4, "adaptations": 4, "adaptation": 4, "disabled facilities": 5, "dfg": 5,
            "decorating": 4, "painting": 3, "glazing": 3, "window": 3, "windows": 3, "doors": 2,
            "building works": 4, "minor works": 4, "housing": 3, "property": 2, "domestic": 2,
            "plastering": 4, "roofing": 1, "landscaping": 2, "grounds": 2, "cleaning": 1,
            "framework": 1, "dynamic purchasing": 2, "dps": 2, "retrofit": 3, "energy efficiency": 2,
        },
        "negatives": ["software", "consultancy", "consultant", "design services", "architect", "highway", "road",
                      "bridge", "it services", "telecom", "legal", "insurance", "audit", "training", "recruitment",
                      "medical", "vehicle", "catering", "agency", "survey", "lift", "lifts", "laundry", "vending", "coffee", "equipment", "pool", "spa", "bridges", "bridge", "grounds maintenance", "gardening", "arboricultural", "tree"],
    },
}


def _dt(v: str) -> datetime | None:
    try:
        return datetime.fromisoformat(v[:19])
    except Exception:  # noqa: BLE001
        return None


def score(row: dict, profile: dict) -> tuple[int, list[str]]:
    reasons: list[str] = []
    if row.get("notice_kind") not in profile["kinds"]:
        return 0, []
    deadline = _dt(row.get("deadline") or "")
    if row.get("notice_kind") == "Tender" and deadline and deadline < datetime.now():
        return 0, []
    region = row.get("region") or ""
    penalty = 0
    if region and region != "United Kingdom":
        if region not in profile["regions"]:
            return 0, []
        reasons.append(f"region {region}")
    elif not profile["allow_blank_region"]:
        return 0, []
    else:
        penalty = 2
        reasons.append("region not stated")
    value = row.get("value_gbp")
    if value:
        if value > profile["value_max"]:
            return 0, []
        if value < profile["value_min"]:
            return 0, []
        reasons.append(f"£{value:,.0f}")

    pts = -penalty
    codes = [c.strip() for c in (row.get("cpv_codes") or "").split(";") if c.strip()]
    best = 0
    best_code = ""
    for code in codes:
        for prefix, p in profile["cpv"].items():
            if code.startswith(prefix) and p > best:
                best, best_code = p, code
    if best:
        pts += best
        reasons.append(f"CPV {best_code}")

    title = (row.get("title") or "").lower()
    desc = (row.get("description") or "").lower()
    title_hits, desc_hits = [], []
    for kw, p in profile["keywords"].items():
        pat = r"\b" + re.escape(kw) + r"\b"
        if re.search(pat, title):
            title_hits.append((kw, p))
        elif re.search(pat, desc):
            desc_hits.append((kw, p))
    title_hits.sort(key=lambda x: -x[1])
    desc_hits.sort(key=lambda x: -x[1])
    tp = sum(p for _, p in title_hits[:3])
    dp = sum(max(1, p // 2) for _, p in desc_hits[:3])
    pts += tp + dp
    if title_hits:
        reasons.append("title: " + ", ".join(k for k, _ in title_hits[:3]))
    if desc_hits:
        reasons.append("text: " + ", ".join(k for k, _ in desc_hits[:3]))

    for neg in profile["negatives"]:
        if re.search(r"\b" + re.escape(neg) + r"\b", title):
            pts -= 6
            reasons.append(f"minus: {neg}")
            break

    if not best and not title_hits:
        return 0, []
    return pts, reasons


def main() -> None:
    src = Path(sys.argv[1])
    out = Path(sys.argv[2])
    rows = [json.loads(l) for l in src.read_text(encoding="utf-8").splitlines() if l.strip()]
    wb = Workbook()
    wb.remove(wb.active)
    summary = {}
    for name, profile in PROFILES.items():
        hits = []
        for r in rows:
            s, why = score(r, profile)
            if s >= 7:
                hits.append((s, why, r))
        best = {}
        for h in hits:
            k = re.sub(r"[^a-z0-9]", "", (h[2].get("title") or "").lower())[:60] + "|" + re.sub(r"[^a-z0-9]", "", (h[2].get("buyer") or "").lower())[:20]
            if k not in best or h[2]["published"] > best[k][2]["published"]:
                best[k] = h
        hits = list(best.values())
        hits.sort(key=lambda h: (-h[0], h[2].get("deadline") or "9"))
        summary[name] = hits
        ws = wb.create_sheet(name[:31])
        heads = ["Score", "Why", "Source", "Type", "Published", "Title", "Buyer", "Region", "Value (GBP)", "Deadline", "CPV", "Link"]
        ws.append(heads)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3A5F")
        for s, why, r in hits:
            ws.append([s, "; ".join(why), r["source"], r["notice_kind"], _dt(r["published"]), r["title"], r["buyer"], r["region"],
                       r.get("value_gbp"), _dt(r["deadline"]) if r.get("deadline") else None, r.get("cpv_codes", "")[:60], "Open notice"])
            i = ws.max_row
            if r.get("link"):
                ws.cell(row=i, column=12).hyperlink = r["link"]
                ws.cell(row=i, column=12).font = Font(color="0563C1", underline="single")
            ws.cell(row=i, column=5).number_format = "dd mmm yyyy"
            ws.cell(row=i, column=10).number_format = "dd mmm yyyy"
            ws.cell(row=i, column=9).number_format = "#,##0"
        for i, w in enumerate([7, 46, 20, 9, 12, 60, 34, 18, 13, 12, 22, 12], start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "C2"
        ws.auto_filter.ref = f"A1:L{max(2, ws.max_row)}"
    wb.save(out)
    for name, hits in summary.items():
        print(f"\n== {name}: {len(hits)} matches (score 7+), {sum(1 for h in hits if h[0] >= 12)} strong (12+)")
        for s, why, r in hits[:12]:
            print(f"  {s:>2}  {r['title'][:70]} | {r['buyer'][:32]} | {r['region']} | {r.get('value_gbp') or ''} | {(r.get('deadline') or '')[:10]}")
            print(f"      {'; '.join(why)[:150]}")


if __name__ == "__main__":
    main()
