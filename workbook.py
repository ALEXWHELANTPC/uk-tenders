"""Builds UK_Tenders.xlsx from the normalised rows kept in data/notices."""

from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

HEADERS = [
    ("source", "Source", 22),
    ("notice_kind", "Type", 11),
    ("notice_form", "Form", 10),
    ("published", "Published", 17),
    ("title", "Title", 60),
    ("buyer", "Buyer", 36),
    ("region", "Region", 22),
    ("value_gbp", "Value (GBP)", 15),
    ("cpv_codes", "CPV codes", 20),
    ("cpv_description", "CPV description", 36),
    ("deadline", "Deadline", 17),
    ("contract_start", "Contract start", 13),
    ("contract_end", "Contract end", 13),
    ("suppliers", "Suppliers", 36),
    ("procedure", "Procedure", 30),
    ("sme_suitable", "SME suitable", 12),
    ("link", "Link", 14),
    ("notice_id", "Notice ID", 24),
    ("ocid", "OCID", 30),
    ("description", "Description", 80),
]

HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
HEADER_FONT = Font(bold=True, color="FFFFFF")
LINK_FONT = Font(color="0563C1", underline="single")


def _dt(value: str) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value[:19])
    except ValueError:
        return None


def _write_sheet(wb: Workbook, title: str, rows: list[dict], table_name: str, desc_limit: int | None = None) -> None:
    ws = wb.create_sheet(title)
    ws.append([h[1] for h in HEADERS])
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")

    for row in rows:
        values = []
        for key, _, _ in HEADERS:
            v = row.get(key)
            if key in ("published", "deadline"):
                v = _dt(v) if v else None
            elif key in ("contract_start", "contract_end"):
                v = _dt(v) if v else None
            elif key == "link":
                v = "Open notice" if v else ""
            elif key == "description" and desc_limit and v and len(v) > desc_limit:
                v = v[: desc_limit - 1].rstrip() + "…"
            values.append(v)
        ws.append(values)
        r = ws.max_row
        link = row.get("link")
        if link:
            cell = ws.cell(row=r, column=17)
            cell.hyperlink = link
            cell.font = LINK_FONT
        ws.cell(row=r, column=4).number_format = "dd mmm yyyy hh:mm"
        ws.cell(row=r, column=11).number_format = "dd mmm yyyy hh:mm"
        ws.cell(row=r, column=12).number_format = "dd mmm yyyy"
        ws.cell(row=r, column=13).number_format = "dd mmm yyyy"
        ws.cell(row=r, column=8).number_format = "#,##0"

    for i, (_, _, width) in enumerate(HEADERS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "E2"

    if rows:
        ref = f"A1:{get_column_letter(len(HEADERS))}{len(rows) + 1}"
        table = Table(displayName=table_name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        ws.add_table(table)
    else:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"


WINNER_HEADERS = [
    ("name", "Supplier", 40),
    ("company_number", "Company number", 15),
    ("address_1", "Address", 40),
    ("town", "Town", 20),
    ("postcode", "Postcode", 11),
    ("country", "Country", 16),
    ("address_source", "Address source", 16),
    ("status", "Company status", 14),
    ("scale", "Size", 10),
    ("sic", "SIC codes", 24),
    ("awards", "Awards", 8),
    ("total_value", "Total value (GBP)", 16),
    ("latest_award", "Latest award", 13),
    ("latest_buyer", "Latest buyer", 34),
    ("sectors", "Sectors (CPV)", 50),
    ("regions", "Regions", 24),
    ("sources", "Sources", 26),
    ("link", "Latest notice", 14),
]

_STOP = {"ltd", "limited", "plc", "llp", "llc", "inc", "the", "co", "company", "uk", "cic", "t/a", "ta"}


def winner_key(name: str, company_number: str = "") -> str:
    if company_number:
        return "coh:" + company_number.strip().upper()
    words = re.sub(r"[^a-z0-9 ]", " ", (name or "").lower()).split()
    return "name:" + " ".join(w for w in words if w not in _STOP)


def award_winners(rows: list[dict], since: datetime, enrichment: dict | None = None) -> list[dict]:
    """One row per winning supplier across all award notices since `since`."""
    enrichment = enrichment or {}
    winners: dict[str, dict] = {}
    for r in rows:
        if r.get("notice_kind") != "Award":
            continue
        published = _dt(r.get("published"))
        if not published or published < since:
            continue
        details = r.get("supplier_details") or [{"name": n.strip()} for n in (r.get("suppliers") or "").split(";") if n.strip()]
        for s in details:
            if not s.get("name"):
                continue
            key = winner_key(s["name"], s.get("company_number", ""))
            w = winners.get(key)
            if w is None:
                w = winners[key] = {
                    "name": s["name"], "company_number": s.get("company_number", ""),
                    "street": "", "town": "", "postcode": "", "country": "", "scale": "",
                    "awards": 0, "total_value": 0.0, "latest_award": None, "latest_buyer": "", "link": "",
                    "sectors": set(), "regions": set(), "sources": set(),
                }
            if not w["postcode"] and s.get("postcode"):
                w.update({k: s.get(k, "") for k in ("street", "town", "postcode", "country")})
            elif not w["street"] and s.get("street"):
                w["street"], w["town"] = s.get("street", ""), s.get("town", "")
            if not w["company_number"] and s.get("company_number"):
                w["company_number"] = s["company_number"]
            if s.get("scale") and not w["scale"]:
                w["scale"] = s["scale"]
            w["awards"] += 1
            if r.get("value_gbp"):
                w["total_value"] += float(r["value_gbp"])
            if w["latest_award"] is None or published > w["latest_award"]:
                w["latest_award"], w["latest_buyer"], w["link"] = published, r.get("buyer", ""), r.get("link", "")
            for d in (r.get("cpv_description") or "").split(";"):
                if d.strip():
                    w["sectors"].add(d.strip())
            if r.get("region"):
                w["regions"].add(r["region"])
            w["sources"].add(r.get("source", ""))

    out: list[dict] = []
    for key, w in winners.items():
        ch = enrichment.get(key) or (enrichment.get("coh:" + w["company_number"]) if w["company_number"] else None) or {}
        reg = ch.get("registered_office") or {}
        if reg.get("postal_code"):
            address_1 = ", ".join(x for x in (reg.get("address_line_1"), reg.get("address_line_2")) if x)
            town, postcode, country, source = reg.get("locality", ""), reg.get("postal_code", ""), reg.get("country", ""), "Companies House"
        else:
            address_1, town, postcode, country = w["street"], w["town"], w["postcode"], w["country"]
            source = "Notice" if (postcode or address_1) else "Not found"
        out.append({
            "name": ch.get("company_name") or w["name"],
            "company_number": ch.get("company_number") or w["company_number"],
            "address_1": address_1, "town": town, "postcode": postcode, "country": country,
            "address_source": source,
            "status": ch.get("company_status", ""),
            "scale": {"sme": "SME", "large": "Large", "micro": "Micro", "medium": "Medium", "small": "Small"}.get((w["scale"] or "").lower(), w["scale"]),
            "sic": "; ".join(ch.get("sic_codes") or []),
            "awards": w["awards"],
            "total_value": round(w["total_value"]) if w["total_value"] else None,
            "latest_award": w["latest_award"],
            "latest_buyer": w["latest_buyer"],
            "sectors": "; ".join(sorted(w["sectors"]))[:300],
            "regions": "; ".join(sorted(w["regions"])),
            "sources": "; ".join(sorted(w["sources"])),
            "link": w["link"],
        })
    out.sort(key=lambda x: (x["latest_award"] or datetime.min), reverse=True)
    return out


def _write_winners(wb: Workbook, rows: list[dict]) -> None:
    ws = wb.create_sheet("Award winners")
    ws.append([h[1] for h in WINNER_HEADERS])
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    for r in rows:
        ws.append([("Open notice" if k == "link" and r.get("link") else r.get(k)) for k, _, _ in WINNER_HEADERS])
        i = ws.max_row
        if r.get("link"):
            c = ws.cell(row=i, column=len(WINNER_HEADERS))
            c.hyperlink = r["link"]
            c.font = LINK_FONT
        ws.cell(row=i, column=12).number_format = "#,##0"
        ws.cell(row=i, column=13).number_format = "dd mmm yyyy"
    for i, (_, _, width) in enumerate(WINNER_HEADERS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "B2"
    if rows:
        table = Table(displayName="AwardWinners", ref=f"A1:{get_column_letter(len(WINNER_HEADERS))}{len(rows) + 1}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        ws.add_table(table)


def build_workbook(rows: list[dict], path: Path, settings: dict, state: dict, enrichment: dict | None = None) -> None:
    cfg = settings["workbook"]
    now = datetime.now(timezone.utc).replace(tzinfo=None)
    new_since = now - timedelta(days=cfg.get("new_days", 7))
    awards_since = now - timedelta(days=cfg.get("awards_days", 90))
    pipeline_since = now - timedelta(days=cfg.get("pipeline_days", 90))

    def published(r: dict) -> datetime:
        return _dt(r["published"]) or datetime.min

    open_tenders = [
        r for r in rows
        if r["notice_kind"] == "Tender" and (_dt(r.get("deadline")) or datetime.max) >= now
    ]
    # One row per ocid for open tenders: keep the most recent release.
    latest: dict[str, dict] = {}
    for r in sorted(open_tenders, key=published):
        latest[r.get("ocid") or r["_key"]] = r
    open_tenders = sorted(latest.values(), key=lambda r: (_dt(r.get("deadline")) or datetime.max, r["title"]))

    new_rows = sorted((r for r in rows if published(r) >= new_since), key=published, reverse=True)
    awards = sorted((r for r in rows if r["notice_kind"] == "Award" and published(r) >= awards_since), key=published, reverse=True)
    pipeline = sorted((r for r in rows if r["notice_kind"] == "Pipeline" and published(r) >= pipeline_since), key=published, reverse=True)

    wb = Workbook()
    about = wb.active
    about.title = "About"
    about.column_dimensions["A"].width = 34
    about.column_dimensions["B"].width = 90
    info = [
        ("UK Tenders", ""),
        ("Last updated (UTC)", now.strftime("%d %b %Y %H:%M")),
        ("Notices held", len(rows)),
        ("Open tenders", len(open_tenders)),
        (f"New in last {cfg.get('new_days', 7)} days", len(new_rows)),
        (f"Awards in last {cfg.get('awards_days', 90)} days", len(awards)),
        ("Award winners (companies)", len(award_winners(rows, awards_since, enrichment))),
        (f"Pipeline notices in last {cfg.get('pipeline_days', 90)} days", len(pipeline)),
        ("", ""),
        ("Sources", ""),
    ]
    for label, stats in (state.get("last_totals") or {}).items():
        info.append((label, f"{stats.get('releases', 0)} notices read, {stats.get('new', 0)} new, {stats.get('updated', 0)} updated, {stats.get('errors', 0)} errors"))
    for key, st in (state.get("sources") or {}).items():
        label = settings["sources"].get(key, {}).get("label", key)
        info.append((f"{label} last success", (st.get("last_success") or "never").replace("T", " ")))
    info += [
        ("", ""),
        ("How to use", "This file is rewritten every hour by the collector. Keep your own notes and client matching in a separate workbook that looks up this one, so they are not overwritten."),
        ("Type", "Tender = open opportunity. Award = contract awarded (shows winning suppliers). Pipeline = early notice of planned procurement. Update = change to an earlier notice."),
        ("Region", "Taken from the delivery address where given, otherwise the buyer's address."),
        ("Award winners", "One row per company that has won a contract in the awards window, with a postal address for direct mail. Address source says where the address came from: Companies House (registered office, the legally recognised address for post) when the lookup is on, otherwise the address printed on the notice. Not found means the notice carried no address and no Companies House match was made."),
        ("Data", "Open Government Licence data from Find a Tender, Contracts Finder, Public Contracts Scotland and Sell2Wales."),
    ]
    for label, value in info:
        about.append([label, value])
    about["A1"].font = Font(bold=True, size=14)
    for row in about.iter_rows(min_row=2):
        row[0].font = Font(bold=True)
        row[1].alignment = Alignment(wrap_text=True, vertical="top")

    _write_sheet(wb, "Open tenders", open_tenders, "OpenTenders")
    _write_sheet(wb, f"New last {cfg.get('new_days', 7)} days", new_rows, "NewNotices")
    _write_sheet(wb, f"Awards ({cfg.get('awards_days', 90)} days)", awards, "Awards", desc_limit=400)
    _write_winners(wb, award_winners(rows, awards_since, enrichment))
    _write_sheet(wb, f"Pipeline ({cfg.get('pipeline_days', 90)} days)", pipeline, "Pipeline", desc_limit=800)

    wb.save(path)
