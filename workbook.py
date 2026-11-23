"""Builds UK_Tenders.xlsx from the normalised rows kept in data/notices."""

from __future__ import annotations

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


def build_workbook(rows: list[dict], path: Path, settings: dict, state: dict) -> None:
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
    _write_sheet(wb, f"Pipeline ({cfg.get('pipeline_days', 90)} days)", pipeline, "Pipeline", desc_limit=800)

    wb.save(path)
