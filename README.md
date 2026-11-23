# UK tenders collector

Collects every notice published through the UK public procurement feeds that offer open data, and keeps them in one Excel workbook.

Sources: Find a Tender, Contracts Finder, Public Contracts Scotland and Sell2Wales. All of it is Open Government Licence data. eTendersNI has no public feed; above-threshold Northern Ireland notices still arrive through Find a Tender.

## What you get

`UK_Tenders.xlsx` is rewritten after every run and has these tabs:

Open tenders: live opportunities, soonest deadline first, one row per procurement.
New last 7 days: every notice of any type published in the last week.
Awards (90 days): contract awards with the winning suppliers.
Pipeline (90 days): early notices of planned procurements.
About: when it last ran and what each source returned.

Columns: Source, Type, Form, Published, Title, Buyer, Region, Value (GBP), CPV codes, CPV description, Deadline, Contract start, Contract end, Suppliers, Procedure, SME suitable, Link, Notice ID, OCID, Description.

The full history is kept in `data/notices/YYYY-MM.jsonl`, one line per notice, for 18 months.

## How it runs

A GitHub Actions workflow (`.github/workflows/collect.yml`) runs at 17 minutes past every hour. Find a Tender and Contracts Finder are read every hour from where the last run left off. Public Contracts Scotland and Sell2Wales only offer whole months, so they are re-read every four hours. The first run goes back 14 days.

To run it by hand: Actions tab, "Collect UK tenders", "Run workflow". Tick force_all to read every source regardless of the four-hour spacing.

## Settings

`settings.json` controls which sources are on, how far back the first run looks, how many months of history to keep, and the day windows for each tab. Set `"enabled": false` on a source to switch it off.

## Direct link to the latest workbook

```
https://raw.githubusercontent.com/<owner>/uk-tenders/main/UK_Tenders.xlsx
```

Excel, Power Query and Power Automate can all read that address, which is how the file reaches Microsoft 365.

## Local test

```
python tests/test_offline.py
```

This checks the parsing against sample notices and writes `tests/sample.xlsx` without touching the network.
