# UK tenders collector

Collects every notice published through the UK public procurement feeds that offer open data, and keeps them in one Excel workbook.

Sources: Find a Tender, Contracts Finder, Public Contracts Scotland and Sell2Wales. All of it is Open Government Licence data. eTendersNI has no public feed; above-threshold Northern Ireland notices still arrive through Find a Tender.

## What you get

`UK_Tenders.xlsx` is rewritten after every run and has these tabs:

Open tenders: live opportunities, newest first, one row per procurement (sort the Deadline column if you want soonest first).
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
https://raw.githubusercontent.com/ALEXWHELANTPC/uk-tenders/main/UK_Tenders.xlsx
```

A Power Automate flow called "UK tenders workbook to OneDrive" (owner: Alex Whelan) fetches that address every hour with the OneDrive "Upload file from URL" action and saves it to OneDrive at My files > UK Tenders > UK_Tenders.xlsx, overwriting the previous copy. Nothing else is needed on the Microsoft side.

## Things learned from the live feeds

Find a Tender's API is slow and sometimes answers 502 or 504. A page of 100 notices can take a minute or more, so the collector reads it one day at a time, records how far it got, and carries on next run. The first 14-day backfill therefore takes a few hourly runs to complete.

Public Contracts Scotland (and Sell2Wales, which runs on the same platform) send an incomplete certificate chain. The collector fetches the missing intermediate certificate from the address inside the server's certificate and verifies against that, so certificate checking stays on.

Sell2Wales's API returned "Error converting data type nvarchar to float" (HTTP 500) for every query on 23 November 2026, a fault on their side. The collector gives up on a source after three failed requests in a row and tries again next run, so when Sell2Wales recovers it will start filling in by itself. Welsh notices above the UK threshold also arrive through Find a Tender in the meantime.

## Local test

```
python tests/test_offline.py
```

This checks the parsing against sample notices and writes `tests/sample.xlsx` without touching the network.
