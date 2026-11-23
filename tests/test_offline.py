"""Offline check: normalise sample releases and build a workbook without the network."""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import collector  # noqa: E402
from workbook import build_workbook  # noqa: E402

SETTINGS = json.loads((Path(__file__).resolve().parents[1] / "settings.json").read_text())

FTS_TENDER = {
    "ocid": "ocds-h6vhtk-07c104", "id": "110456-2026", "tag": ["tender"], "date": "2026-11-23T10:53:13Z",
    "parties": [{"name": "RSSB", "id": "GB-PPON-X", "address": {"locality": "London", "region": "UKI31"}, "roles": ["buyer"]}],
    "buyer": {"id": "GB-PPON-X", "name": "RSSB"},
    "tender": {
        "title": "T1431 - Defective On-Train Equipment research", "description": "RSSB is seeking a supplier.", "status": "active",
        "items": [{"id": "1", "additionalClassifications": [{"scheme": "CPV", "id": "73110000", "description": "Research services"}], "relatedLot": "1"}],
        "value": {"amountGross": 150000.0, "amount": 125000.0, "currency": "GBP"},
        "procurementMethod": "open", "procurementMethodDetails": "Below threshold - open competition",
        "tenderPeriod": {"endDate": "2026-12-17T12:00:00+00:00"},
        "lots": [{"id": "1", "suitability": {"sme": True}, "contractPeriod": {"startDate": "2027-02-01T00:00:00+00:00", "endDate": "2027-07-31T23:59:59+01:00"}}],
        "documents": [{"id": "110456-2026", "documentType": "tenderNotice", "noticeType": "UK4", "url": "https://www.find-tender.service.gov.uk/Notice/110456-2026"}],
    },
}

FTS_AWARD = {
    "ocid": "ocds-h6vhtk-07c103", "id": "110455-2026", "date": "2026-11-23T10:44:36Z", "tag": ["award", "contract"],
    "tender": {"title": "SBU-DCO-102-SP-26", "status": "complete", "classification": {"scheme": "CPV", "id": "72220000", "description": "Systems and technical consultancy services"},
               "description": "Fixed-price strategic consultancy.", "items": [{"id": "1", "deliveryAddresses": [{"region": "UKL18"}]}],
               "procurementMethod": "limited", "procurementMethodDetails": "Award procedure without prior publication"},
    "awards": [{"id": "110455-2026-1", "status": "active", "suppliers": [{"id": "GB-FTS-188521", "name": "CGI UK Ltd"}]}],
    "parties": [{"name": "Swansea Bay University Health Board", "id": "GB-FTS-188520", "address": {"locality": "Cardiff", "region": "UK"}, "roles": ["buyer"]},
                {"name": "CGI UK Ltd", "id": "GB-FTS-188521", "roles": ["supplier"]}],
    "buyer": {"id": "GB-FTS-188520", "name": "Swansea Bay University Health Board"},
    "contracts": [{"id": "110455-2026-1", "awardID": "110455-2026-1", "status": "active", "value": {"amount": 148770, "currency": "GBP"}, "dateSigned": "2026-11-23T00:00:00Z"}],
}

CF_AWARD = {
    "ocid": "ocds-b5fd17-fbc65894", "id": "8c80d531-6cfb-4242-b450-43e69be693e1-922277", "date": "2026-11-23T10:16:28Z", "tag": ["award"],
    "tender": {"id": "CF-0093600D24000000HFzAEAW", "title": "WSCC Route 7412 Burstow Park", "description": "WSCC Route 7412", "status": "complete",
               "classification": {"scheme": "CPV", "id": "60120000", "description": "Taxi services"},
               "items": [{"id": "1", "deliveryAddresses": [{"postalCode": "PO19 1RQ"}, {"region": "South East", "countryName": "United Kingdom"}]}],
               "value": {"amount": 59941.2, "currency": "GBP"}, "procurementMethod": "selective", "procurementMethodDetails": "Call-off from a dynamic purchasing system",
               "tenderPeriod": {"endDate": "2026-09-07T12:00:00+01:00"}, "contractPeriod": {"startDate": "2026-09-21T00:00:00+01:00", "endDate": "2027-07-31T23:59:59+01:00"},
               "suitability": {"sme": True, "vcse": False}},
    "parties": [{"id": "GB-CFS-244021", "name": "West Sussex County Council", "roles": ["buyer"]}, {"id": "GB-CFS-331208", "name": "EASY ACCESSIBLE CARS LTD", "roles": ["supplier"]}],
    "buyer": {"id": "GB-CFS-244021", "name": "West Sussex County Council"},
    "awards": [{"id": "x-1", "status": "active", "value": {"amount": 59941.2, "currency": "GBP"}, "suppliers": [{"id": "GB-CFS-331208", "name": "EASY ACCESSIBLE CARS LTD"}],
                "documents": [{"id": "1", "documentType": "awardNotice", "url": "https://www.contractsfinder.service.gov.uk/Notice/8c80d531-6cfb-4242-b450-43e69be693e1"}]}],
}

PCS_TENDER = {
    "id": "rls-1-NOV569336", "ocid": "ocds-r6ebe6-0000849139", "date": "2026-11-23T00:00:00Z", "tag": ["tender"],
    "buyer": {"name": "Scottish Environment Protection Agency", "id": "org-1"},
    "parties": [{"id": "org-1", "name": "Scottish Environment Protection Agency", "roles": ["buyer"], "address": {"region": "UKM"}}],
    "tender": {"title": "Consultancy Support for Waste Monitoring Data 2027-2031", "description": "Support.", "status": "active",
               "items": [{"id": "1", "additionalClassifications": [{"id": "90700000", "scheme": "CPV"}], "deliveryAddresses": [{"region": "UKM63"}]}],
               "classification": {"id": "90700000", "scheme": "CPV"}, "value": {"amount": 250000, "currency": "GBP"},
               "tenderPeriod": {"endDate": "2026-12-18T12:00:00Z"}, "procurementMethod": "open",
               "documents": [{"id": "NOV569336", "documentType": "contractNotice", "url": "https://www.publiccontractsscotland.gov.uk/search/show/search_view.aspx?ID=NOV569336"},
                             {"id": "economic", "documentType": "economicSelectionCriteria"}]},
    "links": [{"rel": "canonical", "href": "https://api.publiccontractsscotland.gov.uk/v1/Notice?id=ocds-r6ebe6-0000849139"}],
}


def main() -> None:
    rows = [
        collector.normalise(FTS_TENDER, "find_a_tender", SETTINGS),
        collector.normalise(FTS_AWARD, "find_a_tender", SETTINGS),
        collector.normalise(CF_AWARD, "contracts_finder", SETTINGS),
        collector.normalise(PCS_TENDER, "public_contracts_scotland", SETTINGS),
    ]
    for r in rows:
        assert r, "normalise returned nothing"
        print({k: r[k] for k in ("source", "notice_kind", "notice_form", "title", "buyer", "region", "value_gbp", "cpv_codes", "deadline", "contract_end", "suppliers", "link", "notice_id")})
    assert rows[0]["notice_kind"] == "Tender" and rows[0]["region"] == "London" and rows[0]["value_gbp"] == 150000.0
    assert rows[1]["notice_kind"] == "Award" and rows[1]["suppliers"] == "CGI UK Ltd" and rows[1]["region"] == "Wales" and rows[1]["value_gbp"] == 148770
    assert rows[2]["region"] == "South East" and rows[2]["link"].endswith("/Notice/8c80d531-6cfb-4242-b450-43e69be693e1")
    assert rows[3]["region"] == "Scotland" and "NOV569336" in rows[3]["link"] and rows[3]["notice_form"] == "contractNotice"

    out = Path(__file__).resolve().parent / "sample.xlsx"
    state = {"last_totals": {"Find a Tender": {"releases": 2, "new": 2, "updated": 0, "errors": 0}}, "sources": {"find_a_tender": {"last_success": "2026-11-23T11:00:00"}}}
    build_workbook(rows, out, SETTINGS, state)
    print("workbook ok:", out, out.stat().st_size, "bytes")


if __name__ == "__main__":
    main()
