# Business Leads Scraper — City + Category → Excel

Generate B2B lead lists (business name, phone, website, email, address, coordinates) for **any city in the world** and **any business category**, exported as clean CSV + formatted Excel.

```bash
pip install -r requirements.txt

python leads.py --city "London" --category dentist
python leads.py --city "Lahore" --category restaurant --limit 300
python leads.py --city "Berlin" --tag "shop=bakery"     # any OpenStreetMap tag
```

## Sample output

See [`sample_output/`](sample_output/) — every lead in a formatted, filterable Excel sheet.

```
===== SUMMARY =====
Total leads   : 412
With phone    : 268
With website  : 301
```

## Data source & legality

This tool uses the **OpenStreetMap Overpass API** — free, public, open-licensed data (ODbL). That means:

- 100% legal to collect and hand to clients
- No proxies, no CAPTCHAs, no account bans
- Global coverage, 12+ built-in categories, any custom tag supported

> **Note for client projects:** when clients need Google Maps data specifically, the professional route is the official **Google Places API** (paid, reliable, ToS-compliant). Scraping Google Maps directly violates Google's Terms of Service and risks blocks mid-project — I build lead pipelines on data sources that won't collapse. Data completeness on OSM varies by region: bigger international cities return the richest contact data.

## Tech stack

`requests` · Overpass QL · `openpyxl` · dedupe + cleaning built in

## Possible extensions (available on request)

- Email finder / verification layer
- Push leads straight into Google Sheets, Airtable or a CRM (HubSpot, Notion)
- Scheduled weekly lead refresh with only-new-leads detection
- Multi-city batch runs
