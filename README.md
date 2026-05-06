# PHL5 S1 OB Hourly Tracker

Real-time shift dashboard for PHL5 Outbound — Shift 1 (Mon–Thu, 7:30 AM–6:00 PM).

## What it tracks

| Metric | Source | Refresh |
|---|---|---|
| LP Volume & LP UPH | 28-Day LP SharePoint (FY26 LP - PHL5.xlsm) | Daily 7:30 AM |
| Total Backlog | Endgame pt-status (FC 3124) | Hourly at :15 |
| Actual Volume / Hour | DRAX · Stationary Picking (019209516) | Hourly at :15 |
| ASRS UPH / Hour | DRAX · Stationary Picking (019209516) | Hourly at :15 |
| Box Finish UPH / Hour | DRAX · Box Finishing (019034514) | Hourly at :15 |
| Actual Bagging / Hour | DRAX · Bagging - Manual (019034295) | Hourly at :15 |
| Bagging Goal | 10% of Goal Volume (auto-derived) | On load |

## How it works

```
data.js  ←──── fetch_lp_data.py (7:30 AM daily, Windows Task Scheduler)
         ←──── qa-kitten scheduler (Endgame + DRAX, every :15 Mon–Thu)
         │
index.html ──── reads data.js every 15 min via dynamic script tag (no page reload)
```

## Setup

1. Open `index.html` in your browser
2. Run `uv run fetch_lp_data.py` once to authenticate MS Graph (SharePoint)
3. The qa-kitten scheduler handles Endgame + DRAX automatically on Walmart network

## Files

| File | Purpose |
|---|---|
| `index.html` | Dashboard UI (Alpine.js + Tailwind CDN) |
| `data.js` | Auto-generated data payload (read by dashboard) |
| `fetch_lp_data.py` | Pulls LP Volume + UPH from SharePoint via MS Graph |
| `fetch_backlog.py` | Pulls Endgame backlog via API (fallback script) |
| `uv.toml` | Points `uv` at Walmart internal PyPI mirror |
