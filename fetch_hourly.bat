@echo off
:: PHL5 Hourly Fetch — Backlog + DRAX + Gscope
:: Runs at :15 of each hour, Mon–Thu 8:15 AM–5:15 PM via Task Scheduler
cd /d "%~dp0"
uv run --index-url https://pypi.ci.artifacts.walmart.com/artifactory/api/pypi/external-pypi/simple --allow-insecure-host pypi.ci.artifacts.walmart.com fetch_hourly.py >> fetch_hourly.log 2>&1
