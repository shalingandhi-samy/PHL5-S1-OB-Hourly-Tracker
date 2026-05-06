@echo off
:: PHL5 LP Data Fetcher — runs at 7:30 AM Mon–Thu via Task Scheduler
cd /d "%~dp0"
uv run --index-url https://pypi.ci.artifacts.walmart.com/artifactory/api/pypi/external-pypi/simple --allow-insecure-host pypi.ci.artifacts.walmart.com fetch_lp_data.py >> fetch_lp_data.log 2>&1
