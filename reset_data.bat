@echo off
:: PHL5 Daily Reset — zeros data.js at 7:00 AM Mon–Thu via Task Scheduler
:: Dashboard clears before the 7:30 AM shift opens
cd /d "%~dp0"
uv run --index-url https://pypi.ci.artifacts.walmart.com/artifactory/api/pypi/external-pypi/simple --allow-insecure-host pypi.ci.artifacts.walmart.com reset_data.py >> reset_data.log 2>&1
