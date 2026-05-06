@echo off
:: PHL5 Auto-Screenshot -- sends dashboard JPEG to Teams at :20 each hour
:: Runs Mon-Thu 8:20 AM - 5:20 PM via Task Scheduler (PHL5-Screenshot-Teams)
chcp 65001 >nul 2>&1
cd /d "%~dp0"
echo PHL5 Screenshot to Teams -- %TIME% %DATE% >> screenshot_teams.log
"C:\Users\S0G0K3S\.code-puppy-venv\Scripts\uv.exe" run ^
    --index-url https://pypi.ci.artifacts.walmart.com/artifactory/api/pypi/external-pypi/simple ^
    --allow-insecure-host pypi.ci.artifacts.walmart.com ^
    screenshot_teams.py >> screenshot_teams.log 2>&1
