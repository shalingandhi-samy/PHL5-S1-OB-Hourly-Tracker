@echo off
title PHL5 Cookie Extractor
echo.
echo  =======================================================
echo   PHL5 Cookie Extractor -- Run this each morning
echo   Grabs Gscope token. DRAX is handled by qa-kitten.
echo   Edge will open. SSO fires automatically on Walmart VPN.
echo   Gscope may ask you to log in -- do it in that window.
echo  =======================================================
echo.
cd /d "%~dp0"
uv run --index-url https://pypi.ci.artifacts.walmart.com/artifactory/api/pypi/external-pypi/simple --allow-insecure-host pypi.ci.artifacts.walmart.com extract_cookies.py
echo.
if %ERRORLEVEL% EQU 0 (
    echo  Done! Run smoke_test.py to verify.
) else (
    echo  Something went wrong -- check output above.
)
echo.
pause
