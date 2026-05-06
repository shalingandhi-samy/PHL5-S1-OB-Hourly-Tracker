@echo off
cd /d "%~dp0"
uv run fetch_backlog.py >> fetch_backlog.log 2>&1
