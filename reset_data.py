#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = []
# ///
"""
PHL5 Daily Reset -- runs at 7:00 AM Mon-Thu via Task Scheduler.

Zeros out all daily metrics in data.js so the dashboard starts clean
when the 7:30 AM shift opens. Static config (cellVol, cellUPH, capacity
pool labels/colors) is preserved from the previous day's data.
Data refills naturally from 8:15 AM onward via fetch_hourly.py + qa-kitten.
"""
import json
import sys
from datetime import datetime
from pathlib import Path

HERE    = Path(__file__).parent
DATA_JS = HERE / "data.js"

N_SHIFT_SLOTS = 11   # 7:30 AM ... 5:00 PM (11 hourly buckets)
N_OTS_SLOTS   = 20   # 20 cut-time slots in otsData


def _read_current() -> dict:
    """Parse window.PHL5_DATA from data.js. Returns {} on any failure."""
    if not DATA_JS.exists():
        return {}
    try:
        txt   = DATA_JS.read_text(encoding="utf-8")
        start = txt.find("{")
        end   = txt.rfind("}") + 1
        return json.loads(txt[start:end]) if start != -1 else {}
    except Exception as e:
        print(f"  Warning: could not parse existing data.js -- {e}")
        return {}


def _capacity_skeleton(existing: list[dict]) -> list[dict]:
    """Preserve pool type/id/color; zero out planned/consumed/available."""
    default = [
        {"type": "1P",     "pool": 1,  "color": "#0053e2"},
        {"type": "3P/WFS", "pool": 4,  "color": "#ffc220"},
        {"type": "SAMS",   "pool": 27, "color": "#2a8703"},
    ]
    source = existing if len(existing) == 3 else default
    return [
        {
            "type":      row.get("type",  default[i]["type"]),
            "pool":      row.get("pool",  default[i]["pool"]),
            "planned":   0,
            "consumed":  0,
            "available": 0,
            "color":     row.get("color", default[i]["color"]),
        }
        for i, row in enumerate(source)
    ]


def main() -> None:
    print(f"\nPHL5 Daily Reset -- {datetime.now().strftime('%H:%M:%S  %a %b %d %Y')}\n")

    current = _read_current()

    fresh: dict = {
        "lpVolume":     0,
        "lpUPH":        0.0,
        "backlog":      0,
        "blNotShipped": 0,
        "blLoaded":     0,
        "blDiverted":   0,
        "blShipLabel":  0,
        "actualVol":    [0]   * N_SHIFT_SLOTS,
        "asrsUPH":      [0.0] * N_SHIFT_SLOTS,
        "boxUPH":       [0.0] * N_SHIFT_SLOTS,
        "actualBag":    [0]   * N_SHIFT_SLOTS,
        "hoursWorked":  [0.0] * N_SHIFT_SLOTS,
        "capacity":     _capacity_skeleton(current.get("capacity", [])),
        "otsData":      [0]   * N_OTS_SLOTS,
        # Preserve static identifiers -- these don't change day to day
        "cellVol":      current.get("cellVol",  ""),
        "cellUPH":      current.get("cellUPH",  ""),
        "lastUpdated":  datetime.now().isoformat(timespec="seconds"),
        "skipped":      False,
    }

    ts  = fresh["lastUpdated"]
    out = (
        "// Auto-generated -- daily reset at 7:00 AM, data refills from 8:15 AM\n"
        f"// Reset: {ts}\n"
        f"window.PHL5_DATA = {json.dumps(fresh, indent=2)};\n"
    )

    DATA_JS.write_text(out, encoding="utf-8")
    print(f"  [OK] data.js zeroed -- {ts}")
    print("  [i]  First real data pull at 8:15 AM via fetch_hourly.py + qa-kitten\n")


if __name__ == "__main__":
    main()
