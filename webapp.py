# ============================================================
# 🔄 ACOOLE PORTAL — PROFESSIONAL VERSION v4.30
#    • Robust pandas NaT/date normalisation for employee HR reports
#    • Google Drive backup: automatic hourly sync + immediate Super Admin sync
#    • NEW: Employee Items (Issue / Holdings / Leaver Check-in → auto Store Deduction)
# ============================================================
import streamlit as st
import os
import sys
import json
import base64
import shutil
import subprocess
import pandas as pd
import io
import re
import requests
import threading
import queue
import time
import textwrap

# ============================================================
# Employee Weekly Timesheet module embedded in this single file.
# Uses a separate workbook and does not write to payroll.
# ============================================================
import json
import os
import uuid
from datetime import date, datetime, time, timedelta
from typing import Callable, Iterable

import pandas as pd
import streamlit as st

COLUMNS = [
    "Timesheet_ID", "Week_Start", "Week_End", "Department", "Employee_ID", "Employee_Name",
    "Job_Title", "Days_JSON", "Worked_Hours", "Contracted_Hours", "Travel_Hours", "Away_Hours",
    "Net_Balance_Hours", "Status", "Submitted_By", "Submitted_At", "Approved_By", "Approved_At",
    "Director_Comments", "Flagged_Days_JSON", "Updated_At"
]


def _ensure_workbook(path: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        pd.DataFrame(columns=COLUMNS).to_excel(path, index=False, engine="openpyxl")
    else:
        try:
            df = pd.read_excel(path, engine="openpyxl")
            changed = False
            for col in COLUMNS:
                if col not in df.columns:
                    df[col] = ""
                    changed = True
            if changed:
                tmp = path + ".repair.tmp.xlsx"
                df.to_excel(tmp, index=False, engine="openpyxl")
                os.replace(tmp, path)
        except Exception as exc:
            raise RuntimeError(f"Cannot read timesheet workbook safely: {exc}") from exc


def _load(path: str) -> pd.DataFrame:
    _ensure_workbook(path)
    return pd.read_excel(path, engine="openpyxl").fillna("")


def _save(path: str, df: pd.DataFrame, sync_callback=None) -> None:
    tmp = path + ".write.tmp.xlsx"
    df.reindex(columns=COLUMNS).to_excel(tmp, index=False, engine="openpyxl")
    os.replace(tmp, path)
    if callable(sync_callback):
        try:
            sync_callback(path)
        except Exception as exc:
            # Local workbook remains saved; the host app's sync monitor can retry.
            print(f"Timesheet Google Drive sync queued/failed: {exc}")


def _week_days(week_start: date):
    return [week_start + timedelta(days=i) for i in range(7)]


def _float(value, default=0.0):
    try:
        x = float(value)
        return x if x >= 0 else default
    except (TypeError, ValueError):
        return default


def _day_hours(day: dict) -> float:
    if day.get("away_type") in {"Holiday", "Sick", "Other absence"}:
        return 0.0
    start, end = str(day.get("start", "")).strip(), str(day.get("end", "")).strip()
