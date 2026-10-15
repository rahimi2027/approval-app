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
from datetime import datetime, date, timezone, timedelta

from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2 import service_account



from datetime import datetime, date, timedelta
import pandas as pd
import streamlit as st
import os

# ============================================================
# CONSTANTS
# ============================================================

PAYROLL_MONTHS = [
    ("2026-09", "September 2026"), ("2026-10", "October 2026"),
    ("2026-11", "November 2026"),  ("2026-12", "December 2026"),
    ("2027-01", "January 2027"),   ("2027-02", "February 2027"),
    ("2027-03", "March 2027"),     ("2027-04", "April 2027"),
    ("2027-05", "May 2027"),       ("2027-06", "June 2027"),
    ("2027-07", "July 2027"),      ("2027-08", "August 2027"),
    ("2027-09", "September 2027"), ("2027-10", "October 2027"),
    ("2027-11", "November 2027"),  ("2027-12", "December 2027"),
]

PAYROLL_STATUS_OPTIONS = [
    "Worked", "Late In", "Late Out", "Authorised Absence", "HD Holiday",
    "Overtime Approved", "College", "Holiday", "Sick", "Maternity",
    "Lay-off", "Unpaid Holiday", "Unpaid Absence", "Bank Holiday",
    "AWOL", "N/A", "Leaver",
]

PAYROLL_TIME_ENTRY_STATUSES = {
    "Worked", "Late In", "Late Out", "Authorised Absence",
    "HD Holiday", "Overtime Approved", "College",
}

PAYROLL_TIMESHEETS_COLUMNS = [
    "Week_Key", "Department", "Employee_ID", "Day_Index",
    "Time_In", "Time_Out", "Travel", "Time_Away", "Status", "Note",
]
PAYROLL_WEEK_STATES_COLUMNS = [
    "Week_Key", "Department", "Status",
    "Submitted_At", "Submitted_By", "Approved_At", "Approved_By",
]
PAYROLL_FLAGS_COLUMNS = [
    "Week_Key", "Department", "Employee_ID", "Day_Index",
    "Comment", "Flagged_By", "Flagged_At",
]
PAYROLL_APPROVALS_COLUMNS = [
    "Week_Key", "Department", "Employee_ID", "Approved_At", "Approved_By",
]

# ============================================================
# INJECTED CONTEXT (filled by init_payroll)
# ============================================================

_ctx = {
    "app_folder": None,
    "safe_init_excel": None,
    "read_excel": None,
    "sync_to_drive": None,
    "employees_provider": None,
    "invalidate_cache": None,
    "initialised": False,
}


def init_payroll(app_folder, safe_init_excel_fn, read_excel_records_fn,
                 sync_to_drive_fn, employees_provider,
                 invalidate_cache_fn=None):
    """Call once from the main app after infrastructure is available."""
