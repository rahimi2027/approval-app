# ============================================================
# 🔄 ACOOLE PORTAL — PROFESSIONAL VERSION v4.31
#    • Fixed: page no longer blocks on Google Drive at startup
#    • Fixed: no background worker thread (removed)
#    • Google Drive is now FULLY LAZY - only connects on user action
# ============================================================
import streamlit as st

# st.set_page_config MUST be the first Streamlit command.
st.set_page_config(page_title="ACoole Portal", layout="wide", initial_sidebar_state="expanded")

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

st.markdown("""
    <style>
    .block-container {
        padding-top: 2rem !important;
        padding-left: 0.3rem !important;
        padding-right: 2rem !important;
        max-width: 1400px !important;
        width: 90% !important;
    }
    section[data-testid="stSidebar"] { width: 320px !important; }
    section[data-testid="stSidebar"] > div:first-child > div {
        padding-left: 0.2rem !important;
        padding-right: 0.2rem !important;
    }
    section[data-testid="stSidebar"] > div:first-child { align-items: flex-start !important; }
    .streamlit-expander { width: 100% !important; margin: 0 !important; padding-left: 0 !important; }
    .streamlit-expanderHeader { justify-content: flex-start !important; padding-left: 0.5rem !important; }
    .streamlit-expanderContent {
        width: 100% !important; box-sizing: border-box !important;
        padding: 0.5rem !important; padding-left: 0.3rem !important; text-align: left !important;
    }
    .streamlit-expanderContent form { width: 100% !important; box-sizing: border-box !important; padding: 0 !important; margin: 0 !important; }
    .streamlit-expanderContent label { text-align: left !important; justify-content: flex-start !important; padding-left: 0.2rem !important; }
    .streamlit-expanderContent input { width: 100% !important; box-sizing: border-box !important; }
    .streamlit-expanderContent .stButton > button { width: 100% !important; box-sizing: border-box !important; margin-top: 0.5rem !important; }
    </style>
""", unsafe_allow_html=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
APP_FOLDER = os.path.join(BASE_DIR, "Acoole_App_Uploads")
os.makedirs(APP_FOLDER, exist_ok=True)
UPLOAD_DIR = os.path.join(APP_FOLDER, "uploaded_attachments")
AUDIT_LOG_PATH = os.path.join(APP_FOLDER, "audit_log.xlsx")
AUDIT_LOG_FILE = AUDIT_LOG_PATH
AUDIT_COLUMNS = ["AuditID", "Timestamp", "User_Name", "User_Role", "Action", "Request_ID", "Department", "Amount", "Decision_By", "Decision_Date", "Field_Changed", "Old_Value", "New_Value", "IP_Address"]
ALLOWED_CLEAR_ROLES = ["Super Admin"]
ARCHIVE_FOLDER = os.path.join(APP_FOLDER, "audit_archives/")
PDF_DIR = os.path.join(APP_FOLDER, "approved_pdfs")
LOGO_PATH = os.path.join(BASE_DIR, "logo.png")
APPROVED_STAMP_PATH = os.path.join(BASE_DIR, "approved_stamp.png")
REJECTED_STAMP_PATH = os.path.join(BASE_DIR, "rejected_stamp.png")
EXCEL_PATH = os.path.join(APP_FOLDER, "requests.xlsx")
USER_DB_PATH = os.path.join(APP_FOLDER, "user_database.xlsx")
SETTINGS_PATH = os.path.join(APP_FOLDER, "settings.xlsx")
WORK_ORDERS_PATH = os.path.join(APP_FOLDER, "work_orders.xlsx")
WORK_ORDER_PDF_DIR = os.path.join(APP_FOLDER, "work_order_pdfs")
INSPECTOR_BONUS_PATH = os.path.join(APP_FOLDER, "inspector_bonus.xlsx")
INSPECTOR_BONUS_PDF_DIR = os.path.join(APP_FOLDER, "inspector_bonus_pdfs")
GOOGLE_DRIVE_FOLDER_ID = "1g3DsqT_w_tU0QBnrXcZqYjp51SokH4hG"

# ============================================================
# 👥 HR LEAVE SETTLEMENT — PATHS & CONSTANTS
# ============================================================
HR_LEAVE_PATH = os.path.join(APP_FOLDER, "hr_leave_requests.xlsx")
HR_DAILY_RATES_PATH = os.path.join(APP_FOLDER, "hr_daily_rates.xlsx")
HR_LEAVE_PDF_DIR = os.path.join(APP_FOLDER, "hr_leave_pdfs")
HR_EMPLOYEES_PATH = os.path.join(APP_FOLDER, "hr_employees.xlsx")
HR_PORTAL_LEAVE_PATH = os.path.join(APP_FOLDER, "hr_portal_leave_records.xlsx")
os.makedirs(HR_LEAVE_PDF_DIR, exist_ok=True)

HR_EMPLOYEE_COLUMNS = [
    "Employee ID", "Full Name", "Start Date", "Position / Job Title", "Department", "Work Type / Sub-department",
    "Agreement Type", "Status", "Working Pattern", "Days Worked Per Week", "Holiday Entitlement Override", "Entitlement Adjustment Note", "Leaving Date", "Leaving Reason"
]
HR_PORTAL_LEAVE_COLUMNS = [
    "Leave ID", "Employee ID", "Date From", "Date To", "Leave Type", "Days",
    "Status", "Request Source", "Request Reference", "Notes", "Recorded By", "Recorded At",
    "Entry Source", "Requested By", "Requested At", "Approved By", "Approved At", "Rejection Reason"
]

HR_LEAVE_COLUMNS = [
    "ID", "Employee ID", "Employee Name", "Employee Department", "Transaction Type",
    "Category Reason", "Owe Owed", "Date", "Number of Days", "Amount (£)",
    "Line Manager", "Description", "Attachment Name", "Status",
    "Director Comments", "Rejection Reason", "Decision Date", "Decision By",
    "Submitted By", "Submitted Date", "PDF File Path", "Final Holiday Settlement"
]
HR_DAILY_RATES_COLUMNS = ["Department", "Transaction Type", "Daily Rate (£)", "Active"]
DEFAULT_HR_CATEGORIES = [
    "Annual Leave Balance", "Unused Holiday Payout",
    "Overused Holiday Deduction", "Leave Encashment",
]
DEFAULT_OWE_OWED = ["Company Owes Employee", "Employee Owes Company"]

# ============================================================
# 📦 STORE DEPARTMENT DEDUCTION — PATHS & CONSTANTS
# ============================================================
STORE_DEDUCTION_PATH = os.path.join(APP_FOLDER, "store_deduction_requests.xlsx")
STORE_ITEMS_PATH = os.path.join(APP_FOLDER, "store_items.xlsx")
STORE_DEDUCTION_PDF_DIR = os.path.join(APP_FOLDER, "store_deduction_pdfs")
os.makedirs(STORE_DEDUCTION_PDF_DIR, exist_ok=True)

STORE_DEDUCTION_COLUMNS = [
    "ID", "Employee Name", "Date of Leaving", "Employee Department",
    "Line Manager", "Date of Submit", "Type", "Items Deducted JSON", "Total Deduction (£)",
    "Description", "Attachment Name", "Status", "Director Comments",
    "Rejection Reason", "Decision Date", "Decision By", "Submitted By",
    "Submitted Date", "PDF File Path"
]
STORE_ITEMS_COLUMNS = ["Item Name", "Price (£)", "Active"]
DEFAULT_STORE_ITEMS = [
    {"Item Name": "Laptop - Dell Latitude", "Price (£)": 510.00, "Active": True},
    {"Item Name": "Mobile Phone - iPhone 13", "Price (£)": 450.00, "Active": True},
    {"Item Name": "ID Card / Access Badge", "Price (£)": 20.00, "Active": True},
    {"Item Name": "Safety Boots", "Price (£)": 75.00, "Active": True},
    {"Item Name": "Company Vehicle Keys", "Price (£)": 150.00, "Active": True},
]

# ============================================================
# 🧰 EMPLOYEE ITEM ISSUE / RETURN — PATHS & CONSTANTS
# ============================================================
EMPLOYEE_ITEMS_PATH = os.path.join(APP_FOLDER, "employee_items.xlsx")
ITEM_CHECKIN_PATH   = os.path.join(APP_FOLDER, "item_checkins.xlsx")

EMPLOYEE_ITEMS_COLUMNS = [
    "ID", "Employee ID", "Employee Name", "Department",
    "Item Name", "Quantity Issued", "Quantity Returned", "Quantity Outstanding",
    "Unit Price (£)", "Total Value (£)", "Issue Date", "Issued By",
    "Status", "Notes",
]
ITEM_CHECKIN_COLUMNS = [
    "ID", "Employee ID", "Employee Name", "Department",
    "Leaving Date", "Check-in Date", "Checked In By",
    "Returned Items JSON", "Not Returned Items JSON",
    "Total Deduction (£)", "Deduction Request ID", "Return Request ID", "Status", "Notes",
]

# ============================================================
# 🔗 EMPLOYEE LEAVER CLEARANCE — HR + STORE WORKFLOW
# ============================================================
LEAVER_CLEARANCE_PATH = os.path.join(APP_FOLDER, "employee_leaver_clearance.xlsx")
LEAVER_CLEARANCE_COLUMNS = [
    "Clearance ID", "Employee ID", "Employee Name", "Department", "Leaving Date",
    "Leaving Reason", "Created By", "Created At",
    "HR Status", "Holiday Balance (Days)", "Holiday Settlement ID", "Holiday Settlement Status",
    "Holiday Settlement Amount (£)", "Holiday Settlement Type",
    "Store Status", "Store Check-in ID", "Store Deduction ID", "Store Return ID",
    "Outstanding Item Value (£)", "Store Completed By", "Store Completed At",
    "Payroll Status", "Payroll Request IDs", "Payroll Total Addition (£)", "Payroll Total Deduction (£)",
    "Payroll Completed By", "Payroll Completed At",
    "Final Status", "Final Cleared By", "Final Cleared At", "Notes",
]

USER_DB_COLUMNS = [
    "full_name", "username", "password", "role", "dept",
    "can_view_all_dept", "can_generate_pdf", "can_download_data",
    "can_approve_requests", "can_access_inspector_bonus",
    "can_access_addition_deduction", "can_access_work_orders",
    "can_access_wo_total", "can_access_hr_leave", "can_access_leave_request", "can_access_employee_hr_reports",
    "can_access_holiday_calendar", "can_access_hr_reports", "can_access_employee_overview",
    "can_access_holiday_calculator", "can_access_employee_directory",
    "employee_id", "can_access_store_deduction", "can_access_employee_items",
    "is_active"
]

ONEDRIVE_CLIENT_ID = ""
ONEDRIVE_CLIENT_SECRET = ""
ONEDRIVE_TENANT_ID = "common"
ONEDRIVE_FOLDER = "Acoole_App_Uploads/"
USE_ONEDRIVE = False

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)
os.makedirs(ARCHIVE_FOLDER, exist_ok=True)
os.makedirs(WORK_ORDER_PDF_DIR, exist_ok=True)
os.makedirs(INSPECTOR_BONUS_PDF_DIR, exist_ok=True)

SCOPES = ["https://www.googleapis.com/auth/drive"]
drive_service = None
DRIVE_CONNECTION_ERROR = ""
DRIVE_LAST_SYNC = {}
DRIVE_LAST_SYNC_ERROR = {}
_DRIVE_SYNC_FINGERPRINTS = {}
_DRIVE_SYNC_LOCK = threading.RLock()
DRIVE_PENDING_SYNC = set()
_DRIVE_FIND_CACHE = {}
_DRIVE_FIND_CACHE_TTL = 60.0
_DRIVE_ID_CACHE = {}

# Local persistence / recovery. Google Drive is a backup layer; these local files
# are the durable working copy and RAM is never treated as the only copy.
LOCAL_RECOVERY_DIR = os.path.join(APP_FOLDER, "local_recovery")
DRIVE_SYNC_STATE_PATH = os.path.join(APP_FOLDER, "drive_sync_state.json")
_DRIVE_STATE_LOCK = threading.RLock()
os.makedirs(LOCAL_RECOVERY_DIR, exist_ok=True)

# ============================================================
# GOOGLE DRIVE CONNECTION — SERVICE ACCOUNT (LAZY / NON-BLOCKING)
# ============================================================
# IMPORTANT: Never build the Google Drive client or call the Drive API while this
# module is importing. Credentials are decoded locally (fast); the API client is
# created only when a Drive operation is explicitly requested by the user.
_DRIVE_CREDS_DICT = None
try:
    gdrive = st.secrets["gdrive"]
    b64_string = str(gdrive["key_b64"]).replace("\n", "").replace("\r", "").replace(" ", "").replace("\t", "")
    padding_needed = (4 - len(b64_string) % 4) % 4
    if padding_needed:
        b64_string += "=" * padding_needed
    _DRIVE_CREDS_DICT = json.loads(base64.b64decode(b64_string).decode("utf-8"))
    DRIVE_CONNECTION_ERROR = ""
except Exception as e:
    _DRIVE_CREDS_DICT = None
    DRIVE_CONNECTION_ERROR = f"{type(e).__name__}: {e}"
    print(f"Google Drive credentials unavailable; local storage will be used: {DRIVE_CONNECTION_ERROR}")


def _ensure_drive_service():
    """Create the Google Drive client only when backup I/O actually needs it."""
    global drive_service, DRIVE_CONNECTION_ERROR
    if drive_service is not None:
        return True
    if not _DRIVE_CREDS_DICT:
        return False
    try:
        credentials = service_account.Credentials.from_service_account_info(
            _DRIVE_CREDS_DICT, scopes=SCOPES
        )
        drive_service = build("drive", "v3", credentials=credentials, cache_discovery=True)
        DRIVE_CONNECTION_ERROR = ""
        return True
    except Exception as e:
        drive_service = None
        DRIVE_CONNECTION_ERROR = f"{type(e).__name__}: {e}"
        print(f"Google Drive connection deferred/failed; local storage remains active: {DRIVE_CONNECTION_ERROR}")
        return False


def _drive_find_file(filename, parent_id=GOOGLE_DRIVE_FOLDER_ID):
    if drive_service is None:
        return None
    cache_key = f"{parent_id}::{filename}"
    cached = _DRIVE_FIND_CACHE.get(cache_key)
    if cached and (time.monotonic() - cached[0]) < _DRIVE_FIND_CACHE_TTL:
        return cached[1]
    try:
        safe_name = str(filename).replace("'", "\\'")
        q = (f"name = '{safe_name}' and '{parent_id}' in parents and trashed = false")
        result = drive_service.files().list(q=q, spaces="drive", fields="files(id,name,modifiedTime,parents)", orderBy="modifiedTime desc", pageSize=100, includeItemsFromAllDrives=True, supportsAllDrives=True).execute()
        files = result.get("files", [])
        found = files[0] if files else None
        _DRIVE_FIND_CACHE[cache_key] = (time.monotonic(), found)
        return found
    except Exception as e:
        print(f"Drive lookup failed for {filename}: {e}")
        return None


def _drive_upload_path(local_path, filename=None, parent_id=GOOGLE_DRIVE_FOLDER_ID):
    if drive_service is None or not os.path.exists(local_path):
        return None
    filename = filename or os.path.basename(local_path)
    cache_key = f"{parent_id}::{filename}"

    def _do(file_id=None):
        media = MediaFileUpload(local_path, resumable=False)
        if file_id:
            return drive_service.files().update(fileId=file_id, media_body=media, fields="id,name,parents", supportsAllDrives=True).execute()
        metadata = {"name": filename, "parents": [parent_id]}
        return drive_service.files().create(body=metadata, media_body=media, fields="id,name,parents", supportsAllDrives=True).execute()

    try:
        cached_id = _DRIVE_ID_CACHE.get(cache_key)
        if cached_id:
            try:
                return _do(cached_id).get("id")
            except Exception:
                _DRIVE_ID_CACHE.pop(cache_key, None)
        existing = _drive_find_file(filename, parent_id)
        if existing:
            _DRIVE_ID_CACHE[cache_key] = existing["id"]
            return _do(existing["id"]).get("id")
        created = _do(None)
        _DRIVE_ID_CACHE[cache_key] = created.get("id")
        return created.get("id")
    except Exception as e:
        DRIVE_LAST_SYNC_ERROR[filename] = f"{type(e).__name__}: {e}"
        print(f"Drive upload failed for {filename}: {e}")
        return None


def _drive_download_file(file_id, local_path):
    if drive_service is None:
        return False
    tmp_path = f"{local_path}.tmp"
    try:
        request = drive_service.files().get_media(fileId=file_id, supportsAllDrives=True)
        with open(tmp_path, "wb") as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
        os.replace(tmp_path, local_path)
        return True
    except Exception as e:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
        print(f"Drive download failed for {local_path}: {e}")
        return False


DRIVE_BACKUP_FILENAMES = {
    EXCEL_PATH: "BACKUP_requests.xlsx",
    USER_DB_PATH: "BACKUP_users.xlsx",
    SETTINGS_PATH: "BACKUP_settings.xlsx",
    AUDIT_LOG_PATH: "BACKUP_audit_log.xlsx",
    WORK_ORDERS_PATH: "BACKUP_work_orders.xlsx",
    INSPECTOR_BONUS_PATH: "BACKUP_inspector_bonus.xlsx",
    HR_LEAVE_PATH: "BACKUP_hr_leave_requests.xlsx",
    HR_DAILY_RATES_PATH: "BACKUP_hr_daily_rates.xlsx",
    STORE_DEDUCTION_PATH: "BACKUP_store_transactions.xlsx",
    STORE_ITEMS_PATH: "BACKUP_store_items.xlsx",
    HR_EMPLOYEES_PATH: "BACKUP_hr_employee_records.xlsx",
    HR_PORTAL_LEAVE_PATH: "BACKUP_hr_portal_leave_records.xlsx",
    EMPLOYEE_ITEMS_PATH: "BACKUP_employee_items.xlsx",
    ITEM_CHECKIN_PATH:   "BACKUP_item_checkins.xlsx",
    LEAVER_CLEARANCE_PATH: "BACKUP_employee_leaver_clearance.xlsx",
}


def _load_drive_sync_state():
    try:
        if os.path.exists(DRIVE_SYNC_STATE_PATH):
            with open(DRIVE_SYNC_STATE_PATH, "r", encoding="utf-8") as fh:
                data = json.load(fh)
                return data if isinstance(data, dict) else {}
    except Exception as e:
        print(f"Drive sync state could not be loaded: {e}")
    return {}


def _save_drive_sync_state():
    try:
        tmp = DRIVE_SYNC_STATE_PATH + ".tmp"
        state = {
            "last_sync": DRIVE_LAST_SYNC,
            "last_error": DRIVE_LAST_SYNC_ERROR,
            "pending": sorted(os.path.basename(p) for p in DRIVE_PENDING_SYNC),
            "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(state, fh, indent=2, default=str)
        os.replace(tmp, DRIVE_SYNC_STATE_PATH)
    except Exception as e:
        print(f"Drive sync state could not be saved: {e}")


def _restore_drive_sync_state():
    state = _load_drive_sync_state()
    for key, value in (state.get("last_sync") or {}).items():
        DRIVE_LAST_SYNC[key] = value
    for key, value in (state.get("last_error") or {}).items():
        DRIVE_LAST_SYNC_ERROR[key] = value


def _create_local_recovery_snapshot(local_path):
    """Keep a rolling set of local snapshots so a bad write can be recovered."""
    if not local_path or not os.path.exists(local_path):
        return None
    try:
        os.makedirs(LOCAL_RECOVERY_DIR, exist_ok=True)
        filename = os.path.basename(local_path)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        snapshot = os.path.join(LOCAL_RECOVERY_DIR, f"{filename}.{stamp}.bak")
        shutil.copy2(local_path, snapshot)

        matches = sorted(
            [
                os.path.join(LOCAL_RECOVERY_DIR, n)
                for n in os.listdir(LOCAL_RECOVERY_DIR)
                if n.startswith(filename + ".") and n.endswith(".bak")
            ],
            key=lambda x: os.path.getmtime(x),
            reverse=True,
        )
        for old in matches[5:]:
            try:
                os.remove(old)
            except Exception:
                pass
        return snapshot
    except Exception as e:
        print(f"Local recovery snapshot failed for {local_path}: {e}")
        return None


def _mark_drive_file_dirty(local_path, make_recovery_snapshot=True):
    """Queue a local save for the next manual Drive upload. Never blocks the user."""
    if not local_path or not os.path.exists(local_path):
        return False
    with _DRIVE_SYNC_LOCK:
        if make_recovery_snapshot:
            _create_local_recovery_snapshot(local_path)
        DRIVE_PENDING_SYNC.add(os.path.abspath(local_path))
        _save_drive_sync_state()
    return True


def _sync_saved_file_to_drive_now(local_path):
    """Actually upload one file. Called only from an explicit user action."""
    _ensure_drive_service()
    local_path = os.path.abspath(local_path)
    filename = os.path.basename(local_path)
    if drive_service is None or not os.path.exists(local_path):
        msg = "Google Drive service is not connected." if drive_service is None else "Local file does not exist."
        DRIVE_LAST_SYNC_ERROR[filename] = msg
        _save_drive_sync_state()
        return False

    with _DRIVE_SYNC_LOCK:
        live_ok = False
        backup_ok = False
        live_error = ""
        backup_name = DRIVE_BACKUP_FILENAMES.get(local_path)

        try:
            live_id = _drive_upload_path(local_path, filename)
            if live_id:
                live_ok = True
                DRIVE_LAST_SYNC[filename] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                DRIVE_LAST_SYNC_ERROR.pop(filename, None)
            else:
                live_error = DRIVE_LAST_SYNC_ERROR.get(filename, "Live Google Drive upload returned no file ID.")
        except Exception as e:
            live_error = f"{type(e).__name__}: {e}"
            DRIVE_LAST_SYNC_ERROR[filename] = live_error

        if backup_name:
            try:
                backup_id = _drive_upload_path(local_path, backup_name)
                if backup_id:
                    backup_ok = True
                    synced_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    DRIVE_LAST_SYNC[backup_name] = synced_at
                    DRIVE_LAST_SYNC_ERROR.pop(backup_name, None)
                else:
                    DRIVE_LAST_SYNC_ERROR[backup_name] = DRIVE_LAST_SYNC_ERROR.get(backup_name, "Backup Google Drive upload returned no file ID.")
            except Exception as e:
                DRIVE_LAST_SYNC_ERROR[backup_name] = f"{type(e).__name__}: {e}"

        if live_error and not live_ok:
            print(f"Google Drive live workbook sync unavailable for {filename}: {live_error}")
        if backup_name and not backup_ok:
            print(f"Google Drive backup sync failed for {backup_name}: {DRIVE_LAST_SYNC_ERROR.get(backup_name, '')}")

        if live_ok or backup_ok:
            try:
                st_info = os.stat(local_path)
                _DRIVE_SYNC_FINGERPRINTS[local_path] = f"{st_info.st_size}:{int(st_info.st_mtime_ns)}"
            except Exception:
                pass
            DRIVE_PENDING_SYNC.discard(local_path)
            _save_drive_sync_state()
            return True

        DRIVE_PENDING_SYNC.add(local_path)
        _DRIVE_SYNC_FINGERPRINTS.pop(local_path, None)
        _save_drive_sync_state()
        return False


def sync_saved_file_to_drive(local_path):
    """Save hook. Super Admin writes sync immediately; other users queue for next manual sync."""
    filename = os.path.basename(local_path)
    if not os.path.exists(local_path):
        DRIVE_LAST_SYNC_ERROR[filename] = "Local file does not exist."
        return False

    role = ""
    try:
        role = str(st.session_state.get("user_info", {}).get("role", "")).strip()
    except Exception:
        role = ""
    if role == "Super Admin":
        # Super Admin changes are business-critical: attempt immediate upload.
        # If Drive is unavailable, fall back to the pending queue without raising.
        _ensure_drive_service()
        if drive_service is not None:
            try:
                return _sync_saved_file_to_drive_now(local_path)
            except Exception as e:
                print(f"Immediate Super Admin sync failed, queued instead: {e}")
                return _mark_drive_file_dirty(local_path)

    return _mark_drive_file_dirty(local_path)


def _upload_to_drive_bg(local_path, filename):
    # Compatibility shim. Queue for the next manual sync.
    _mark_drive_file_dirty(local_path, make_recovery_snapshot=False)


def _force_sync_all_drive_files():
    """Explicit Super Admin/Director action: upload every existing workbook now."""
    results = []
    for local_path in DRIVE_BACKUP_FILENAMES:
        if not os.path.exists(local_path):
            continue
        ok = _sync_saved_file_to_drive_now(local_path)
        results.append((os.path.basename(local_path), ok))
    return results


def _drive_status_rows():
    files = [
        ("HR Leave Requests", HR_LEAVE_PATH),
        ("HR Daily Rates", HR_DAILY_RATES_PATH),
        ("Store Transactions", STORE_DEDUCTION_PATH),
        ("Store Items", STORE_ITEMS_PATH),
        ("HR Employee Records", HR_EMPLOYEES_PATH),
        ("HR Portal Leave Records", HR_PORTAL_LEAVE_PATH),
        ("Employee Items", EMPLOYEE_ITEMS_PATH),
        ("Item Check-ins", ITEM_CHECKIN_PATH),
    ]
    rows = []
    for label, path in files:
        live = os.path.basename(path)
        backup = DRIVE_BACKUP_FILENAMES.get(path, "")
        local_exists = os.path.exists(path)
        pending = os.path.abspath(path) in DRIVE_PENDING_SYNC
        err = DRIVE_LAST_SYNC_ERROR.get(live) or DRIVE_LAST_SYNC_ERROR.get(backup, "")
        rows.append({
            "File": label,
            "Local": "OK" if local_exists else "Missing",
            "Drive Status": "Pending" if pending else ("Synced" if DRIVE_LAST_SYNC.get(live) else "Not yet synced"),
            "Last Sync": DRIVE_LAST_SYNC.get(live, "-"),
            "Error": err or "",
        })
    return rows


def render_google_drive_status():
    if not st.session_state.get("logged_in"):
        return
    role = str(st.session_state.get("user_info", {}).get("role", "")).strip()
    if role not in {"Super Admin", "Director"}:
        return

    with st.sidebar.expander("☁️ Google Drive Backup Status", expanded=False):
        if drive_service is None:
            st.warning("Google Drive is NOT currently connected.")
            if DRIVE_CONNECTION_ERROR:
                st.caption(f"Reason: {DRIVE_CONNECTION_ERROR}")
            st.caption("Local data is being saved safely. Click the button below to attempt a connection.")
        else:
            st.success("Google Drive backup is connected.")
            st.caption("Google Drive is updated when you click the manual sync button below.")

        if st.button("🔄 Save All Data to Google Drive Now", key="force_drive_sync_all"):
            with st.spinner("Connecting to Google Drive and uploading all workbooks..."):
                _ensure_drive_service()
                if drive_service is None:
                    st.error(f"Cannot sync. {DRIVE_CONNECTION_ERROR or 'Google Drive is not connected.'}")
                    st.info("Local data remains safe on disk.")
                else:
                    results = _force_sync_all_drive_files()
                    for label, ok in results:
                        st.write(("✅ " if ok else "❌ ") + label)
                    st.rerun()

        rows = _drive_status_rows()
        st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
        pending = len(DRIVE_PENDING_SYNC)
        st.caption(f"Pending Drive uploads: {pending}. Local saves never wait for Google Drive.")


def get_onedrive_token():
    if not ONEDRIVE_CLIENT_ID or not ONEDRIVE_CLIENT_SECRET:
        return None
    try:
        url = f"https://login.microsoftonline.com/{ONEDRIVE_TENANT_ID}/oauth2/v2.0/token"
        data = {"grant_type": "client_credentials", "client_id": ONEDRIVE_CLIENT_ID, "client_secret": ONEDRIVE_CLIENT_SECRET, "scope": "https://graph.microsoft.com/.default"}
        res = requests.post(url, data=data, timeout=30)
        if res.status_code == 200:
            return res.json().get("access_token")
    except Exception as e:
        st.warning(f"⚠️ OneDrive connection: {e}")
    return None


def upload_to_onedrive(local_file_path, remote_filename=None):
    if not USE_ONEDRIVE:
        return False
    token = get_onedrive_token()
    if not token:
        return False
    filename = remote_filename or os.path.basename(local_file_path)
    remote_path = f"{ONEDRIVE_FOLDER}{filename}"
    try:
        with open(local_file_path, 'rb') as f:
            file_content = f.read()
        url = f"https://graph.microsoft.com/v1.0/drives/me/items/root:/{remote_path}:/content"
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/octet-stream"}
        res = requests.put(url, data=file_content, headers=headers, timeout=60)
        if res.status_code in (200, 201):
            return True
    except Exception:
        pass
    return False


DEFAULT_CATEGORIES = ["Food Allowance", "Others", "Parking", "Parking Fine", "GYM Membership", "Item Not Returned", "Item Missing"]
DEFAULT_ROLES = ["Manager", "Staff", "Team Member", "Employee", "Work Order Employee", "Work Order Manager", "Director", "Payroll", "Super Admin"]
DEFAULT_DEPARTMENTS = [
    "National Grid",
    "Isolator",
    "Projects",
    "Project",
    "Accounts",
    "Payroll Department",
    "ACoole Electrical Ltd",
    "Store",
    "HR",
]
EXCEL_COLUMNS = ["ID", "Employee Name", "Department", "Transaction Type", "Category Reason", "Date", "Amount (£)", "Line Manager", "Description", "Attachment Name", "Status", "Director Comments", "Decision Date", "Decision By", "Submitted By", "PDF File Path", "Edited From ID", "Old Data"]
WORK_ORDER_COLUMNS = ["Work Order ID", "Manual Work Order No.", "Employee Name", "Department", "Work Date", "Hours", "Amount (£)", "Manager", "Description", "Attachment Name", "Status", "Site Address", "Customer Job No.", "Manager Comments", "Manager Decision Date", "Manager Decision By", "Director Comments", "Director Decision Date", "Director Decision By", "Submitted By", "Submitted Date", "Payroll Status", "Payroll Date", "Payroll By", "PDF File Path"]
INSPECTOR_BONUS_COLUMNS = ["ID", "Inspector Name", "Month & Year", "Days Absent", "Reasons for Absence", "Total Jobs Completed", "Bonus Amount (£)", "Status", "Director Comments", "Director Decision Date", "Director Decision By", "Submitted By", "Submitted Date", "PDF File Path"]

DEFAULT_USERS = [
    {"full_name": "National Grid Manager", "username": "national_grid", "password": "acoole123", "role": "Manager", "dept": "National Grid", "can_access_inspector_bonus": True, "can_access_addition_deduction": True, "can_access_work_orders": False, "can_access_wo_total": False, "can_access_hr_leave": True, "can_access_leave_request": True, "can_access_store_deduction": True, "is_active": True},
    {"full_name": "Isolator Manager", "username": "isolator", "password": "acoole123", "role": "Manager", "dept": "Isolator", "can_access_hr_leave": True, "can_access_leave_request": True, "can_access_store_deduction": True, "is_active": True},
    {"full_name": "Project Manager", "username": "project", "password": "acoole123", "role": "Manager", "dept": "Project", "can_access_hr_leave": True, "can_access_leave_request": True, "can_access_store_deduction": True, "is_active": True},
    {"full_name": "Accounts Manager", "username": "accounts", "password": "acoole123", "role": "Manager", "dept": "Accounts", "can_access_hr_leave": True, "can_access_leave_request": True, "can_access_store_deduction": True, "is_active": True},
    {"full_name": "Andy Acoole", "username": "andy", "password": "andy2026", "role": "Director", "dept": "ACoole Electrical Ltd", "can_access_hr_leave": False, "can_access_store_deduction": True, "is_active": True},
    {"full_name": "System Administrator", "username": "wais", "password": "superadmin123", "role": "Super Admin", "dept": "System Administration", "can_access_hr_leave": False, "can_access_store_deduction": True, "is_active": True},
    {"full_name": "Payroll Team", "username": "payroll", "password": "payroll2026", "role": "Payroll", "dept": "Payroll Department", "can_access_store_deduction": True, "is_active": True}
]
PERMISSION_DEFAULTS = {
    "Employee": {"can_view_all_dept": False, "can_generate_pdf": False, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False, "can_access_addition_deduction": False, "can_access_work_orders": False, "can_access_wo_total": False, "can_access_hr_leave": False, "can_access_leave_request": False, "can_access_employee_hr_reports": True, "can_access_store_deduction": False, "can_access_employee_items": False},
    "Work Order Employee": {"can_view_all_dept": False, "can_generate_pdf": False, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False, "can_access_addition_deduction": False, "can_access_work_orders": True, "can_access_wo_total": False, "can_access_hr_leave": False, "can_access_leave_request": False, "can_access_store_deduction": False, "can_access_employee_items": False},
    "Work Order Manager": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False, "can_access_addition_deduction": False, "can_access_work_orders": True, "can_access_wo_total": True, "can_access_hr_leave": False, "can_access_leave_request": False, "can_access_store_deduction": False, "can_access_employee_items": False},
    "Staff": {"can_view_all_dept": False, "can_generate_pdf": False, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False, "can_access_addition_deduction": True, "can_access_work_orders": False, "can_access_wo_total": False, "can_access_hr_leave": False, "can_access_store_deduction": True, "can_access_employee_items": True},
    "Team Member": {"can_view_all_dept": True, "can_generate_pdf": False, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False, "can_access_addition_deduction": True, "can_access_work_orders": False, "can_access_wo_total": False, "can_access_hr_leave": False, "can_access_store_deduction": True, "can_access_employee_items": True},
    "Manager": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": True, "can_access_addition_deduction": True, "can_access_work_orders": False, "can_access_wo_total": False, "can_access_hr_leave": True, "can_access_leave_request": False, "can_access_store_deduction": True, "can_access_employee_items": True},
    "Director": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True, "can_approve_requests": True, "can_access_inspector_bonus": True, "can_access_addition_deduction": True, "can_access_work_orders": True, "can_access_wo_total": True, "can_access_hr_leave": False, "can_access_store_deduction": True, "can_access_employee_items": True},
    "Payroll": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True, "can_approve_requests": False, "can_access_inspector_bonus": True, "can_access_addition_deduction": True, "can_access_work_orders": True, "can_access_wo_total": True, "can_access_hr_leave": True, "can_access_leave_request": False, "can_access_store_deduction": True, "can_access_employee_items": True},
    "Super Admin": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True, "can_approve_requests": True, "can_access_inspector_bonus": True, "can_access_addition_deduction": True, "can_access_work_orders": True, "can_access_wo_total": True, "can_access_hr_leave": False, "can_access_store_deduction": True, "can_access_employee_items": True}
}

_HR_REPORT_PERMISSION_DEFAULTS = {
    "can_access_holiday_calendar": False,
    "can_access_hr_reports": False,
    "can_access_employee_overview": False,
    "can_access_holiday_calculator": False,
    "can_access_employee_directory": False,
}
for _role_name in PERMISSION_DEFAULTS:
    PERMISSION_DEFAULTS[_role_name].update(_HR_REPORT_PERMISSION_DEFAULTS)

PERMISSION_LABELS = {
    "can_view_all_dept": "👁️ View All Department Requests",
    "can_generate_pdf": "📄 Generate & Download PDFs",
    "can_download_data": "📥 Download Data Backups",
    "can_approve_requests": "✅ Approve/Reject Requests",
    "can_access_inspector_bonus": "💰 National Grid Inspector Bonus",
    "can_access_addition_deduction": "➕ Addition & Deduction",
    "can_access_work_orders": "🛠️ Work Orders",
    "can_access_wo_total": "💷 Approved Work Order Total",
    "can_access_hr_leave": "🏢 HR Department (Employee Management + Leave Settlement)",
    "can_access_leave_request": "📝 Leave Request (Department Manager)",
    "can_access_employee_hr_reports": "👤 Employee HR Reports (View Only)",
    "can_access_holiday_calendar": "📅 Holiday Calendar (View Only)",
    "can_access_hr_reports": "📥 HR Reports (Download)",
    "can_access_employee_overview": "📊 Employee Overview (View Only)",
    "can_access_holiday_calculator": "🧮 Holiday Calculator (View Only)",
    "can_access_employee_directory": "👥 Employee Directory (View Only)",
    "can_access_store_deduction": "📦 Store Department Deduction",
    "can_access_employee_items": "🧰 Employee Items / Holdings"
}

try:
    from fpdf import FPDF
    PDF_AVAILABLE = True
except ImportError:
    FPDF = None
    PDF_AVAILABLE = False

_EXCEL_INIT_LOCK = threading.Lock()
_HR_PORTAL_LEAVE_LOCK = threading.RLock()


def _write_empty_excel(path, columns):
    parent = os.path.dirname(os.path.abspath(path))
    os.makedirs(parent, exist_ok=True)
    df = pd.DataFrame(columns=columns)
    stem = os.path.splitext(os.path.basename(path))[0]
    temp_path = os.path.join(parent, f".{stem}.clear_{os.getpid()}_{threading.get_ident()}.xlsx")
    try:
        with _EXCEL_INIT_LOCK:
            df.to_excel(temp_path, index=False, engine="openpyxl")
            os.replace(temp_path, path)
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


def safe_init_excel(path, columns):
    """Fast startup initializer. Never opens an existing workbook during app import."""
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    if os.path.exists(path) and os.path.getsize(path) > 0:
        return True
    tmp_path = f"{path}.init.tmp.xlsx"
    try:
        pd.DataFrame(columns=columns).to_excel(tmp_path, index=False, engine="openpyxl")
        os.replace(tmp_path, path)
        return True
    except Exception as e:
        print(f"Excel initialisation for {path} failed: {e}")
        return False
    finally:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass


def _invalidate_data_cache(*names):
    for name in names:
        st.session_state.pop(name, None)


def _set_data_cache(name, value):
    st.session_state[name] = value
    return value


def _read_excel_records(path):
    if not os.path.exists(path):
        return pd.DataFrame()
    return pd.read_excel(path, engine="openpyxl").fillna("")


def init_audit_log():
    safe_init_excel(AUDIT_LOG_PATH, AUDIT_COLUMNS)


def load_audit_log(force=False):
    init_audit_log()
    if not force and "_audit_log_cache" in st.session_state:
        return list(st.session_state["_audit_log_cache"])
    try:
        df = _read_excel_records(AUDIT_LOG_PATH)
        records = df.to_dict(orient="records")
        return _set_data_cache("_audit_log_cache", records).copy()
    except Exception as e:
        print(f"⚠️ Failed to load audit log: {e}")
        return []


def _get_next_audit_id():
    if "_audit_log_count" not in st.session_state:
        try:
            df = _read_excel_records(AUDIT_LOG_PATH)
            st.session_state["_audit_log_count"] = len(df)
        except Exception:
            st.session_state["_audit_log_count"] = 0
    st.session_state["_audit_log_count"] += 1
    return st.session_state["_audit_log_count"]


def save_audit_entry(entry):
    init_audit_log()
    try:
        df = pd.read_excel(AUDIT_LOG_PATH, engine="openpyxl").fillna("")
        df = pd.concat([df, pd.DataFrame([entry])], ignore_index=True)
        df.to_excel(AUDIT_LOG_PATH, index=False, engine="openpyxl")
        _invalidate_data_cache("_audit_log_cache")
        if not st.session_state.get("_audit_sync_queued"):
            st.session_state["_audit_sync_queued"] = True
            sync_saved_file_to_drive(AUDIT_LOG_PATH)
            st.session_state["_audit_sync_queued"] = False
    except Exception as e:
        print(f"⚠️ Failed to save audit entry: {e}")


def clear_audit_log_file():
    _write_empty_excel(AUDIT_LOG_FILE, AUDIT_COLUMNS)
    _invalidate_data_cache("_audit_log_cache")
    st.session_state["_audit_log_count"] = 0
    sync_saved_file_to_drive(AUDIT_LOG_FILE)


def clear_all_requests_file():
    _write_empty_excel(EXCEL_PATH, EXCEL_COLUMNS)
    _set_data_cache("_records_cache", [])
    sync_saved_file_to_drive(EXCEL_PATH)
    st.session_state["_live_data_reset"] = datetime.now().isoformat()


def clear_all_inspector_bonus():
    _write_empty_excel(INSPECTOR_BONUS_PATH, INSPECTOR_BONUS_COLUMNS)
    _set_data_cache("_inspector_bonus_cache", [])
    sync_saved_file_to_drive(INSPECTOR_BONUS_PATH)


def clear_all_hr_leave():
    _write_empty_excel(HR_LEAVE_PATH, HR_LEAVE_COLUMNS)
    _set_data_cache("_hr_leave_cache", [])
    sync_saved_file_to_drive(HR_LEAVE_PATH)


def clear_all_store_deductions():
    _write_empty_excel(STORE_DEDUCTION_PATH, STORE_DEDUCTION_COLUMNS)
    _set_data_cache("_store_deduction_cache", [])
    sync_saved_file_to_drive(STORE_DEDUCTION_PATH)


def clear_live_request_and_audit_data():
    clear_all_requests_file()
    clear_audit_log_file()


def get_request_details(req_id):
    dept, amount, decision_by, decision_date = "-", "-", "-", "-"
    try:
        all_recs = load_records_from_excel()
        req = next((r for r in all_recs if str(r.get("id")) == str(req_id)), None)
        if req:
            dept = req.get("dept", "-")
            amt = req.get("amount", "-")
            amount = f"£{amt:.2f}" if amt and amt != "-" else "-"
            decision_by = req.get("decision_by", "-") or "-"
            decision_date = req.get("decision_date", "-") or "-"
    except Exception as e:
        print(f"⚠️ Audit lookup failed: {e}")
    return dept, amount, decision_by, decision_date


def _audit_json(value):
    try:
        return json.dumps(value, ensure_ascii=False, default=str)[:300] if value else "-"
    except Exception:
        return str(value)[:300] if value is not None else "-"


def log_action(action, req_id="-", old_data=None, new_data=None, fields_changed=None, decision_by=None, decision_date=None):
    if not st.session_state.get("logged_in"):
        return
    user_info = st.session_state.user_info
    username = user_info.get("full_name", user_info.get("username", "Unknown"))
    role = user_info.get("role", "Unknown")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    SETTING_ACTIONS = [
        "CATEGORY_ADDED", "CATEGORY_EDITED", "CATEGORY_DELETED",
        "DEPARTMENT_ADDED", "DEPARTMENT_EDITED", "DEPARTMENT_DELETED",
        "ROLE_ADDED", "ROLE_EDITED", "ROLE_DELETED",
        "USER_CREATED", "USER_EDITED", "USER_DELETED",
        "PASSWORD_CHANGED", "PASSWORD_RESET",
        "INSPECTOR_BONUS_CREATED", "INSPECTOR_BONUS_APPROVED",
        "INSPECTOR_BONUS_REJECTED", "INSPECTOR_BONUS_EDITED",
        "INSPECTOR_BONUS_STATUS_CHANGED", "SUPER_ADMIN_CLEAR_INSPECTOR_BONUS",
        "HR_CATEGORY_ADDED", "HR_CATEGORY_DELETED",
        "HR_RATE_ADDED", "HR_RATE_UPDATED", "HR_RATE_DELETED",
        "STORE_ITEM_ADDED", "STORE_ITEM_EDITED", "STORE_ITEM_DELETED",
        "HR_EMPLOYEE_ADDED", "HR_EMPLOYEE_EDITED", "HR_LEAVE_RECORDED", "HR_LEAVE_EDITED", "HR_LEAVE_DELETED",
        "HR_ENTITLEMENT_ADJUSTED", "HR_PORTAL_LEAVE_REQUESTED", "HR_PORTAL_LEAVE_APPROVED", "HR_PORTAL_LEAVE_REJECTED",
        "EMPLOYEE_ITEMS_ISSUED", "ITEM_CHECKIN_COMPLETED",
    ]
    if action in SETTING_ACTIONS:
        action_labels = {
            "CATEGORY_ADDED": "🏷️ Category Added", "CATEGORY_EDITED": "🏷️ Category Edited", "CATEGORY_DELETED": "🏷️ Category Deleted",
            "DEPARTMENT_ADDED": "🏢 Department Added", "DEPARTMENT_EDITED": "🏢 Department Edited", "DEPARTMENT_DELETED": "🏢 Department Deleted",
            "ROLE_ADDED": "🎖️ Role Added", "ROLE_EDITED": "🎖️ Role/Permissions Edited", "ROLE_DELETED": "🎖️ Role Deleted",
            "USER_CREATED": "👤 User Account Created", "USER_EDITED": "👤 User Account Edited",
            "USER_DELETED": "👤 User Account Deleted", "PASSWORD_CHANGED": "🔑 Password Changed", "PASSWORD_RESET": "🔑 Password Reset",
            "INSPECTOR_BONUS_CREATED": "💰 Inspector Bonus Submitted", "INSPECTOR_BONUS_APPROVED": "💰 Inspector Bonus Approved",
            "INSPECTOR_BONUS_REJECTED": "💰 Inspector Bonus Rejected", "INSPECTOR_BONUS_EDITED": "💰 Inspector Bonus Edited",
            "INSPECTOR_BONUS_STATUS_CHANGED": "💰 Inspector Bonus Status Changed",
            "SUPER_ADMIN_CLEAR_INSPECTOR_BONUS": "🧹 Super Admin — All Inspector Bonuses Cleared",
            "HR_CATEGORY_ADDED": "🏷️ HR Category Added", "HR_CATEGORY_DELETED": "🏷️ HR Category Deleted",
            "HR_RATE_ADDED": "💷 HR Daily Rate Added", "HR_RATE_UPDATED": "💷 HR Daily Rate Updated", "HR_RATE_DELETED": "💷 HR Daily Rate Deleted",
            "STORE_ITEM_ADDED": "📦 Store Item Added", "STORE_ITEM_EDITED": "📦 Store Item Edited", "STORE_ITEM_DELETED": "📦 Store Item Deleted",
            "HR_EMPLOYEE_ADDED": "🧑💼 HR Employee Added", "HR_EMPLOYEE_EDITED": "🧑💼 HR Employee Edited",
            "HR_LEAVE_RECORDED": "📅 HR Leave Recorded", "HR_LEAVE_EDITED": "✏️ HR Leave Edited", "HR_LEAVE_DELETED": "🗑️ HR Leave Deleted", "HR_ENTITLEMENT_ADJUSTED": "📊 HR Entitlement Adjusted",
            "HR_PORTAL_LEAVE_REQUESTED": "📝 Department Leave Requested", "HR_PORTAL_LEAVE_APPROVED": "✅ Department Leave Approved", "HR_PORTAL_LEAVE_REJECTED": "❌ Department Leave Rejected",
            "EMPLOYEE_ITEMS_ISSUED": "🧰 Employee Items Issued", "ITEM_CHECKIN_COMPLETED": "📋 Item Check-in Completed",
        }
        display_action = action_labels.get(action, action)
        old_val = _audit_json(old_data)
        new_val = _audit_json(new_data)
        save_audit_entry({"AuditID": _get_next_audit_id(), "Timestamp": timestamp, "User_Name": username, "User_Role": role, "Action": display_action, "Request_ID": str(req_id), "Department": "-", "Amount": "-", "Decision_By": decision_by or "-", "Decision_Date": decision_date or "-", "Field_Changed": "Settings", "Old_Value": old_val, "New_Value": new_val, "IP_Address": "Auto-Logged"})
        return
    dept, amount, saved_decision_by, saved_decision_date = get_request_details(req_id)
    final_decision_by = decision_by or saved_decision_by
    final_decision_date = decision_date or saved_decision_date
    if action.startswith("WORK_ORDER_"):
        wo_labels = {
            "WORK_ORDER_CREATED": "🛠️ Work Order Created", "WORK_ORDER_MANAGER_CREATED": "🛠️ Work Order Submitted to Director",
            "WORK_ORDER_MANAGER_APPROVED": "🛠️ Work Order Approved by Manager → Director", "WORK_ORDER_MANAGER_REJECTED": "🛠️ Work Order Rejected by Manager → Returned to Employee",
            "WORK_ORDER_RETURNED": "🛠️ Work Order Returned to Employee", "WORK_ORDER_DIRECTOR_APPROVED": "🛠️ Work Order Approved for Payment",
            "WORK_ORDER_DIRECTOR_REJECTED": "🛠️ Work Order Rejected by Director", "WORK_ORDER_STATUS_CHANGED": "🛠️ Work Order Status Changed",
            "WORK_ORDER_EDITED": "🛠️ Work Order Edited", "WORK_ORDER_PAID": "🛠️ Work Order Paid",
        }
        old_v = _audit_json(old_data)
        new_v = _audit_json(new_data)
        save_audit_entry({"AuditID": _get_next_audit_id(), "Timestamp": timestamp, "User_Name": username, "User_Role": role, "Action": wo_labels.get(action, action), "Request_ID": str(req_id), "Department": dept, "Amount": amount, "Decision_By": final_decision_by or "-", "Decision_Date": final_decision_date or timestamp, "Field_Changed": "Work Order Status", "Old_Value": old_v if old_data else "-", "New_Value": new_v if new_data else action.replace("WORK_ORDER_", "").replace("_", " ").title(), "IP_Address": "Auto-Logged"})
        return
    if action.startswith("HR_LEAVE_"):
        hr_labels = {
            "HR_LEAVE_CREATED": "👥 HR Leave Settlement Created",
            "HR_LEAVE_APPROVED": "👥 HR Leave Settlement Approved",
            "HR_LEAVE_REJECTED": "👥 HR Leave Settlement Rejected",
            "HR_LEAVE_STATUS_CHANGED": "👥 HR Leave Settlement Status Changed",
        }
        old_v = _audit_json(old_data)
        new_v = _audit_json(new_data)
        save_audit_entry({"AuditID": _get_next_audit_id(), "Timestamp": timestamp, "User_Name": username, "User_Role": role, "Action": hr_labels.get(action, action), "Request_ID": str(req_id), "Department": "-", "Amount": "-", "Decision_By": final_decision_by or "-", "Decision_Date": final_decision_date or timestamp, "Field_Changed": "HR Leave Status", "Old_Value": old_v, "New_Value": new_v, "IP_Address": "Auto-Logged"})
        return
    if action.startswith("STORE_DEDUCTION_"):
        store_labels = {
            "STORE_DEDUCTION_CREATED": "📦 Store Deduction Created",
            "STORE_DEDUCTION_APPROVED": "📦 Store Deduction Approved",
            "STORE_DEDUCTION_REJECTED": "📦 Store Deduction Rejected",
            "STORE_DEDUCTION_STATUS_CHANGED": "📦 Store Deduction Status Changed",
        }
        old_v = _audit_json(old_data)
        new_v = _audit_json(new_data)
        save_audit_entry({"AuditID": _get_next_audit_id(), "Timestamp": timestamp, "User_Name": username, "User_Role": role, "Action": store_labels.get(action, action), "Request_ID": str(req_id), "Department": "-", "Amount": "-", "Decision_By": final_decision_by or "-", "Decision_Date": final_decision_date or timestamp, "Field_Changed": "Store Deduction Status", "Old_Value": old_v, "New_Value": new_v, "IP_Address": "Auto-Logged"})
        return
    if action in ["CREATED", "DELETED"]:
        save_audit_entry({"AuditID": _get_next_audit_id(), "Timestamp": timestamp, "User_Name": username, "User_Role": role, "Action": action, "Request_ID": str(req_id), "Department": dept, "Amount": amount, "Decision_By": "-", "Decision_Date": "-", "Field_Changed": "-", "Old_Value": "-", "New_Value": "New Request Created" if action == "CREATED" else "Request Permanently Deleted", "IP_Address": "Auto-Logged"})
    elif action in ["APPROVED", "REJECTED", "STATUS_CHANGED"]:
        status_text = "Approved" if action == "APPROVED" else "Rejected" if action == "REJECTED" else "Status Changed"
        save_audit_entry({"AuditID": _get_next_audit_id(), "Timestamp": timestamp, "User_Name": username, "User_Role": role, "Action": action, "Request_ID": str(req_id), "Department": dept, "Amount": amount, "Decision_By": final_decision_by, "Decision_Date": final_decision_date, "Field_Changed": "Status", "Old_Value": "Pending", "New_Value": status_text, "IP_Address": "Auto-Logged"})
    elif action == "EDITED" and old_data and new_data:
        field_labels = {"emp_name": "Employee Name", "dept": "Department", "type": "Transaction Type", "category": "Category", "date": "Date", "amount": "Amount (£)", "manager": "Line Manager", "desc": "Description", "status": "Status", "attachment_name": "Attachments"}
        for key, label in field_labels.items():
            old = str(old_data.get(key, "")).strip()
            new = str(new_data.get(key, "")).strip()
            if old != new:
                save_audit_entry({"AuditID": _get_next_audit_id(), "Timestamp": timestamp, "User_Name": username, "User_Role": role, "Action": "EDITED", "Request_ID": str(req_id), "Department": dept, "Amount": amount, "Decision_By": "-", "Decision_Date": "-", "Field_Changed": label, "Old_Value": old, "New_Value": new, "IP_Address": "Auto-Logged"})


def display_audit_log_panel():
    st.subheader("📖 Full System Audit Log — Complete History")
    st.info("🔒 Super Admin Only — The audit history can be cleared from the Danger Zone below.")
    st.divider()
    logs = load_audit_log()
    if not logs:
        st.info("📋 No activity recorded yet.")
        return
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        filter_user = st.multiselect("👤 Filter by User", sorted(set([l["User_Name"] for l in logs])))
    with c2:
        dept_list = sorted(set([l.get("Department", "") for l in logs if l.get("Department") != "-"]))
        filter_dept = st.multiselect("🏢 Filter by Department", dept_list)
    with c3:
        filter_action = st.multiselect("🔧 Filter by Action", sorted(set([l["Action"] for l in logs])))
    with c4:
        req_list = sorted(set([str(l["Request_ID"]) for l in logs if str(l["Request_ID"]) != "-"]))
        filter_req = st.multiselect("🆔 Filter by Request ID", req_list)
    filtered = logs
    if filter_user:
        filtered = [l for l in filtered if l["User_Name"] in filter_user]
    if filter_dept:
        filtered = [l for l in filtered if l.get("Department", "") in filter_dept]
    if filter_action:
        filtered = [l for l in filtered if l["Action"] in filter_action]
    if filter_req:
        filtered = [l for l in filtered if str(l["Request_ID"]) in filter_req]
    st.metric("📄 Total Entries", len(filtered))
    st.divider()
    for entry in reversed(filtered):
        aid, ts, user, role, action, req_id = entry["AuditID"], entry["Timestamp"], entry["User_Name"], entry["User_Role"], entry["Action"], entry["Request_ID"]
        dept, amount, dec_by, dec_date = entry.get("Department", "-"), entry.get("Amount", "-"), entry.get("Decision_By", "-"), entry.get("Decision_Date", "-")
        field, old_val, new_val = entry["Field_Changed"], entry["Old_Value"], entry["New_Value"]
        icon = {"CREATED": "➕", "EDITED": "✏️", "APPROVED": "✅", "REJECTED": "❌", "DELETED": "🗑️", "STATUS_CHANGED": "🔄"}.get(action, "ℹ️")
        title = f"{icon} {action}" + (f" — Request #{req_id}" if str(req_id) != "-" else "") + f" | {user} ({role}) | {ts}"
        with st.expander(title):
            st.write(f"**🕐 Time:** {ts}")
            st.write(f"**👤 User:** {user} — *{role}*")
            if str(req_id) != "-":
                st.write(f"**🆔 Request ID:** #{req_id}")
            if dept != "-":
                st.write(f"**🏢 Department:** {dept}")
            if amount != "-":
                st.write(f"**💷 Amount:** {amount}")
            if action in ["APPROVED", "REJECTED", "STATUS_CHANGED"]:
                st.write(f"**🎯 Decision By:** {dec_by}")
                st.write(f"**📅 Decision Date:** {dec_date}")
            if field and field != "-" and field != "No Changes":
                st.write(f"**📝 Field Changed:** {field}")
                if old_val and old_val != "-":
                    st.markdown(f"**⬅️ Old:** `{old_val}`")
                if new_val and new_val != "-":
                    st.markdown(f"**➡️ New:** `{new_val}`")
            else:
                st.write(f"**📋 Details:** {new_val}")
    st.divider()
    df_export = pd.DataFrame(filtered)
    st.download_button("📥 Download Full Audit Log (CSV)", df_export.to_csv(index=False).encode("utf-8"), "Acoole_Audit_Log.csv", type="primary")


def format_date(d):
    if not d or str(d).strip().lower() in ["", "none", "nan"]:
        return "-"
    return str(d).strip()


def get_next_id(all_records):
    if not all_records:
        return 1
    return max(int(r.get("id", 0)) for r in all_records) + 1


def show_old_new_comparison(old_json, new_rec):
    try:
        old = json.loads(old_json) if old_json and old_json != "{}" else {}
    except Exception:
        old = {}
    if not old:
        st.info("📋 New request — no previous version.")
        return
    st.markdown("#### 🔄 Changes (Previous → New)")
    fields = [("emp_name", "Employee Name"), ("dept", "Department"), ("type", "Transaction Type"), ("category", "Category"), ("date", "Date"), ("amount", "Amount (£)"), ("manager", "Line Manager"), ("desc", "Description")]
    changed = False
    for key, label in fields:
        o, n = str(old.get(key, "")).strip(), str(new_rec.get(key, "")).strip()
        if o != n:
            changed = True
            st.markdown(f"**{label}**: ~~`{o}`~~ → **`{n}`**")
    if not changed:
        st.info("✅ No changes detected.")


def refresh_data_button():
    if st.button("🔄 Refresh Data", type="secondary", key="refresh_data_btn"):
        with st.spinner("Refreshing..."):
            _ensure_drive_service()
            if drive_service is not None:
                for path in (EXCEL_PATH, USER_DB_PATH, SETTINGS_PATH, AUDIT_LOG_PATH, INSPECTOR_BONUS_PATH, WORK_ORDERS_PATH, HR_LEAVE_PATH, HR_DAILY_RATES_PATH, STORE_DEDUCTION_PATH, STORE_ITEMS_PATH, HR_EMPLOYEES_PATH, HR_PORTAL_LEAVE_PATH, EMPLOYEE_ITEMS_PATH, ITEM_CHECKIN_PATH, LEAVER_CLEARANCE_PATH):
                    with _DRIVE_SYNC_LOCK:
                        pending = path in DRIVE_PENDING_SYNC
                    if pending:
                        continue
                    remote = _drive_find_file(os.path.basename(path))
                    if remote:
                        _drive_download_file(remote["id"], path)
            _invalidate_data_cache("_records_cache", "_users_cache", "_settings_cache", "_audit_log_cache", "_work_orders_cache", "_inspector_bonus_cache", "_audit_log_count", "_hr_leave_cache", "_hr_daily_rates_cache", "_store_deduction_cache", "_store_items_cache", "_employee_items_cache", "_item_checkins_cache", "_leaver_clearance_cache")
            st.session_state["hrp_force_reload"] = True
            st.session_state["hrp_data_loaded"] = False
            st.session_state["_last_refresh"] = datetime.now().isoformat()
        st.rerun()


def init_settings():
    if not os.path.exists(SETTINGS_PATH):
        pd.DataFrame([
            {"setting": "categories", "value": "|".join(DEFAULT_CATEGORIES)},
            {"setting": "roles", "value": "|".join(DEFAULT_ROLES)},
            {"setting": "departments", "value": "|".join(DEFAULT_DEPARTMENTS)},
            {"setting": "hr_categories", "value": "|".join(DEFAULT_HR_CATEGORIES)},
        ]).to_excel(SETTINGS_PATH, index=False, engine="openpyxl")


def _load_setting_value(setting_name, default_values):
    init_settings()
    if "_settings_cache" not in st.session_state:
        try:
            df = _read_excel_records(SETTINGS_PATH)
            settings = {str(r.get("setting", "")).strip(): str(r.get("value", "")) for r in df.to_dict(orient="records")}
            st.session_state["_settings_cache"] = settings
        except Exception:
            st.session_state["_settings_cache"] = {}
    raw = st.session_state["_settings_cache"].get(setting_name, "")
    vals = [v.strip() for v in str(raw).split("|") if v.strip()]
    return vals if vals else list(default_values)


def load_departments():
    departments = _load_setting_value("departments", DEFAULT_DEPARTMENTS)
    departments = [str(d).strip() for d in departments if str(d).strip()]
    for required in DEFAULT_DEPARTMENTS:
        if required not in departments:
            departments.append(required)
    return departments


def save_departments(dept_list):
    init_settings()
    df = _read_excel_records(SETTINGS_PATH)
    found = False
    for idx, r in df.iterrows():
        if r["setting"] == "departments":
            df.at[idx, "value"] = "|".join(dept_list)
            found = True
    if not found:
        df = pd.concat([df, pd.DataFrame([{"setting": "departments", "value": "|".join(dept_list)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False, engine="openpyxl")
    _invalidate_data_cache("_settings_cache")
    sync_saved_file_to_drive(SETTINGS_PATH)


def load_categories():
    return _load_setting_value("categories", DEFAULT_CATEGORIES)


def save_categories(cat_list):
    init_settings()
    df = _read_excel_records(SETTINGS_PATH)
    found = False
    for idx, r in df.iterrows():
        if r["setting"] == "categories":
            df.at[idx, "value"] = "|".join(cat_list)
            found = True
    if not found:
        df = pd.concat([df, pd.DataFrame([{"setting": "categories", "value": "|".join(cat_list)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False, engine="openpyxl")
    _invalidate_data_cache("_settings_cache")
    sync_saved_file_to_drive(SETTINGS_PATH)


def load_roles():
    roles = _load_setting_value("roles", DEFAULT_ROLES)
    missing = [r for r in DEFAULT_ROLES if r not in roles]
    if missing:
        roles = roles + missing
        save_roles(roles)
    return roles


def save_roles(roles_list):
    init_settings()
    df = _read_excel_records(SETTINGS_PATH)
    found = False
    for idx, r in df.iterrows():
        if r["setting"] == "roles":
            df.at[idx, "value"] = "|".join(roles_list)
            found = True
    if not found:
        df = pd.concat([df, pd.DataFrame([{"setting": "roles", "value": "|".join(roles_list)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False, engine="openpyxl")
    _invalidate_data_cache("_settings_cache")
    sync_saved_file_to_drive(SETTINGS_PATH)


def load_hr_categories():
    return _load_setting_value("hr_categories", DEFAULT_HR_CATEGORIES)


def save_hr_categories(cats):
    init_settings()
    df = _read_excel_records(SETTINGS_PATH)
    found = False
    for idx, r in df.iterrows():
        if r["setting"] == "hr_categories":
            df.at[idx, "value"] = "|".join(cats)
            found = True
    if not found:
        df = pd.concat([df, pd.DataFrame([{"setting": "hr_categories", "value": "|".join(cats)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False, engine="openpyxl")
    _invalidate_data_cache("_settings_cache")
    sync_saved_file_to_drive(SETTINGS_PATH)


def init_user_db():
    safe_init_excel(USER_DB_PATH, USER_DB_COLUMNS)
    try:
        df = pd.read_excel(USER_DB_PATH, engine="openpyxl")
        if df.empty:
            pd.DataFrame(DEFAULT_USERS).to_excel(USER_DB_PATH, index=False, engine="openpyxl")
    except Exception:
        pd.DataFrame(DEFAULT_USERS).to_excel(USER_DB_PATH, index=False, engine="openpyxl")


def save_users(users_dict):
    rows = []
    for username, u in users_dict.items():
        rows.append({
            "full_name": u.get("full_name", username), "username": username,
            "password": u.get("password", ""), "role": u.get("role", "Staff"),
            "dept": u.get("dept", ""),
            "can_view_all_dept": u.get("can_view_all_dept", False),
            "can_generate_pdf": u.get("can_generate_pdf", False),
            "can_download_data": u.get("can_download_data", False),
            "can_approve_requests": u.get("can_approve_requests", False),
            "can_access_inspector_bonus": u.get("can_access_inspector_bonus", False),
            "can_access_addition_deduction": u.get("can_access_addition_deduction", False),
            "can_access_work_orders": u.get("can_access_work_orders", False),
            "can_access_wo_total": u.get("can_access_wo_total", False),
            "can_access_hr_leave": u.get("can_access_hr_leave", False),
            "can_access_leave_request": u.get("can_access_leave_request", False),
            "can_access_employee_hr_reports": u.get("can_access_employee_hr_reports", False),
            "can_access_holiday_calendar": u.get("can_access_holiday_calendar", False),
            "can_access_hr_reports": u.get("can_access_hr_reports", False),
            "can_access_employee_overview": u.get("can_access_employee_overview", False),
            "can_access_holiday_calculator": u.get("can_access_holiday_calculator", False),
            "can_access_employee_directory": u.get("can_access_employee_directory", False),
            "employee_id": u.get("employee_id", ""),
            "can_access_store_deduction": u.get("can_access_store_deduction", False),
            "can_access_employee_items": u.get("can_access_employee_items", False),
            "is_active": u.get("is_active", True)
        })
    pd.DataFrame(rows, columns=USER_DB_COLUMNS).to_excel(USER_DB_PATH, index=False, engine="openpyxl")
    _invalidate_data_cache("_users_cache")
    sync_saved_file_to_drive(USER_DB_PATH)


def _flag_or_default(raw_value, role, key):
    s = str(raw_value).strip().lower()
    if s in ("true", "false"):
        return s == "true"
    return PERMISSION_DEFAULTS.get(role, PERMISSION_DEFAULTS["Staff"]).get(key, False)


def _active_or_default(raw_value):
    s = str(raw_value).strip().lower()
    if s in ("true", "false"):
        return s == "true"
    return True


def load_users(force=False):
    init_user_db()
    if not force and "_users_cache" in st.session_state:
        return dict(st.session_state["_users_cache"])
    try:
        df = _read_excel_records(USER_DB_PATH)
        has_leave_request_column = "can_access_leave_request" in df.columns
        users = {}
        for _, r in df.iterrows():
            username = str(r.get("username", "")).strip()
            if not username:
                continue
            user_role = str(r.get("role", "Staff"))
            users[username] = {
                "full_name": str(r.get("full_name", username)).strip(),
                "password": str(r.get("password", "")),
                "role": user_role,
                "dept": str(r.get("dept", "")),
                "can_view_all_dept": _flag_or_default(r.get("can_view_all_dept", ""), user_role, "can_view_all_dept"),
                "can_generate_pdf": _flag_or_default(r.get("can_generate_pdf", ""), user_role, "can_generate_pdf"),
                "can_download_data": _flag_or_default(r.get("can_download_data", ""), user_role, "can_download_data"),
                "can_approve_requests": _flag_or_default(r.get("can_approve_requests", ""), user_role, "can_approve_requests"),
                "can_access_inspector_bonus": _flag_or_default(r.get("can_access_inspector_bonus", ""), user_role, "can_access_inspector_bonus"),
                "can_access_addition_deduction": _flag_or_default(r.get("can_access_addition_deduction", ""), user_role, "can_access_addition_deduction"),
                "can_access_work_orders": _flag_or_default(r.get("can_access_work_orders", ""), user_role, "can_access_work_orders"),
                "can_access_wo_total": _flag_or_default(r.get("can_access_wo_total", ""), user_role, "can_access_wo_total"),
                "can_access_hr_leave": _flag_or_default(r.get("can_access_hr_leave", ""), user_role, "can_access_hr_leave"),
                "can_access_leave_request": _flag_or_default(r.get("can_access_leave_request", ""), user_role, "can_access_leave_request"),
                "can_access_employee_hr_reports": _flag_or_default(r.get("can_access_employee_hr_reports", ""), user_role, "can_access_employee_hr_reports"),
                "can_access_holiday_calendar": _flag_or_default(r.get("can_access_holiday_calendar", ""), user_role, "can_access_holiday_calendar"),
                "can_access_hr_reports": _flag_or_default(r.get("can_access_hr_reports", ""), user_role, "can_access_hr_reports"),
                "can_access_employee_overview": _flag_or_default(r.get("can_access_employee_overview", ""), user_role, "can_access_employee_overview"),
                "can_access_holiday_calculator": _flag_or_default(r.get("can_access_holiday_calculator", ""), user_role, "can_access_holiday_calculator"),
                "can_access_employee_directory": _flag_or_default(r.get("can_access_employee_directory", ""), user_role, "can_access_employee_directory"),
                "employee_id": str(r.get("employee_id", "")).strip(),
                "can_access_store_deduction": _flag_or_default(r.get("can_access_store_deduction", ""), user_role, "can_access_store_deduction"),
                "can_access_employee_items": _flag_or_default(r.get("can_access_employee_items", ""), user_role, "can_access_employee_items"),
                "is_active": _active_or_default(r.get("is_active", ""))
            }
            if user_role == "Super Admin":
                users[username].update({
                    "can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True,
                    "can_approve_requests": True, "can_access_inspector_bonus": True,
                    "can_access_addition_deduction": True, "can_access_work_orders": True,
                    "can_access_wo_total": True, "can_access_hr_leave": False, "can_access_leave_request": False,
                    "can_access_store_deduction": True, "can_access_employee_items": True, "is_active": True
                })
            elif user_role == "Director":
                users[username].update({
                    "can_view_all_dept": True,
                    "can_generate_pdf": True,
                    "can_download_data": True,
                    "can_approve_requests": True,
                    "can_access_inspector_bonus": True,
                    "can_access_addition_deduction": True,
                    "can_access_work_orders": True,
                    "can_access_wo_total": True,
                    "can_access_hr_leave": users[username].get("can_access_hr_leave", False),
                    "can_access_store_deduction": True,
                    "can_access_employee_items": True,
                    "is_active": users[username].get("is_active", True),
                })
        for username, u in users.items():
            if str(u.get("dept", "")).strip().casefold() == "store":
                u["can_access_store_deduction"] = True
                u["can_access_employee_items"] = True
        migrated_leave_request = False
        for username, u in users.items():
            if (not has_leave_request_column) and str(u.get("role", "")).strip().lower() == "manager" and str(u.get("dept", "")).strip().casefold() != "hr" and u.get("can_access_hr_leave"):
                u["can_access_leave_request"] = True
                migrated_leave_request = True
        migrated_builtin_hr = False
        for builtin_username in ("andy", "wais"):
            if builtin_username in users and users[builtin_username].get("can_access_hr_leave"):
                users[builtin_username]["can_access_hr_leave"] = False
                migrated_builtin_hr = True
        if migrated_builtin_hr or migrated_leave_request:
            save_users(users)
        _set_data_cache("_users_cache", users)
        return dict(users)
    except Exception as e:
        st.error(f"User DB Load Error: {e}")
        return {}


# ════════════════════════════════════════════════════════════
# 🏢 HR PORTAL MODULE — Employee / Holiday / Leave Management
# ════════════════════════════════════════════════════════════
HR_PORTAL_DEPARTMENTS = ["HR", "Operations", "Sales", "Admin", "Finance"]
HR_PORTAL_AGREEMENT_TYPES = ["Permanent", "Fixed Term", "Part-Time", "Temporary", "Apprentice", "Contractor"]

HR_DEPARTMENT_WORK_TYPES = {
    "Projects": ["Office Staff", "Site Electrician"],
    "Project": ["Office Staff", "Site Electrician"],
    "National Grid": ["Office Staff", "National Grid Inspector"],
    "Isolator": ["Office Staff", "Isolator Installer"],
}


def _hrp_work_type_options(department):
    return HR_DEPARTMENT_WORK_TYPES.get(str(department or "").strip(), ["Office Staff"])


def _hrp_normalize_work_type(department, work_type="", job_title=""):
    options = _hrp_work_type_options(department)
    raw, title = str(work_type or "").strip(), str(job_title or "").strip()
    if raw in options:
        return raw
    if title in options:
        return title
    aliases = {"site electrician": "Site Electrician", "national grid inspector": "National Grid Inspector", "isolator installer": "Isolator Installer", "office staff": "Office Staff"}
    alias = aliases.get(title.casefold())
    return alias if alias in options else "Office Staff"


HR_PORTAL_LEAVE_TYPES = ["Full Day Holiday", "Half Day Holiday", "Sick Leave", "Family / Emergency Leave", "Unpaid Holiday", "Unpaid Absence", "Maternity Leave", "Paternity Leave", "Other Absence", "College (Apprenticeship)", "Training Course"]
HR_PORTAL_HOLIDAY_LEAVE_TYPES = ["Full Day Holiday", "Half Day Holiday"]


def _hrp_init_storage():
    safe_init_excel(HR_EMPLOYEES_PATH, HR_EMPLOYEE_COLUMNS)
    safe_init_excel(HR_PORTAL_LEAVE_PATH, HR_PORTAL_LEAVE_COLUMNS)


def _hrp_normalize_employee_date(value, default=None):
    if value is None:
        return default
    try:
        missing = pd.isna(value)
        if isinstance(missing, bool) and missing:
            return default
    except Exception:
        pass
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date) and not isinstance(value, pd.Timestamp):
        return value
    try:
        parsed = pd.to_datetime(value, errors="coerce")
        if pd.isna(parsed):
            return default
        return parsed.date()
    except Exception:
        return default


def _hrp_normalize_employee_record(employee):
    if not isinstance(employee, dict):
        return employee
    employee["start_date"] = _hrp_normalize_employee_date(employee.get("start_date"), date.today())
    employee["leaving_date"] = _hrp_normalize_employee_date(employee.get("leaving_date"), None)
    return employee


def _hrp_normalize_employee_records(employees):
    if not isinstance(employees, list):
        return []
    for employee in employees:
        _hrp_normalize_employee_record(employee)
    return employees


def _hrp_load_employees():
    _hrp_init_storage()
    try:
        df = _read_excel_records(HR_EMPLOYEES_PATH)
        records = []
        for r in df.to_dict(orient="records"):
            emp_id = str(r.get("Employee ID", "")).strip()
            if not emp_id:
                continue
            raw_start = r.get("Start Date", "")
            start_date = _hrp_normalize_employee_date(raw_start, date.today())
            override_raw = r.get("Holiday Entitlement Override", "")
            try:
                override = float(override_raw) if str(override_raw).strip() else None
            except Exception:
                override = None
            records.append({
                "emp_id": emp_id,
                "name": str(r.get("Full Name", "")).strip(),
                "start_date": start_date,
                "department": str(r.get("Department", "")).strip() or "Other",
                "job_title": str(r.get("Position / Job Title", "")).strip(),
                "work_type": _hrp_normalize_work_type(str(r.get("Department", "")).strip() or "Other", r.get("Work Type / Sub-department", ""), r.get("Position / Job Title", "")),
                "agreement_type": str(r.get("Agreement Type", "")).strip() or "Permanent",
                "status": str(r.get("Status", "Active")).strip() or "Active",
                "working_pattern": str(r.get("Working Pattern", "Regular hours")).strip() or "Regular hours",
                "days_per_week": float(r.get("Days Worked Per Week", 5) or 5),
                "entitlement_override": override,
                "adjustment_note": str(r.get("Entitlement Adjustment Note", "")).strip(),
                "leaving_date": _hrp_normalize_employee_date(r.get("Leaving Date", ""), None),
                "leaving_reason": str(r.get("Leaving Reason", "")).strip(),
            })
        return records
    except Exception as e:
        print(f"HR employee data load failed: {e}")
        return []


def _hrp_save_employees():
    rows = []
    for e in st.session_state.get("hrp_employees", []):
        rows.append({
            "Employee ID": e.get("emp_id", ""),
            "Full Name": e.get("name", ""),
            "Start Date": e.get("start_date", ""),
            "Position / Job Title": e.get("job_title", ""),
            "Department": e.get("department", ""),
            "Work Type / Sub-department": e.get("work_type", _hrp_normalize_work_type(e.get("department", ""), "", e.get("job_title", ""))),
            "Agreement Type": e.get("agreement_type", ""),
            "Status": e.get("status", "Active"),
            "Working Pattern": e.get("working_pattern", "Regular hours"),
            "Days Worked Per Week": e.get("days_per_week", 5),
            "Holiday Entitlement Override": e.get("entitlement_override", "") if e.get("entitlement_override") is not None else "",
            "Entitlement Adjustment Note": e.get("adjustment_note", ""),
            "Leaving Date": e.get("leaving_date", "") or "",
            "Leaving Reason": e.get("leaving_reason", ""),
        })
    pd.DataFrame(rows, columns=HR_EMPLOYEE_COLUMNS).to_excel(HR_EMPLOYEES_PATH, index=False, engine="openpyxl")
    sync_saved_file_to_drive(HR_EMPLOYEES_PATH)


def _hrp_load_leave_records():
    _hrp_init_storage()
    try:
        df = _read_excel_records(HR_PORTAL_LEAVE_PATH)
        records = []
        max_no = 0
        for row_no, r in enumerate(df.to_dict(orient="records"), start=2):
            leave_id = str(r.get("Leave ID", "")).strip()
            if not leave_id:
                if all(str(v).strip() == "" for v in r.values()):
                    continue
                raise ValueError(f"HR leave workbook row {row_no} has no Leave ID")
            try:
                n = int(str(leave_id).replace("LV-", ""))
                max_no = max(max_no, n)
            except Exception:
                pass

            raw_from = r.get("Date From", "")
            raw_to = r.get("Date To", "")
            d_from_ts = pd.to_datetime(raw_from, errors="coerce")
            d_to_ts = pd.to_datetime(raw_to, errors="coerce")
            if pd.isna(d_from_ts) or pd.isna(d_to_ts):
                raise ValueError(
                    f"HR leave workbook row {row_no} (Leave ID {leave_id}) has an invalid "
                    f"Date From/Date To value: {raw_from!r} / {raw_to!r}"
                )
            d_from = d_from_ts.date()
            d_to = d_to_ts.date()

            try:
                days = float(r.get("Days", 0) or 0)
            except Exception as exc:
                raise ValueError(
                    f"HR leave workbook row {row_no} (Leave ID {leave_id}) has an invalid Days value: {r.get('Days')!r}"
                ) from exc

            records.append({
                "leave_id": leave_id,
                "employee_id": str(r.get("Employee ID", "")).strip(),
                "date_from": d_from, "date_to": d_to,
                "type": str(r.get("Leave Type", "")).strip(),
                "days": days,
                "status": str(r.get("Status", "Approved")).strip() or "Approved",
                "request_source": str(r.get("Request Source", "")).strip() or "Email",
                "request_reference": str(r.get("Request Reference", "")).strip(),
                "notes": str(r.get("Notes", "")).strip(),
                "recorded_by": str(r.get("Recorded By", "")).strip(),
                "recorded_at": str(r.get("Recorded At", "")).strip(),
                "entry_source": str(r.get("Entry Source", "HR Direct")).strip() or "HR Direct",
                "requested_by": str(r.get("Requested By", "")).strip(),
                "requested_at": str(r.get("Requested At", "")).strip(),
                "approved_by": str(r.get("Approved By", "")).strip(),
                "approved_at": str(r.get("Approved At", "")).strip(),
                "rejection_reason": str(r.get("Rejection Reason", "")).strip(),
            })
        st.session_state.hrp_next_leave_number = max_no + 1
        return records
    except Exception as e:
        print(f"HR leave data load failed: {e}")
        raise


def _hrp_save_leave_records(records=None, sync_drive=True):
    source_records = list(records if records is not None else st.session_state.get("hrp_leave_records", []))
    rows = []
    for r in source_records:
        rows.append({
            "Leave ID": r.get("leave_id", ""), "Employee ID": r.get("employee_id", ""),
            "Date From": r.get("date_from", ""), "Date To": r.get("date_to", ""),
            "Leave Type": r.get("type", ""), "Days": r.get("days", 0),
            "Status": r.get("status", "Approved"), "Request Source": r.get("request_source", "Email"),
            "Request Reference": r.get("request_reference", ""), "Notes": r.get("notes", ""),
            "Recorded By": r.get("recorded_by", ""), "Recorded At": r.get("recorded_at", ""),
            "Entry Source": r.get("entry_source", "HR Direct"), "Requested By": r.get("requested_by", ""),
            "Requested At": r.get("requested_at", ""), "Approved By": r.get("approved_by", ""),
            "Approved At": r.get("approved_at", ""), "Rejection Reason": r.get("rejection_reason", ""),
        })
    df = pd.DataFrame(rows, columns=HR_PORTAL_LEAVE_COLUMNS)
    parent = os.path.dirname(os.path.abspath(HR_PORTAL_LEAVE_PATH))
    tmp_path = os.path.join(parent, f".hr_portal_leave_records.save_{os.getpid()}_{threading.get_ident()}.xlsx")
    with _HR_PORTAL_LEAVE_LOCK:
        try:
            df.to_excel(tmp_path, index=False, engine="openpyxl")
            os.replace(tmp_path, HR_PORTAL_LEAVE_PATH)
            verify = pd.read_excel(HR_PORTAL_LEAVE_PATH, engine="openpyxl").fillna("")
            saved_ids = {str(x).strip() for x in verify.get("Leave ID", pd.Series(dtype=str)).tolist() if str(x).strip()}
            expected_ids = {str(r.get("leave_id", "")).strip() for r in source_records if str(r.get("leave_id", "")).strip()}
            if expected_ids - saved_ids:
                raise IOError(f"HR leave save verification failed; missing Leave ID(s): {sorted(expected_ids - saved_ids)}")
            st.session_state.hrp_leave_records = _hrp_load_leave_records()
        finally:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass
    if sync_drive:
        sync_saved_file_to_drive(HR_PORTAL_LEAVE_PATH)
    return True


def _hrp_enforce_ace_id_format(employees):
    employees = list(employees or [])
    used = set()
    next_no = 1
    for e in employees:
        eid = str(e.get("emp_id", "")).strip().upper()
        m = re.fullmatch(r"ACE-ID([0-9]+)", eid)
        if m:
            n = int(m.group(1))
            used.add(n)
            next_no = max(next_no, n + 1)
    changes = []
    for e in employees:
        old = str(e.get("emp_id", "")).strip()
        if re.fullmatch(r"ACE-ID[0-9]+", old.upper()):
            e["emp_id"] = old.upper()
            continue
        while next_no in used:
            next_no += 1
        new = f"ACE-ID{next_no:03d}"
        used.add(next_no)
        next_no += 1
        e["emp_id"] = new
        changes.append((old, new))
    if not changes:
        return employees

    users = load_users()
    for old, new in changes:
        for u in users.values():
            if str(u.get("employee_id", "")).strip().casefold() == old.casefold():
                u["employee_id"] = new
    save_users(users)

    hr_leave = load_hr_leave(force=True)
    for r in hr_leave:
        for old, new in changes:
            if str(r.get("employee_id", "")).strip().casefold() == old.casefold():
                r["employee_id"] = new
    save_all_hr_leave(hr_leave)

    portal_leave = _hrp_load_leave_records()
    for r in portal_leave:
        for old, new in changes:
            if str(r.get("employee_id", "")).strip().casefold() == old.casefold():
                r["employee_id"] = new
    st.session_state.hrp_leave_records = portal_leave
    _hrp_save_leave_records()
    return employees


def _hrp_reconcile_leaver_statuses():
    employees = st.session_state.get("hrp_employees", [])
    settlements = [
        r for r in load_hr_leave(force=False)
        if _hrp_is_final_holiday_settlement_record(r)
        and str(r.get("status", "")).strip().casefold() in {"pending", "approved"}
    ]
    changed = False
    for emp in employees:
        eid = str(emp.get("emp_id", "")).strip().casefold()
        matches = [r for r in settlements if str(r.get("employee_id", "")).strip().casefold() == eid]
        if not matches:
            continue
        if str(emp.get("status", "")).strip().casefold() != "left":
            emp["status"] = "Left"
            changed = True
        if not emp.get("leaving_date"):
            raw_date = matches[0].get("date", "")
            try:
                parsed = pd.to_datetime(raw_date, errors="coerce")
                if pd.notna(parsed):
                    emp["leaving_date"] = parsed.date()
                    changed = True
            except Exception:
                pass
        if not str(emp.get("leaving_reason", "")).strip():
            emp["leaving_reason"] = "Final holiday settlement"
            changed = True
    if changed:
        _hrp_save_employees()
    return changed


def _hr_portal_init():
    if "hrp_role" not in st.session_state:
        st.session_state.hrp_role = "hr"

    needs_reload = (
        not st.session_state.get("hrp_data_loaded", False)
        or bool(st.session_state.pop("hrp_force_reload", False))
    )

    if needs_reload:
        _hrp_init_storage()
        st.session_state.hrp_employees = _hrp_enforce_ace_id_format(_hrp_load_employees())
        st.session_state.hrp_leave_records = _hrp_load_leave_records()
        st.session_state.hrp_storage_version = 4
        _hrp_reconcile_leaver_statuses()
        st.session_state.hrp_data_loaded = True

    active = [e["emp_id"] for e in st.session_state.hrp_employees if e.get("status") == "Active"]
    current_emp_id = str(st.session_state.get("hrp_current_emp_id", "") or "").strip()
    active_lookup = {str(x).casefold() for x in active}
    if current_emp_id.casefold() not in active_lookup:
        st.session_state.hrp_current_emp_id = active[0] if active else ""
    if "hrp_entitlement_overrides" not in st.session_state:
        st.session_state.hrp_entitlement_overrides = {e["emp_id"]: e["entitlement_override"] for e in st.session_state.hrp_employees if e.get("entitlement_override") is not None}
    if "hrp_adjustment_notes" not in st.session_state:
        st.session_state.hrp_adjustment_notes = {e["emp_id"]: e.get("adjustment_note", "") for e in st.session_state.hrp_employees if e.get("adjustment_note")}
    if "hrp_next_leave_number" not in st.session_state:
        st.session_state.hrp_next_leave_number = 1
    if "hrp_leave_form_version" not in st.session_state:
        st.session_state.hrp_leave_form_version = 0


def _hrp_get_employee(employee_id):
    target = str(employee_id or "").strip().casefold()
    employees = st.session_state.get("hrp_employees", []) or []
    for emp in employees:
        if str(emp.get("emp_id", "")).strip().casefold() == target:
            return emp

    try:
        refreshed = _hrp_load_employees()
        if refreshed:
            st.session_state.hrp_employees = refreshed
            for emp in refreshed:
                if str(emp.get("emp_id", "")).strip().casefold() == target:
                    return emp
    except Exception as e:
        print(f"HR employee refresh failed while resolving {employee_id}: {e}")
    return None


def _hrp_employee_label(employee_id, include_status=False):
    emp = _hrp_get_employee(employee_id)
    if not emp:
        return f"Employee not found — {employee_id}"
    name = str(emp.get("name", "")).strip() or "Unnamed employee"
    label = f"{name} — {employee_id}"
    if include_status:
        label += f" — {emp.get('status', 'Active')}"
    return label


def _hrp_validate_employee_id(employee_id, exclude_id=None):
    employee_id = str(employee_id).strip().upper()
    if not employee_id:
        return False, "Employee ID is required. Use ACE-ID followed by the employee number, e.g. ACE-ID001."
    if not re.fullmatch(r"ACE-ID[0-9]+", employee_id):
        return False, "Employee ID must start with ACE-ID and then contain numbers only, e.g. ACE-ID001."
    for emp in st.session_state.get("hrp_employees", []):
        existing = str(emp.get("emp_id", "")).strip().upper()
        if exclude_id is not None and existing == str(exclude_id).strip().upper():
            continue
        if existing == employee_id:
            return False, "This Employee ID already exists."
    return True, employee_id


def _hrp_is_final_holiday_settlement_record(record):
    if not bool(record.get("final_holiday_settlement", False)):
        return False
    category = str(record.get("category", "")).strip().casefold()
    rtype = str(record.get("type", "")).strip().casefold()
    desc = str(record.get("desc", "")).strip().casefold()
    valid_category = category in {"unused holiday payout", "overused holiday deduction"}
    valid_type = rtype in {"addition", "deduction"}
    return valid_type and (valid_category or "final holiday settlement" in desc)


def _hrp_get_working_days(start_date, end_date):
    if end_date < start_date:
        return 0
    total = 0
    current = start_date
    while current <= end_date:
        if current.weekday() < 5 and current not in _hrp_non_working_dates(current.year):
            total += 1
        current += timedelta(days=1)
    return total


def _hrp_calculate_leave_days(leave_type, start_date, end_date, half_day=False):
    if end_date < start_date:
        return 0.0
    if leave_type == "Half Day Holiday" or half_day:
        return 0.5 if _hrp_get_working_days(start_date, end_date) > 0 else 0.0
    return float(_hrp_get_working_days(start_date, end_date))


def _hrp_calculate_service_years(start_date):
    days = (date.today() - start_date).days
    if days < 0:
        return 0.0
    return days / 365.25


def _hrp_calculate_holiday_entitlement(employee):
    start_date = employee["start_date"]
    today = date.today()
    leave_year_start_month, leave_year_start_day = 1, 1
    if today.month > leave_year_start_month or (today.month == leave_year_start_month and today.day >= leave_year_start_day):
        holiday_year_start = date(today.year, leave_year_start_month, leave_year_start_day)
    else:
        holiday_year_start = date(today.year - 1, leave_year_start_month, leave_year_start_day)
    holiday_year_end = date(holiday_year_start.year + 1, leave_year_start_month, leave_year_start_day) - timedelta(days=1)

    try:
        days_per_week = float(employee.get("days_per_week", 5) or 5)
    except Exception:
        days_per_week = 5.0
    days_per_week = max(0.0, min(days_per_week, 7.0))
    full_year_entitlement = min(28.0, 5.6 * days_per_week)
    service_years = _hrp_calculate_service_years(start_date)
    working_pattern = str(employee.get("working_pattern", "Regular hours") or "Regular hours")

    if working_pattern != "Regular hours":
        return full_year_entitlement, (
            "Irregular/part-year worker: GOV.UK requires leave to be accrued using the "
            "12.07% of hours worked in each pay period method for leave years beginning "
            "on or after 1 April 2024. Enter/pay-period hours are not stored in this portal yet, "
            "so no 12.07% hours calculation has been assumed here."
        ), service_years

    if start_date <= holiday_year_start:
        return round(full_year_entitlement, 1), (
            f"UK statutory minimum: {full_year_entitlement:.1f} days for {days_per_week:g} contracted days per week "
            f"(5.6 weeks, capped at 28 days). Leave year: {holiday_year_start:%d/%m/%Y} to {holiday_year_end:%d/%m/%Y}."
        ), service_years

    employed_days = (holiday_year_end - start_date).days + 1
    year_days = (holiday_year_end - holiday_year_start).days + 1
    raw = full_year_entitlement * max(0.0, min(employed_days / year_days, 1.0))
    entitlement = (int(raw * 2 + 0.999999) / 2.0)
    return entitlement, (
        f"Gross UK pro-rata: {full_year_entitlement:.1f} days full-year entitlement × "
        f"{employed_days}/{year_days} calendar days remaining in the leave year, "
        f"rounded up to the next half day. Bank holidays are separate non-working days and are not deducted from the allowance. "
        f"Leave year: {holiday_year_start:%d/%m/%Y} to {holiday_year_end:%d/%m/%Y}."
    ), service_years


def _hrp_get_employee_entitlement(employee):
    employee_id = employee["emp_id"]
    calculated, note, service_years = _hrp_calculate_holiday_entitlement(employee)
    entitlement = calculated
    if employee.get("entitlement_override") is not None:
        entitlement = float(employee["entitlement_override"])
        note = "HR-adjusted pure holiday entitlement; bank holidays are separate and are not deducted."
    else:
        entitlement_overrides = st.session_state.get("hrp_entitlement_overrides", {}) or {}
        if employee_id in entitlement_overrides:
            entitlement = float(entitlement_overrides[employee_id])
            note = "HR-adjusted pure holiday entitlement; bank holidays are separate and are not deducted."
    entitlement = max(0.0, round(entitlement, 1))
    note = f"{note} Bank holidays are not deducted from holiday entitlement."
    return entitlement, note, service_years, calculated


def _hrp_get_employee_leave(employee_id):
    records = st.session_state.get("hrp_leave_records")
    if records is None:
        records = _hrp_load_leave_records()
        st.session_state.hrp_leave_records = records
    return [r for r in records if str(r.get("employee_id", "")).strip().casefold() == str(employee_id).strip().casefold()]


def _hrp_get_approved_holiday_days(employee_id):
    employee = _hrp_get_employee(employee_id)
    records = st.session_state.get("hrp_leave_records")
    if records is None:
        records = _hrp_load_leave_records()
        st.session_state.hrp_leave_records = records
    recorded_days = sum(
        float(r.get("days", 0) or 0)
        for r in records
        if str(r.get("employee_id", "")).strip().casefold() == str(employee_id).strip().casefold()
        and str(r.get("status", "Approved")).strip().casefold() == "approved"
        and r.get("type") in HR_PORTAL_HOLIDAY_LEAVE_TYPES
    )
    return round(recorded_days, 1)


def _hrp_get_company_closure_holiday_days(employee, start_date, end_date):
    if not employee or not start_date or not end_date or end_date < start_date:
        return 0.0
    try:
        days_per_week = float(employee.get("days_per_week", 5) or 5)
    except Exception:
        days_per_week = 5.0
    factor = min(1.0, max(0.0, days_per_week / 5.0))
    total = 0.0
    for year in range(start_date.year, end_date.year + 1):
        for closure_day in sorted(_hrp_company_closure_dates(year)):
            if start_date <= closure_day <= end_date and closure_day.weekday() < 5 and closure_day not in HRP_BANK_HOLIDAYS.get(year, {}):
                total += factor
    return round(total, 1)


def _hrp_get_bank_holiday_days(employee, start_date, end_date):
    if not employee or not start_date or not end_date or end_date < start_date:
        return 0.0
    try:
        days_per_week = float(employee.get("days_per_week", 5) or 5)
    except Exception:
        days_per_week = 5.0
    factor = min(1.0, max(0.0, days_per_week / 5.0))
    total = 0.0
    for year in range(start_date.year, end_date.year + 1):
        for bank_day in HRP_BANK_HOLIDAYS.get(year, {}):
            if start_date <= bank_day <= end_date and bank_day.weekday() < 5:
                total += factor
    return round(total, 1)


def _hrp_get_leaving_entitlement(employee, leaving_date):
    if not employee or not leaving_date:
        return {"gross": 0.0, "bank_holidays": 0.0, "net": 0.0, "note": ""}
    start_date = employee.get("start_date")
    if not isinstance(start_date, date):
        try:
            start_date = pd.to_datetime(start_date).date()
        except Exception:
            start_date = leaving_date
    if leaving_date < start_date:
        return {"gross": 0.0, "bank_holidays": 0.0, "net": 0.0, "note": "Leaving date is before the employee start date."}
    try:
        days_per_week = float(employee.get("days_per_week", 5) or 5)
    except Exception:
        days_per_week = 5.0
    full_year = min(28.0, 5.6 * max(0.0, min(days_per_week, 7.0)))
    year_start = date(leaving_date.year, 1, 1)
    year_end = date(leaving_date.year, 12, 31)
    employed_start = max(start_date, year_start)
    employed_end = min(leaving_date, year_end)
    if employed_end < employed_start:
        gross = 0.0
    elif employed_start == year_start and employed_end == year_end:
        gross = full_year
    else:
        employed_days = (employed_end - employed_start).days + 1
        year_days = (year_end - year_start).days + 1
        gross = full_year * (employed_days / year_days)
        gross = int(gross * 2 + 0.999999) / 2.0
    if employee.get("entitlement_override") is not None:
        try:
            override = float(employee.get("entitlement_override"))
            if employed_start == year_start and employed_end == year_end:
                gross = override
            else:
                gross = int((override * ((employed_end - employed_start).days + 1) / ((year_end - year_start).days + 1)) * 2 + 0.999999) / 2.0
        except Exception:
            pass
    bank_days = _hrp_get_bank_holiday_days(employee, employed_start, employed_end)
    pure_entitlement = max(0.0, round(gross, 1))
    return {
        "gross": pure_entitlement,
        "bank_holidays": bank_days,
        "net": pure_entitlement,
        "note": f"Pure holiday entitlement to leaving date: {pure_entitlement:.1f} days. {bank_days:.1f} bank holiday day(s) occurred in the employment period; none are deducted.",
    }


def _hrp_get_approved_holiday_days_to_date(employee_id, end_date):
    employee = _hrp_get_employee(employee_id)
    if not employee or not end_date:
        return 0.0
    records = st.session_state.get("hrp_leave_records")
    if records is None:
        records = _hrp_load_leave_records()
        st.session_state.hrp_leave_records = records
    total = 0.0
    for r in records:
        if str(r.get("employee_id", "")).strip().casefold() != str(employee_id).strip().casefold():
            continue
        if str(r.get("status", "Approved")).strip().casefold() != "approved" or r.get("type") not in HR_PORTAL_HOLIDAY_LEAVE_TYPES:
            continue
        d_from, d_to = r.get("date_from"), r.get("date_to")
        if not isinstance(d_from, date) or not isinstance(d_to, date):
            continue
        if d_from > end_date:
            continue
        effective_to = min(d_to, end_date)
        if effective_to < d_from:
            continue
        total += _hrp_calculate_leave_days(r.get("type"), d_from, effective_to)
    return round(total, 1)


def _hrp_get_upcoming_bank_holidays(from_date=None, limit=5):
    start = from_date or date.today()
    upcoming = []
    for year in sorted(HRP_BANK_HOLIDAYS):
        for bank_day, name in sorted(HRP_BANK_HOLIDAYS.get(year, {}).items()):
            if bank_day >= start:
                upcoming.append((bank_day, name))
    return upcoming[:max(1, int(limit))]


def _hrp_get_upcoming_bank_holiday_days_for_employee(employee, from_date=None, end_date=None):
    if not employee:
        return 0.0
    start = from_date or date.today()
    start_date = employee.get("start_date")
    if isinstance(start_date, date):
        start = max(start, start_date)
    elif start_date:
        try:
            start = max(start, pd.to_datetime(start_date).date())
        except Exception:
            pass

    if end_date is None:
        leaving_date = employee.get("leaving_date")
        if isinstance(leaving_date, date):
            end_date = leaving_date
        elif leaving_date:
            try:
                end_date = pd.to_datetime(leaving_date).date()
            except Exception:
                end_date = date(start.year, 12, 31)
        else:
            end_date = date(start.year, 12, 31)
    if end_date < start:
        return 0.0

    try:
        days_per_week = float(employee.get("days_per_week", 5) or 5)
    except Exception:
        days_per_week = 5.0
    factor = min(1.0, max(0.0, days_per_week / 5.0))

    total = 0.0
    for year in range(start.year, end_date.year + 1):
        for bank_day in HRP_BANK_HOLIDAYS.get(year, {}):
            if start <= bank_day <= end_date and bank_day.weekday() < 5:
                total += factor
    return round(total, 1)


def _hrp_get_approved_final_settlement_offset(employee_id):
    try:
        records = load_hr_leave(force=False)
    except Exception:
        return 0.0
    total = 0.0
    target = str(employee_id or "").strip().casefold()
    for r in records:
        if str(r.get("status", "")).strip().casefold() != "approved":
            continue
        if not r.get("final_holiday_settlement", False):
            continue
        rid = str(r.get("employee_id", "")).strip().casefold()
        if rid != target:
            continue
        try:
            days = abs(float(r.get("days", 0) or 0))
        except Exception:
            days = 0.0
        if str(r.get("type", "")).strip().casefold() == "addition":
            total += days
        elif str(r.get("type", "")).strip().casefold() == "deduction":
            total -= days
    return round(total, 1)


def _hrp_get_final_settlement_records(employee_id):
    target = str(employee_id or "").strip().casefold()
    try:
        records = load_hr_leave(force=False)
    except Exception:
        return []
    out = []
    for r in records:
        if not r.get("final_holiday_settlement", False):
            continue
        if str(r.get("employee_id", "")).strip().casefold() != target:
            continue
        out.append(r)
    out.sort(key=lambda r: int(r.get("id", 0) or 0), reverse=True)
    return out


def _hrp_has_active_final_settlement(employee_id):
    return any(
        str(r.get("status", "")).strip().casefold() in {"pending", "approved"}
        for r in _hrp_get_final_settlement_records(employee_id)
    )


def _hrp_get_holiday_position(employee_id):
    employee = _hrp_get_employee(employee_id)
    if not employee:
        return {
            "entitlement": 0.0, "used": 0.0, "bank_holidays": 0.0,
            "balance": 0.0, "employee_owes_company": 0.0, "company_owes_employee": 0.0,
        }
    if str(employee.get("status", "")).strip().casefold() == "left" or employee.get("leaving_date"):
        return {
            "entitlement": 0.0, "used": 0.0, "recorded_used": 0.0,
            "company_closure": 0.0, "bank_holidays": 0.0, "raw_balance": 0.0,
            "final_settlement_offset": 0.0, "balance": 0.0,
            "employee_owes_company": 0.0, "company_owes_employee": 0.0,
        }

    entitlement, _, _, _ = _hrp_get_employee_entitlement(employee)
    recorded_used = _hrp_get_approved_holiday_days(employee_id)

    today = date.today()
    start_date = employee.get("start_date")
    if not isinstance(start_date, date):
        try:
            start_date = pd.to_datetime(start_date).date()
        except Exception:
            start_date = today
    end_date = date(today.year, 12, 31)
    leaving_date = employee.get("leaving_date")
    if leaving_date:
        try:
            if not isinstance(leaving_date, date):
                leaving_date = pd.to_datetime(leaving_date).date()
            end_date = min(end_date, leaving_date)
        except Exception:
            pass
    bank_holidays_passed = _hrp_get_bank_holiday_days(employee, start_date, min(today, end_date))
    upcoming_bank_holidays = _hrp_get_upcoming_bank_holiday_days_for_employee(
        employee, today + timedelta(days=1), end_date
    )
    bank_holidays = round(bank_holidays_passed + upcoming_bank_holidays, 1)
    company_closure = _hrp_get_company_closure_holiday_days(employee, start_date, end_date)
    used = round(recorded_used + company_closure, 1)
    raw_balance = round(entitlement - used - bank_holidays, 1)
    final_settlement_offset = _hrp_get_approved_final_settlement_offset(employee_id)
    balance = round(raw_balance - final_settlement_offset, 1)
    return {
        "entitlement": entitlement,
        "used": used,
        "recorded_used": recorded_used,
        "company_closure": company_closure,
        "bank_holidays": bank_holidays,
        "bank_holidays_passed": round(bank_holidays_passed, 1),
        "upcoming_bank_holidays": round(upcoming_bank_holidays, 1),
        "raw_balance": raw_balance,
        "final_settlement_offset": final_settlement_offset,
        "balance": balance,
        "employee_owes_company": round(abs(balance), 1) if balance < 0 else 0.0,
        "company_owes_employee": round(balance, 1) if balance > 0 else 0.0,
    }


def _hrp_get_leave_summary(employee_id):
    records = _hrp_get_employee_leave(employee_id)
    return {
        "holiday": _hrp_get_approved_holiday_days(employee_id),
        "sick": sum(r["days"] for r in records if r["status"] == "Approved" and r["type"] == "Sick Leave"),
        "family": sum(r["days"] for r in records if r["status"] == "Approved" and r["type"] == "Family / Emergency Leave"),
        "other": sum(r["days"] for r in records if r["status"] == "Approved" and r["type"] == "Other Absence"),
        "unpaid": sum(r["days"] for r in records if r["status"] == "Approved" and r["type"] in ["Unpaid Holiday", "Unpaid Absence"]),
        "maternity": sum(r["days"] for r in records if r["status"] == "Approved" and r["type"] == "Maternity Leave"),
        "paternity": sum(r["days"] for r in records if r["status"] == "Approved" and r["type"] == "Paternity Leave"),
    }


def _hrp_create_leave_id():
    leave_id = f"LV-{st.session_state.hrp_next_leave_number:03d}"
    st.session_state.hrp_next_leave_number += 1
    return leave_id


# ════════════════════════════════════════════════════════════
# 📅 HR HOLIDAY CALENDAR — Excel-style yearly employee calendar
# ════════════════════════════════════════════════════════════
HRP_BANK_HOLIDAYS = {
    2026: {
        date(2026, 1, 1): "New Year's Day",
        date(2026, 4, 3): "Good Friday",
        date(2026, 4, 6): "Easter Monday",
        date(2026, 5, 4): "Early May bank holiday",
        date(2026, 5, 25): "Spring bank holiday",
        date(2026, 8, 31): "Summer bank holiday",
        date(2026, 12, 25): "Christmas Day",
        date(2026, 12, 28): "Boxing Day (substitute day)",
    },
    2027: {
        date(2027, 1, 1): "New Year's Day",
        date(2027, 3, 26): "Good Friday",
        date(2027, 3, 29): "Easter Monday",
        date(2027, 5, 3): "Early May bank holiday",
        date(2027, 5, 31): "Spring bank holiday",
        date(2027, 8, 30): "Summer bank holiday",
        date(2027, 12, 27): "Christmas Day (substitute day)",
        date(2027, 12, 28): "Boxing Day (substitute day)",
    },
    2028: {
        date(2028, 1, 3): "New Year's Day (substitute day)",
        date(2028, 4, 14): "Good Friday",
        date(2028, 4, 17): "Easter Monday",
        date(2028, 5, 1): "Early May bank holiday",
        date(2028, 5, 29): "Spring bank holiday",
        date(2028, 8, 28): "Summer bank holiday",
        date(2028, 12, 25): "Christmas Day",
        date(2028, 12, 26): "Boxing Day",
    },
}


def _hrp_company_closure_dates(year):
    return {date(year, 12, day) for day in (29, 30, 31)}


def _hrp_non_working_dates(year):
    dates = set(HRP_BANK_HOLIDAYS.get(int(year), {}).keys())
    dates.update(_hrp_company_closure_dates(int(year)))
    return dates


def _hrp_non_working_reason(day):
    if day in HRP_BANK_HOLIDAYS.get(day.year, {}):
        return "BH", HRP_BANK_HOLIDAYS[day.year][day]
    if day in _hrp_company_closure_dates(day.year):
        return "H", "Company Christmas closure"
    if day.weekday() >= 5:
        return "NA", "Weekend / non-working day"
    return "", ""


def _hrp_blocked_leave_dates(start_date, end_date):
    if end_date < start_date:
        return []
    blocked = []
    current = start_date
    while current <= end_date:
        if current in _hrp_non_working_dates(current.year):
            code, reason = _hrp_non_working_reason(current)
            blocked.append((current, code, reason))
        current += timedelta(days=1)
    return blocked


HRP_CALENDAR_CODES = {
    "Full Day Holiday": "H",
    "Half Day Holiday": "HD",
    "Bank Holiday": "BH",
    "Sick Leave": "S",
    "Family / Emergency Leave": "FE",
    "Unpaid Holiday": "UH",
    "Unpaid Absence": "UA",
    "Maternity Leave": "M",
    "Paternity Leave": "P",
    "Other Absence": "O",
    "College (Apprenticeship)": "C",
    "Training Course": "T",
}

HRP_CALENDAR_COLOURS = {
    "H": "#92D050",
    "HD": "#FFC000",
    "BH": "#E76153",
    "UH": "#22989E",
    "C": "#D86DCD",
    "NA": "#B793FF",
    "T": "#FFFF00",
    "M": "#B8D3EF",
    "UA": "#00FFFF",
    "S": "#AFABAB",
    "LEFT": "#FF0000",
}


def _hrp_calendar_leave_code(leave_type, status):
    code = HRP_CALENDAR_CODES.get(str(leave_type or "").strip(), "L")
    status = str(status or "Approved").strip().casefold()
    if status == "pending":
        return f"{code}*"
    if status == "rejected":
        return ""
    return code


def _hrp_render_holiday_calendar():
    st.subheader("📅 Holiday & Absence Calendar")
    st.caption(
        "Excel-style yearly calendar showing every employee, department, start date and approved leave. "
        "Bank holidays are shown as BH and 29–31 December company closure is shown as H. "
        "Pending department-manager requests are marked with * and are not treated as approved leave."
    )

    employees = _hrp_load_employees()
    leave_records = _hrp_load_leave_records()
    st.session_state.hrp_employees = employees
    st.session_state.hrp_leave_records = leave_records

    c1, c2, c3 = st.columns([1, 2, 2])
    with c1:
        calendar_year = st.number_input(
            "Calendar Year", min_value=2020, max_value=2100,
            value=date.today().year, step=1, key="hrp_calendar_year"
        )
    with c2:
        departments = sorted({str(e.get("department", "")).strip() for e in employees if str(e.get("department", "")).strip()})
        department_filter = st.selectbox(
            "Department", ["All Departments"] + departments,
            key="hrp_calendar_department"
        )
    with c3:
        status_filter = st.selectbox(
            "Leave shown", ["Approved + Pending", "Approved only"],
            key="hrp_calendar_status"
        )

    filtered_employees = [
        e for e in employees
        if department_filter == "All Departments"
        or str(e.get("department", "")).strip().casefold() == department_filter.casefold()
    ]
    filtered_employees.sort(key=lambda e: (str(e.get("department", "")).casefold(), str(e.get("name", "")).casefold()))

    leave_lookup = {}
    non_working_dates = _hrp_non_working_dates(int(calendar_year))
    for record in leave_records:
        status = str(record.get("status", "Approved")).strip()
        if status.casefold() == "rejected":
            continue
        if status_filter == "Approved only" and status.casefold() != "approved":
            continue
        try:
            start = record["date_from"]
            end = record["date_to"]
            if isinstance(start, str):
                start = pd.to_datetime(start).date()
            if isinstance(end, str):
                end = pd.to_datetime(end).date()
        except Exception:
            continue
        if end < start:
            start, end = end, start
        code = _hrp_calendar_leave_code(record.get("type", ""), status)
        if not code:
            continue
        employee_id = str(record.get("employee_id", "")).strip().casefold()
        current = start
        while current <= end:
            if current.year == int(calendar_year) and current.weekday() < 5:
                key = (employee_id.casefold(), current)
                if current in non_working_dates:
                    current += timedelta(days=1)
                    continue
                existing = leave_lookup.get(key, "")
                if code.rstrip("*") == "S":
                    leave_lookup[key] = code
                elif existing.rstrip("*") == "S":
                    pass
                elif existing == "":
                    leave_lookup[key] = code
                elif existing.endswith("*") and not code.endswith("*"):
                    leave_lookup[key] = code
                elif code not in existing.split("/"):
                    leave_lookup[key] = f"{existing}/{code}"
            current += timedelta(days=1)

    first_day = date(int(calendar_year), 1, 1)
    last_day = date(int(calendar_year), 12, 31)
    dates = []
    current = first_day
    while current <= last_day:
        dates.append(current)
        current += timedelta(days=1)

    non_working_dates = _hrp_non_working_dates(int(calendar_year))
    for employee in filtered_employees:
        employee_id = str(employee.get("emp_id", "")).strip().casefold()
        for d in dates:
            code, _reason = _hrp_non_working_reason(d)
            if code:
                leave_lookup[(employee_id, d)] = code

    columns = ["Employee Name", "Department", "Start Date"] + [d.strftime("%d %b") for d in dates]
    rows = []
    for employee in filtered_employees:
        row = {
            "Employee Name": employee.get("name", ""),
            "Department": employee.get("department", ""),
            "Start Date": employee.get("start_date", ""),
        }
        raw_start = employee.get("start_date", "")
        try:
            if raw_start is None or pd.isna(raw_start) or not str(raw_start).strip():
                employee_start = None
            elif isinstance(raw_start, date) and not isinstance(raw_start, pd.Timestamp):
                employee_start = raw_start
            else:
                parsed_start = pd.to_datetime(raw_start, errors="coerce")
                employee_start = None if pd.isna(parsed_start) else parsed_start.date()
        except Exception:
            employee_start = None

        raw_leaving = employee.get("leaving_date", "")
        try:
            if raw_leaving is None or pd.isna(raw_leaving) or not str(raw_leaving).strip():
                employee_leaving = None
            elif isinstance(raw_leaving, date) and not isinstance(raw_leaving, pd.Timestamp):
                employee_leaving = raw_leaving
            else:
                parsed_leaving = pd.to_datetime(raw_leaving, errors="coerce")
                employee_leaving = None if pd.isna(parsed_leaving) else parsed_leaving.date()
        except Exception:
            employee_leaving = None
        for d in dates:
            if employee_start and d < employee_start:
                row[d.strftime("%d %b")] = "NA"
            elif employee_leaving and d > employee_leaving:
                row[d.strftime("%d %b")] = "LEFT"
            else:
                row[d.strftime("%d %b")] = leave_lookup.get((str(employee.get("emp_id", "")).strip().casefold(), d), "")
        rows.append(row)

    df = pd.DataFrame(rows, columns=columns)
    if not df.empty:
        df["Start Date"] = pd.to_datetime(df["Start Date"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")

    date_columns = [d.strftime("%d %b") for d in dates]

    def _style_calendar(dataframe):
        styles = pd.DataFrame("", index=dataframe.index, columns=dataframe.columns)
        for row_idx in dataframe.index:
            for d, col in zip(dates, date_columns):
                value = str(dataframe.at[row_idx, col] or "").strip()
                base_codes = [part.rstrip("*").strip() for part in value.split("/") if part.strip()]
                base_code = base_codes[0] if base_codes else ""
                bg = HRP_CALENDAR_COLOURS.get(base_code)
                if bg:
                    styles.at[row_idx, col] = f"background-color: {bg}; color: #000000; font-weight: 800; text-align: center; vertical-align: middle;"
                elif value.endswith("*"):
                    styles.at[row_idx, col] = "font-weight: 700; text-align: center; vertical-align: middle;"
                elif value:
                    styles.at[row_idx, col] = "font-weight: 700; text-align: center; vertical-align: middle;"
        return styles

    st.markdown(
        "**Legend:** H = Holiday · HD = Half Day Holiday · BH = Bank Holiday · UH = Unpaid Holiday · "
        "C = College (Apprenticeship) · NA = Closed (Weekend) / Not yet started · T = Training Course · "
        "M = Maternity Leave · UA = Unpaid Absence · S = Sick · FE = Family/Emergency · "
        "P = Paternity · O = Other Absence · * = Pending approval"
    )

    bank_days = HRP_BANK_HOLIDAYS.get(int(calendar_year), {})
    if bank_days:
        st.info("**Bank holidays / blocked dates:** " + ", ".join(f"{d.strftime('%d %b')} — {name}" for d, name in sorted(bank_days.items())))
    st.info(f"**Company closure:** 29, 30 and 31 December are company-closed days and are shown as **H**. They are treated as pre-booked annual holiday and reduce the available holiday balance; no separate leave entry is required.")
    st.caption(f"{len(filtered_employees)} employees · {len(dates)} calendar days · {calendar_year}")

    if df.empty:
        st.info("No employees match the selected department.")
    else:
        styled = (
            df.style
            .apply(_style_calendar, axis=None)
            .set_properties(**{"text-align": "center", "vertical-align": "middle"})
            .set_table_styles([
                {"selector": "th", "props": [("text-align", "center"), ("vertical-align", "middle")]},
                {"selector": "td", "props": [("text-align", "center"), ("vertical-align", "middle")]},
            ])
        )
        column_config = {
            "Employee Name": st.column_config.TextColumn("Employee Name", width="medium"),
            "Department": st.column_config.TextColumn("Department", width="medium"),
            "Start Date": st.column_config.TextColumn("Start Date", width="small"),
        }
        for col in date_columns:
            column_config[col] = st.column_config.TextColumn(col, width="small")
        st.dataframe(
            styled,
            width="stretch",
            hide_index=True,
            height=650,
            column_config=column_config,
        )


def _hrp_work_duration(start_date, end_date):
    try:
        start = pd.to_datetime(start_date).date() if not isinstance(start_date, date) else start_date
        end = pd.to_datetime(end_date).date() if not isinstance(end_date, date) else end_date
        if end < start:
            return "0 days"
        years = end.year - start.year
        months = end.month - start.month
        days = end.day - start.day
        if days < 0:
            months -= 1
            prev_month = end.month - 1 or 12
            prev_year = end.year if end.month > 1 else end.year - 1
            import calendar as _calendar
            days += _calendar.monthrange(prev_year, prev_month)[1]
        if months < 0:
            years -= 1
            months += 12
        parts = []
        if years:
            parts.append(f"{years} year" + ("s" if years != 1 else ""))
        if months:
            parts.append(f"{months} month" + ("s" if months != 1 else ""))
        if days or not parts:
            parts.append(f"{days} day" + ("s" if days != 1 else ""))
        return ", ".join(parts)
    except Exception:
        return "-"


def _hrp_leaver_history(employee):
    emp_id = str(employee.get("emp_id", "")).strip()
    leaving_date = employee.get("leaving_date")
    if not leaving_date:
        return {"leave_records": [], "settlements": [], "holiday_taken": 0.0, "types": {}}
    try:
        leaving_date = pd.to_datetime(leaving_date).date()
    except Exception:
        pass
    leave_records = []
    types = {}
    holiday_taken = 0.0
    for r in _hrp_load_leave_records():
        if str(r.get("employee_id", "")).strip().casefold() != emp_id.casefold():
            continue
        if str(r.get("status", "")).strip().casefold() != "approved":
            continue
        try:
            d_from = pd.to_datetime(r.get("date_from")).date()
        except Exception:
            d_from = None
        if d_from and isinstance(leaving_date, date) and d_from > leaving_date:
            continue
        days = float(r.get("days", 0) or 0)
        leave_type = str(r.get("type", r.get("leave_type", "Other")) or "Other").strip()
        types[leave_type] = round(types.get(leave_type, 0.0) + days, 1)
        if leave_type.casefold() in {"holiday", "annual holiday", "annual leave", "holiday leave"}:
            holiday_taken += days
        leave_records.append(r)
    settlements = [
        r for r in load_hr_leave(force=False)
        if r.get("final_holiday_settlement", False)
        and str(r.get("employee_id", "")).strip().casefold() == emp_id.casefold()
    ]
    return {"leave_records": leave_records, "settlements": settlements, "holiday_taken": round(holiday_taken, 1), "types": types}


# ============================================================
# 👥 HR LEAVE SETTLEMENT — DATA LAYER
# ============================================================
def initialise_hr_leave():
    safe_init_excel(HR_LEAVE_PATH, HR_LEAVE_COLUMNS)
    safe_init_excel(HR_DAILY_RATES_PATH, HR_DAILY_RATES_COLUMNS)


def load_hr_leave(force=False):
    if not force and "_hr_leave_cache" in st.session_state:
        return list(st.session_state["_hr_leave_cache"])
    initialise_hr_leave()
    try:
        df = _read_excel_records(HR_LEAVE_PATH)
        records = []
        for r in df.to_dict(orient="records"):
            try:
                rid = int(r.get("ID", 0))
            except Exception:
                rid = 0
            try:
                amount = float(r.get("Amount (£)", 0) or 0)
            except Exception:
                amount = 0.0
            try:
                days = float(r.get("Number of Days", 0) or 0)
            except Exception:
                days = 0.0
            records.append({
                "id": rid,
                "employee_id": str(r.get("Employee ID", "")).strip(),
                "emp_name": str(r.get("Employee Name", "")).strip(),
                "emp_dept": str(r.get("Employee Department", "")).strip(),
                "type": str(r.get("Transaction Type", "")).strip(),
                "category": str(r.get("Category Reason", "")).strip(),
                "owe_owed": str(r.get("Owe Owed", "")).strip(),
                "date": str(r.get("Date", "")).strip(),
                "days": days, "amount": amount,
                "manager": str(r.get("Line Manager", "")).strip(),
                "desc": str(r.get("Description", "")).strip(),
                "attachment_name": str(r.get("Attachment Name", "None")).strip(),
                "status": str(r.get("Status", "pending")).strip().lower(),
                "director_comments": str(r.get("Director Comments", "")).strip(),
                "rejection_reason": str(r.get("Rejection Reason", "")).strip(),
                "decision_date": str(r.get("Decision Date", "")).strip(),
                "decision_by": str(r.get("Decision By", "")).strip(),
                "submitted_by": str(r.get("Submitted By", "")).strip(),
                "submitted_date": str(r.get("Submitted Date", "")).strip(),
                "pdf_path": str(r.get("PDF File Path", "")).strip(),
                "final_holiday_settlement": str(r.get("Final Holiday Settlement", "")).strip().casefold() in {"yes", "true", "1"},
            })
        _set_data_cache("_hr_leave_cache", records)
        return list(records)
    except Exception as e:
        st.error(f"HR Leave Load Error: {e}")
        return []


def save_all_hr_leave(records, sync=True):
    rows = [{
        "ID": int(r.get("id", 0)),
        "Employee ID": str(r.get("employee_id", "")),
        "Employee Name": str(r.get("emp_name", "")),
        "Employee Department": str(r.get("emp_dept", "")),
        "Transaction Type": str(r.get("type", "")),
        "Category Reason": str(r.get("category", "")),
        "Owe Owed": str(r.get("owe_owed", "")),
        "Date": str(r.get("date", "")),
        "Number of Days": float(r.get("days", 0)),
        "Amount (£)": float(r.get("amount", 0)),
        "Line Manager": str(r.get("manager", "")),
        "Description": str(r.get("desc", "")),
        "Attachment Name": str(r.get("attachment_name", "None")),
        "Status": str(r.get("status", "pending")).lower(),
        "Director Comments": str(r.get("director_comments", "")),
        "Rejection Reason": str(r.get("rejection_reason", "")),
        "Decision Date": str(r.get("decision_date", "")),
        "Decision By": str(r.get("decision_by", "")),
        "Submitted By": str(r.get("submitted_by", "")),
        "Submitted Date": str(r.get("submitted_date", "")),
        "PDF File Path": str(r.get("pdf_path", "")),
        "Final Holiday Settlement": "Yes" if r.get("final_holiday_settlement", False) else "No",
    } for r in records]
    pd.DataFrame(rows, columns=HR_LEAVE_COLUMNS).to_excel(HR_LEAVE_PATH, index=False, engine="openpyxl")
    _set_data_cache("_hr_leave_cache", list(records))
    if sync:
        sync_saved_file_to_drive(HR_LEAVE_PATH)


def get_next_hr_leave_id(records):
    if not records:
        return 1
    return max(int(r.get("id", 0)) for r in records) + 1


def load_hr_daily_rates(force=False):
    if not force and "_hr_daily_rates_cache" in st.session_state:
        return list(st.session_state["_hr_daily_rates_cache"])
    initialise_hr_leave()
    try:
        df = _read_excel_records(HR_DAILY_RATES_PATH)
        records = []
        for r in df.to_dict(orient="records"):
            try:
                rate = float(r.get("Daily Rate (£)", 0) or 0)
            except Exception:
                rate = 0.0
            active = str(r.get("Active", "True")).strip().lower() in ("true", "yes", "1")
            records.append({
                "dept": str(r.get("Department", "")).strip(),
                "type": str(r.get("Transaction Type", "")).strip(),
                "rate": rate, "active": active,
            })
        _set_data_cache("_hr_daily_rates_cache", records)
        return list(records)
    except Exception as e:
        st.error(f"HR Daily Rates Load Error: {e}")
        return []


def save_hr_daily_rates(records):
    rows = [{"Department": r["dept"], "Transaction Type": r["type"],
             "Daily Rate (£)": float(r["rate"]), "Active": bool(r["active"])} for r in records]
    pd.DataFrame(rows, columns=HR_DAILY_RATES_COLUMNS).to_excel(HR_DAILY_RATES_PATH, index=False, engine="openpyxl")
    _set_data_cache("_hr_daily_rates_cache", list(records))
    sync_saved_file_to_drive(HR_DAILY_RATES_PATH)


def get_hr_daily_rate(dept, transaction_type):
    for r in load_hr_daily_rates():
        if r["dept"] == dept and r["type"] == transaction_type and r["active"]:
            return r["rate"]
    return None


# ============================================================
# 📦 STORE DEPARTMENT DEDUCTION — DATA LAYER
# ============================================================
def initialise_store_deduction():
    safe_init_excel(STORE_DEDUCTION_PATH, STORE_DEDUCTION_COLUMNS)
    if not os.path.exists(STORE_ITEMS_PATH) or os.path.getsize(STORE_ITEMS_PATH) == 0:
        try:
            pd.DataFrame(DEFAULT_STORE_ITEMS).to_excel(STORE_ITEMS_PATH, index=False, engine="openpyxl")
        except Exception as e:
            print(f"Store items initialisation failed: {e}")


def load_store_deductions(force=False):
    if not force and "_store_deduction_cache" in st.session_state:
        return list(st.session_state["_store_deduction_cache"])
    initialise_store_deduction()
    try:
        df = _read_excel_records(STORE_DEDUCTION_PATH)
        records = []
        for r in df.to_dict(orient="records"):
            try:
                rid = int(r.get("ID", 0))
            except Exception:
                rid = 0
            try:
                total = float(r.get("Total Deduction (£)", 0) or 0)
            except Exception:
                total = 0.0
            try:
                items_json = r.get("Items Deducted JSON", "[]")
            except Exception:
                items_json = "[]"
            try:
                items = json.loads(items_json) if items_json else []
            except Exception:
                items = []
            records.append({
                "id": rid,
                "emp_name": str(r.get("Employee Name", "")).strip(),
                "date_leaving": str(r.get("Date of Leaving", "")).strip(),
                "emp_dept": str(r.get("Employee Department", "")).strip(),
                "manager": str(r.get("Line Manager", "")).strip(),
                "date_submit": str(r.get("Date of Submit", "")).strip(),
                "type": str(r.get("Type", "Deduction")).strip(),
                "items": items,
                "total_deduction": total,
                "desc": str(r.get("Description", "")).strip(),
                "attachment_name": str(r.get("Attachment Name", "None")).strip(),
                "status": str(r.get("Status", "pending")).strip().lower(),
                "director_comments": str(r.get("Director Comments", "")).strip(),
                "rejection_reason": str(r.get("Rejection Reason", "")).strip(),
                "decision_date": str(r.get("Decision Date", "")).strip(),
                "decision_by": str(r.get("Decision By", "")).strip(),
                "submitted_by": str(r.get("Submitted By", "")).strip(),
                "submitted_date": str(r.get("Submitted Date", "")).strip(),
                "pdf_path": str(r.get("PDF File Path", "")).strip(),
            })
        _set_data_cache("_store_deduction_cache", records)
        return list(records)
    except Exception as e:
        st.error(f"Store Deduction Load Error: {e}")
        return []


def save_all_store_deductions(records, sync=True):
    rows = [{
        "ID": int(r.get("id", 0)),
        "Employee Name": str(r.get("emp_name", "")),
        "Date of Leaving": str(r.get("date_leaving", "")),
        "Employee Department": str(r.get("emp_dept", "")),
        "Line Manager": str(r.get("manager", "")),
        "Date of Submit": str(r.get("date_submit", "")),
        "Type": str(r.get("type", "Deduction")),
        "Items Deducted JSON": json.dumps(r.get("items", []), ensure_ascii=False),
        "Total Deduction (£)": float(r.get("total_deduction", 0)),
        "Description": str(r.get("desc", "")),
        "Attachment Name": str(r.get("attachment_name", "None")),
        "Status": str(r.get("status", "pending")).lower(),
        "Director Comments": str(r.get("director_comments", "")),
        "Rejection Reason": str(r.get("rejection_reason", "")),
        "Decision Date": str(r.get("decision_date", "")),
        "Decision By": str(r.get("decision_by", "")),
        "Submitted By": str(r.get("submitted_by", "")),
        "Submitted Date": str(r.get("submitted_date", "")),
        "PDF File Path": str(r.get("pdf_path", "")),
    } for r in records]
    pd.DataFrame(rows, columns=STORE_DEDUCTION_COLUMNS).to_excel(STORE_DEDUCTION_PATH, index=False, engine="openpyxl")
    _set_data_cache("_store_deduction_cache", list(records))
    if sync:
        sync_saved_file_to_drive(STORE_DEDUCTION_PATH)


def get_next_store_deduction_id(records):
    if not records:
        return 1
    return max(int(r.get("id", 0)) for r in records) + 1


def load_store_items(force=False):
    if not force and "_store_items_cache" in st.session_state:
        return list(st.session_state["_store_items_cache"])
    initialise_store_deduction()
    try:
        df = _read_excel_records(STORE_ITEMS_PATH)
        records = []
        for r in df.to_dict(orient="records"):
            try:
                price = float(r.get("Price (£)", 0) or 0)
            except Exception:
                price = 0.0
            active = str(r.get("Active", "True")).strip().lower() in ("true", "yes", "1")
            records.append({
                "name": str(r.get("Item Name", "")).strip(),
                "price": price,
                "active": active,
            })
        _set_data_cache("_store_items_cache", records)
        return list(records)
    except Exception as e:
        st.error(f"Store Items Load Error: {e}")
        return []


def save_store_items(records):
    rows = [{"Item Name": r["name"], "Price (£)": float(r["price"]), "Active": bool(r["active"])} for r in records]
    pd.DataFrame(rows, columns=STORE_ITEMS_COLUMNS).to_excel(STORE_ITEMS_PATH, index=False, engine="openpyxl")
    _set_data_cache("_store_items_cache", list(records))
    sync_saved_file_to_drive(STORE_ITEMS_PATH)


# ============================================================
# 🔗 EMPLOYEE LEAVER CLEARANCE — COMPANY-WIDE WORKFLOW
# ============================================================
def initialise_leaver_clearance():
    safe_init_excel(LEAVER_CLEARANCE_PATH, LEAVER_CLEARANCE_COLUMNS)


def _normalise_clearance_record(r):
    def _money(v):
        try:
            return float(v or 0)
        except Exception:
            return 0.0
    return {
        "clearance_id": str(r.get("Clearance ID", r.get("clearance_id", ""))).strip(),
        "employee_id": str(r.get("Employee ID", r.get("employee_id", ""))).strip(),
        "emp_name": str(r.get("Employee Name", r.get("emp_name", ""))).strip(),
        "emp_dept": str(r.get("Department", r.get("emp_dept", ""))).strip(),
        "leaving_date": str(r.get("Leaving Date", r.get("leaving_date", ""))).strip(),
        "leaving_reason": str(r.get("Leaving Reason", r.get("leaving_reason", ""))).strip(),
        "created_by": str(r.get("Created By", r.get("created_by", ""))).strip(),
        "created_at": str(r.get("Created At", r.get("created_at", ""))).strip(),
        "hr_status": str(r.get("HR Status", r.get("hr_status", "Pending"))).strip() or "Pending",
        "holiday_balance": _money(r.get("Holiday Balance (Days)", r.get("holiday_balance", 0))),
        "holiday_settlement_id": str(r.get("Holiday Settlement ID", r.get("holiday_settlement_id", ""))).strip(),
        "holiday_settlement_status": str(r.get("Holiday Settlement Status", r.get("holiday_settlement_status", "Not Required"))).strip() or "Not Required",
        "holiday_settlement_amount": _money(r.get("Holiday Settlement Amount (£)", r.get("holiday_settlement_amount", 0))),
        "holiday_settlement_type": str(r.get("Holiday Settlement Type", r.get("holiday_settlement_type", ""))).strip(),
        "store_status": str(r.get("Store Status", r.get("store_status", "Pending"))).strip() or "Pending",
        "store_checkin_id": str(r.get("Store Check-in ID", r.get("store_checkin_id", ""))).strip(),
        "store_deduction_id": str(r.get("Store Deduction ID", r.get("store_deduction_id", ""))).strip(),
        "store_return_id": str(r.get("Store Return ID", r.get("store_return_id", ""))).strip(),
        "outstanding_item_value": _money(r.get("Outstanding Item Value (£)", r.get("outstanding_item_value", 0))),
        "store_completed_by": str(r.get("Store Completed By", r.get("store_completed_by", ""))).strip(),
        "store_completed_at": str(r.get("Store Completed At", r.get("store_completed_at", ""))).strip(),
        "payroll_status": str(r.get("Payroll Status", r.get("payroll_status", "Pending"))).strip() or "Pending",
        "payroll_request_ids": str(r.get("Payroll Request IDs", r.get("payroll_request_ids", ""))).strip(),
        "payroll_total_addition": _money(r.get("Payroll Total Addition (£)", r.get("payroll_total_addition", 0))),
        "payroll_total_deduction": _money(r.get("Payroll Total Deduction (£)", r.get("payroll_total_deduction", 0))),
        "payroll_completed_by": str(r.get("Payroll Completed By", r.get("payroll_completed_by", ""))).strip(),
        "payroll_completed_at": str(r.get("Payroll Completed At", r.get("payroll_completed_at", ""))).strip(),
        "final_status": str(r.get("Final Status", r.get("final_status", "Open"))).strip() or "Open",
        "final_cleared_by": str(r.get("Final Cleared By", r.get("final_cleared_by", ""))).strip(),
        "final_cleared_at": str(r.get("Final Cleared At", r.get("final_cleared_at", ""))).strip(),
        "notes": str(r.get("Notes", r.get("notes", ""))).strip(),
    }


def load_leaver_clearances(force=False):
    if not force and "_leaver_clearance_cache" in st.session_state:
        return list(st.session_state["_leaver_clearance_cache"])
    initialise_leaver_clearance()
    try:
        df = _read_excel_records(LEAVER_CLEARANCE_PATH)
        records = [_normalise_clearance_record(r) for r in df.to_dict(orient="records")]
        records = [r for r in records if r.get("clearance_id") and r.get("employee_id")]
        _set_data_cache("_leaver_clearance_cache", records)
        return list(records)
    except Exception as e:
        st.error(f"Leaver Clearance Load Error: {e}")
        return []


def save_all_leaver_clearances(records, sync=True):
    rows = []
    for r in records:
        rows.append({
            "Clearance ID": str(r.get("clearance_id", "")), "Employee ID": str(r.get("employee_id", "")),
            "Employee Name": str(r.get("emp_name", "")), "Department": str(r.get("emp_dept", "")),
            "Leaving Date": str(r.get("leaving_date", "")), "Leaving Reason": str(r.get("leaving_reason", "")),
            "Created By": str(r.get("created_by", "")), "Created At": str(r.get("created_at", "")),
            "HR Status": str(r.get("hr_status", "Pending")), "Holiday Balance (Days)": float(r.get("holiday_balance", 0) or 0),
            "Holiday Settlement ID": str(r.get("holiday_settlement_id", "")), "Holiday Settlement Status": str(r.get("holiday_settlement_status", "Not Required")),
            "Holiday Settlement Amount (£)": float(r.get("holiday_settlement_amount", 0) or 0), "Holiday Settlement Type": str(r.get("holiday_settlement_type", "")),
            "Store Status": str(r.get("store_status", "Pending")), "Store Check-in ID": str(r.get("store_checkin_id", "")),
            "Store Deduction ID": str(r.get("store_deduction_id", "")), "Store Return ID": str(r.get("store_return_id", "")),
            "Outstanding Item Value (£)": float(r.get("outstanding_item_value", 0) or 0), "Store Completed By": str(r.get("store_completed_by", "")),
            "Store Completed At": str(r.get("store_completed_at", "")),
            "Payroll Status": str(r.get("payroll_status", "Pending")), "Payroll Request IDs": str(r.get("payroll_request_ids", "")),
            "Payroll Total Addition (£)": float(r.get("payroll_total_addition", 0) or 0), "Payroll Total Deduction (£)": float(r.get("payroll_total_deduction", 0) or 0),
            "Payroll Completed By": str(r.get("payroll_completed_by", "")), "Payroll Completed At": str(r.get("payroll_completed_at", "")),
            "Final Status": str(r.get("final_status", "Open")), "Final Cleared By": str(r.get("final_cleared_by", "")),
            "Final Cleared At": str(r.get("final_cleared_at", "")), "Notes": str(r.get("notes", "")),
        })
    pd.DataFrame(rows, columns=LEAVER_CLEARANCE_COLUMNS).to_excel(LEAVER_CLEARANCE_PATH, index=False, engine="openpyxl")
    _set_data_cache("_leaver_clearance_cache", list(records))
    if sync:
        sync_saved_file_to_drive(LEAVER_CLEARANCE_PATH)


def get_next_leaver_clearance_id(records):
    nums = []
    for r in records or []:
        m = re.fullmatch(r"LC-(\d+)", str(r.get("clearance_id", "")).strip().upper())
        if m:
            nums.append(int(m.group(1)))
    return f"LC-{(max(nums) + 1 if nums else 1):06d}"


def _get_latest_final_settlement(employee_id):
    records = _hrp_get_final_settlement_records(employee_id)
    return records[0] if records else None


def _get_leaver_holiday_balance(employee, leaving_date):
    try:
        if not isinstance(leaving_date, date):
            leaving_date = pd.to_datetime(leaving_date).date()
    except Exception:
        leaving_date = date.today()
    calc = _hrp_get_leaving_entitlement(employee, leaving_date)
    used = _hrp_get_approved_holiday_days_to_date(employee.get("emp_id", ""), leaving_date)
    try:
        start = employee.get("start_date")
        if not isinstance(start, date):
            start = pd.to_datetime(start).date()
    except Exception:
        start = leaving_date
    closure = _hrp_get_company_closure_holiday_days(employee, start, leaving_date)
    return round(round(calc["net"] - calc["bank_holidays"], 1) - round(used + closure, 1), 1)


def _get_payroll_requests_for_clearance(rec):
    ids = {str(x).strip() for x in str(rec.get("payroll_request_ids", "")).split(",") if str(x).strip()}
    if not ids:
        return []
    return [r for r in load_records_from_excel(force=True) if str(r.get("id", "")) in ids]


def _sync_clearance_payroll_status(rec, records=None):
    reqs = _get_payroll_requests_for_clearance(rec)
    if not reqs:
        if str(rec.get("payroll_status", "")).casefold() not in {"cleared", "not required"}:
            rec["payroll_status"] = "Pending"
        rec["payroll_total_addition"] = 0.0
        rec["payroll_total_deduction"] = 0.0
        return rec
    additions = sum(float(r.get("amount", 0) or 0) for r in reqs if str(r.get("type", "")).casefold() == "addition" and str(r.get("status", "")).casefold() == "approved")
    deductions = sum(float(r.get("amount", 0) or 0) for r in reqs if str(r.get("type", "")).casefold() == "deduction" and str(r.get("status", "")).casefold() == "approved")
    rec["payroll_total_addition"] = additions
    rec["payroll_total_deduction"] = deductions
    statuses = {str(r.get("status", "pending")).casefold() for r in reqs}
    if str(rec.get("payroll_status", "")).casefold() == "cleared" and "pending" not in statuses:
        return rec
    if "pending" in statuses:
        rec["payroll_status"] = "Pending Director Approval"
    elif "approved" in statuses:
        rec["payroll_status"] = "Approved"
    elif statuses and statuses.issubset({"rejected"}):
        rec["payroll_status"] = "Rejected"
    else:
        rec["payroll_status"] = "Pending"
    return rec


def _ensure_leaver_clearance(employee, created_by="HR"):
    if not employee or not employee.get("leaving_date"):
        return None
    emp_id = str(employee.get("emp_id", "")).strip()
    if not emp_id:
        return None
    records = load_leaver_clearances(force=True)
    rec = next((r for r in records if str(r.get("employee_id", "")).casefold() == emp_id.casefold()), None)
    try:
        leaving_date = employee.get("leaving_date") if isinstance(employee.get("leaving_date"), date) else pd.to_datetime(employee.get("leaving_date")).date()
    except Exception:
        leaving_date = employee.get("leaving_date")
    holiday_balance = _get_leaver_holiday_balance(employee, leaving_date)
    settlement = _get_latest_final_settlement(emp_id)
    if settlement:
        settlement_status = str(settlement.get("status", "pending")).strip().title()
        settlement_id = str(settlement.get("id", ""))
        settlement_amount = float(settlement.get("amount", 0) or 0)
        settlement_type = str(settlement.get("type", ""))
    elif abs(holiday_balance) < 0.05:
        settlement_status = "Not Required"
        settlement_id = ""
        settlement_amount = 0.0
        settlement_type = ""
    else:
        settlement_status = "Required"
        settlement_id = ""
        settlement_amount = 0.0
        settlement_type = ""
    all_items = load_employee_items()
    outstanding = get_employee_outstanding_items(emp_id, all_items)
    outstanding_value = sum(float(r.get("qty_outstanding", 0) or 0) * float(r.get("unit_price", 0) or 0) for r in outstanding)
    checkins = load_item_checkins(force=True)
    emp_checkins = [c for c in checkins if str(c.get("employee_id", "")).casefold() == emp_id.casefold()]
    latest = emp_checkins[-1] if emp_checkins else None
    created_new = rec is None
    if created_new:
        rec = {
            "clearance_id": get_next_leaver_clearance_id(records), "employee_id": emp_id,
            "emp_name": employee.get("name", ""), "emp_dept": employee.get("department", ""),
            "leaving_date": str(leaving_date), "leaving_reason": str(employee.get("leaving_reason", "")),
            "created_by": created_by or "HR", "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "hr_status": "Complete", "store_status": "Pending", "payroll_status": "Pending",
            "payroll_request_ids": "", "final_status": "Open",
            "final_cleared_by": "", "final_cleared_at": "", "notes": "",
        }
        records.append(rec)
    rec.update({
        "emp_name": employee.get("name", rec.get("emp_name", "")),
        "emp_dept": employee.get("department", rec.get("emp_dept", "")),
        "leaving_date": str(leaving_date),
        "leaving_reason": str(employee.get("leaving_reason", rec.get("leaving_reason", ""))),
        "hr_status": "Complete", "holiday_balance": holiday_balance,
        "holiday_settlement_id": settlement_id, "holiday_settlement_status": settlement_status,
        "holiday_settlement_amount": settlement_amount, "holiday_settlement_type": settlement_type,
        "outstanding_item_value": outstanding_value
    })
    if latest:
        rec["store_checkin_id"] = str(latest.get("id", ""))
        rec["store_deduction_id"] = str(latest.get("deduction_request_id", ""))
        rec["store_return_id"] = str(latest.get("return_request_id", ""))
    if outstanding_value > 0:
        rec["store_status"] = "Pending"
    elif str(rec.get("store_status", "")).casefold() not in {"completed", "cleared"}:
        rec["store_status"] = "Not Required" if not latest else "Check-in Complete"
    _sync_clearance_payroll_status(rec)
    save_all_leaver_clearances(records, sync=created_new)
    return rec


def _clearance_hr_ready(rec):
    return str(rec.get("hr_status", "")).casefold() == "complete"


def _clearance_director_hr_ready(rec):
    status = str(rec.get("holiday_settlement_status", "")).casefold()
    return status in {"approved", "not required"} or (not status and abs(float(rec.get("holiday_balance", 0) or 0)) < 0.05)


def _clearance_store_ready(rec):
    if str(rec.get("store_status", "")).casefold() not in {"not required", "check-in complete", "completed", "cleared"}:
        return False
    if float(rec.get("outstanding_item_value", 0) or 0) > 0:
        return False
    linked_ids = {x.strip() for x in str(rec.get("store_deduction_id", "")).split(",") if x.strip()}
    if not linked_ids:
        return True
    try:
        store_reqs = load_store_deductions(force=True)
        linked = [r for r in store_reqs if str(r.get("id", "")) in linked_ids]
        if not linked:
            return True
        return all(str(r.get("status", "pending")).casefold() == "approved" for r in linked)
    except Exception:
        return False


def _maybe_finalize_clearance(rec, records):
    _sync_clearance_payroll_status(rec)
    if _clearance_all_ready(rec) and str(rec.get("final_status", "")).casefold() != "cleared":
        rec["final_status"] = "Cleared"
        rec["final_cleared_by"] = "System"
        rec["final_cleared_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        save_all_leaver_clearances(records)
        log_action("LEAVER_FINAL_CLEARANCE_COMPLETED", rec.get("clearance_id"), new_data=rec)
        return True
    return False


def _clearance_payroll_ready(rec):
    status = str(rec.get("payroll_status", "Pending")).casefold()
    return status in {"approved", "not required", "cleared"}


def _clearance_all_ready(rec):
    return _clearance_hr_ready(rec) and _clearance_director_hr_ready(rec) and _clearance_store_ready(rec) and _clearance_payroll_ready(rec)


def _create_payroll_clearance_request(rec, user_name, trans_type, amount, reason):
    trans_type = str(trans_type or "").strip().title()
    if trans_type not in {"Addition", "Deduction"}:
        raise ValueError("Payroll adjustment must be Addition or Deduction.")
    amount = float(amount or 0)
    if amount <= 0:
        raise ValueError("Payroll adjustment amount must be greater than £0.")
    reason = str(reason or "").strip()
    if not reason:
        raise ValueError("A reason is required for a Payroll Addition/Deduction.")

    requests = load_records_from_excel(force=True)
    clearance_id = str(rec.get("clearance_id", "")).strip()
    for existing in requests:
        if (str(existing.get("category", "")).casefold() == "leaver payroll adjustment"
                and clearance_id
                and clearance_id in str(existing.get("desc", ""))
                and str(existing.get("type", "")).casefold() == trans_type.casefold()
                and str(existing.get("status", "")).casefold() in {"pending", "approved"}):
            raise ValueError(f"A {trans_type} for {clearance_id} is already pending/approved (Request #{existing.get('id')}).")

    new_id = get_next_id(requests)
    request = {
        "id": new_id, "emp_name": rec.get("emp_name", ""),
        "dept": "Payroll Department", "type": trans_type,
        "category": "Leaver Payroll Adjustment",
        "date": str(rec.get("leaving_date", date.today())),
        "amount": amount, "manager": "Payroll",
        "desc": f"Leaver Clearance {clearance_id}: {reason}",
        "attachment_name": "None", "status": "pending",
        "director_comments": "", "decision_date": "", "decision_by": "",
        "submitted_by": user_name, "pdf_path": "",
        "edited_from_id": "", "old_data": "",
    }
    requests.append(request)
    save_all_records(requests)

    current_ids = [x.strip() for x in str(rec.get("payroll_request_ids", "")).split(",") if x.strip()]
    if str(new_id) not in current_ids:
        current_ids.append(str(new_id))
    rec["payroll_request_ids"] = ",".join(current_ids)
    rec["payroll_status"] = "Pending Director Approval"
    log_action("LEAVER_PAYROLL_ADJUSTMENT_CREATED", new_id, new_data={"clearance_id": clearance_id, **request})
    return request


# ============================================================
# 🧰 EMPLOYEE ITEMS — DATA LAYER
# ============================================================
def initialise_employee_items():
    safe_init_excel(EMPLOYEE_ITEMS_PATH, EMPLOYEE_ITEMS_COLUMNS)
    safe_init_excel(ITEM_CHECKIN_PATH, ITEM_CHECKIN_COLUMNS)


def load_employee_items(force=False):
    if not force and "_employee_items_cache" in st.session_state:
        return list(st.session_state["_employee_items_cache"])
    initialise_employee_items()
    try:
        df = _read_excel_records(EMPLOYEE_ITEMS_PATH)
        records = []
        for r in df.to_dict(orient="records"):
            def _f(k):
                try:
                    return float(r.get(k, 0) or 0)
                except Exception:
                    return 0.0
            try:
                rid = int(r.get("ID", 0))
            except Exception:
                rid = 0
            records.append({
                "id": rid,
                "employee_id": str(r.get("Employee ID", "")).strip(),
                "emp_name": str(r.get("Employee Name", "")).strip(),
                "emp_dept": str(r.get("Department", "")).strip(),
                "item_name": str(r.get("Item Name", "")).strip(),
                "qty_issued": _f("Quantity Issued"),
                "qty_returned": _f("Quantity Returned"),
                "qty_outstanding": _f("Quantity Outstanding"),
                "unit_price": _f("Unit Price (£)"),
                "total_value": _f("Total Value (£)"),
                "issue_date": str(r.get("Issue Date", "")).strip(),
                "issued_by": str(r.get("Issued By", "")).strip(),
                "status": str(r.get("Status", "Issued")).strip(),
                "notes": str(r.get("Notes", "")).strip(),
            })
        _set_data_cache("_employee_items_cache", records)
        return list(records)
    except Exception as e:
        st.error(f"Employee Items Load Error: {e}")
        return []


def save_all_employee_items(records, sync=True):
    rows = [{
        "ID": int(r.get("id", 0)),
        "Employee ID": str(r.get("employee_id", "")),
        "Employee Name": str(r.get("emp_name", "")),
        "Department": str(r.get("emp_dept", "")),
        "Item Name": str(r.get("item_name", "")),
        "Quantity Issued": float(r.get("qty_issued", 0)),
        "Quantity Returned": float(r.get("qty_returned", 0)),
        "Quantity Outstanding": float(r.get("qty_outstanding", 0)),
        "Unit Price (£)": float(r.get("unit_price", 0)),
        "Total Value (£)": float(r.get("total_value", 0)),
        "Issue Date": str(r.get("issue_date", "")),
        "Issued By": str(r.get("issued_by", "")),
        "Status": str(r.get("status", "Issued")),
        "Notes": str(r.get("notes", "")),
    } for r in records]
    pd.DataFrame(rows, columns=EMPLOYEE_ITEMS_COLUMNS).to_excel(EMPLOYEE_ITEMS_PATH, index=False, engine="openpyxl")
    _set_data_cache("_employee_items_cache", list(records))
    if sync:
        sync_saved_file_to_drive(EMPLOYEE_ITEMS_PATH)


def get_next_employee_item_id(records):
    if not records:
        return 1
    return max(int(r.get("id", 0)) for r in records) + 1


def get_employee_outstanding_items(employee_id, records=None):
    if records is None:
        records = load_employee_items()
    target = str(employee_id or "").strip().casefold()
    return [r for r in records
            if str(r.get("employee_id", "")).strip().casefold() == target
            and float(r.get("qty_outstanding", 0) or 0) > 0]


def load_item_checkins(force=False):
    if not force and "_item_checkins_cache" in st.session_state:
        return list(st.session_state["_item_checkins_cache"])
    initialise_employee_items()
    try:
        df = _read_excel_records(ITEM_CHECKIN_PATH)
        records = []
        for r in df.to_dict(orient="records"):
            try:
                rid = int(r.get("ID", 0))
            except Exception:
                rid = 0
            try:
                total = float(r.get("Total Deduction (£)", 0) or 0)
            except Exception:
                total = 0.0
            try:
                returned = json.loads(r.get("Returned Items JSON", "[]") or "[]")
            except Exception:
                returned = []
            try:
                not_returned = json.loads(r.get("Not Returned Items JSON", "[]") or "[]")
            except Exception:
                not_returned = []
            records.append({
                "id": rid,
                "employee_id": str(r.get("Employee ID", "")).strip(),
                "emp_name": str(r.get("Employee Name", "")).strip(),
                "emp_dept": str(r.get("Department", "")).strip(),
                "leaving_date": str(r.get("Leaving Date", "")).strip(),
                "checkin_date": str(r.get("Check-in Date", "")).strip(),
                "checked_in_by": str(r.get("Checked In By", "")).strip(),
                "returned_items": returned,
                "not_returned_items": not_returned,
                "total_deduction": total,
                "deduction_request_id": str(r.get("Deduction Request ID", "")).strip(),
                "return_request_id": str(r.get("Return Request ID", "")).strip(),
                "status": str(r.get("Status", "Completed")).strip(),
                "notes": str(r.get("Notes", "")).strip(),
            })
        _set_data_cache("_item_checkins_cache", records)
        return list(records)
    except Exception as e:
        st.error(f"Item Check-in Load Error: {e}")
        return []


def save_all_item_checkins(records, sync=True):
    rows = [{
        "ID": int(r.get("id", 0)),
        "Employee ID": str(r.get("employee_id", "")),
        "Employee Name": str(r.get("emp_name", "")),
        "Department": str(r.get("emp_dept", "")),
        "Leaving Date": str(r.get("leaving_date", "")),
        "Check-in Date": str(r.get("checkin_date", "")),
        "Checked In By": str(r.get("checked_in_by", "")),
        "Returned Items JSON": json.dumps(r.get("returned_items", []), ensure_ascii=False),
        "Not Returned Items JSON": json.dumps(r.get("not_returned_items", []), ensure_ascii=False),
        "Total Deduction (£)": float(r.get("total_deduction", 0)),
        "Deduction Request ID": str(r.get("deduction_request_id", "")),
        "Return Request ID": str(r.get("return_request_id", "")),
        "Status": str(r.get("status", "Completed")),
        "Notes": str(r.get("notes", "")),
    } for r in records]
    pd.DataFrame(rows, columns=ITEM_CHECKIN_COLUMNS).to_excel(ITEM_CHECKIN_PATH, index=False, engine="openpyxl")
    _set_data_cache("_item_checkins_cache", list(records))
    if sync:
        sync_saved_file_to_drive(ITEM_CHECKIN_PATH)


def get_next_item_checkin_id(records):
    if not records:
        return 1
    return max(int(r.get("id", 0)) for r in records) + 1


# ============================================================
# INITIALISE ALL WORKBOOKS (fast — never reads existing files)
# ============================================================
def initialise_excel():
    safe_init_excel(EXCEL_PATH, EXCEL_COLUMNS)


initialise_excel()
safe_init_excel(WORK_ORDERS_PATH, WORK_ORDER_COLUMNS)
safe_init_excel(INSPECTOR_BONUS_PATH, INSPECTOR_BONUS_COLUMNS)
initialise_hr_leave()
initialise_store_deduction()
initialise_employee_items()
initialise_leaver_clearance()


# ============================================================
# MAIN EXCEL DATA LAYER
# ============================================================
def load_records_from_excel(force=False):
    if not force and "_records_cache" in st.session_state:
        return list(st.session_state["_records_cache"])
    try:
        if not os.path.exists(EXCEL_PATH):
            return _set_data_cache("_records_cache", []).copy()
        df = _read_excel_records(EXCEL_PATH)
        if df.empty:
            return _set_data_cache("_records_cache", []).copy()
        parsed = []
        for r in df.to_dict(orient="records"):
            try:
                record_id = int(r.get("ID", 0))
            except Exception:
                record_id = 0
            try:
                amount = float(r.get("Amount (£)", 0))
            except Exception:
                amount = 0.0
            parsed.append({
                "id": record_id,
                "emp_name": str(r.get("Employee Name", "Not Specified")).strip(),
                "dept": str(r.get("Department", "Not Specified")).strip(),
                "type": str(r.get("Transaction Type", "Not Specified")).strip(),
                "category": str(r.get("Category Reason", "Not Specified")).strip(),
                "date": str(r.get("Date", "")).strip(),
                "amount": amount,
                "manager": str(r.get("Line Manager", "Not Specified")).strip(),
                "desc": str(r.get("Description", "")).strip(),
                "attachment_name": str(r.get("Attachment Name", "None")).strip(),
                "status": str(r.get("Status", "pending")).strip().lower(),
                "director_comments": str(r.get("Director Comments", "")).strip(),
                "decision_date": str(r.get("Decision Date", "")).strip(),
                "decision_by": str(r.get("Decision By", "")).strip(),
                "submitted_by": str(r.get("Submitted By", "")).strip(),
                "pdf_path": str(r.get("PDF File Path", "")).strip(),
                "edited_from_id": str(r.get("Edited From ID", "")).strip(),
                "old_data": str(r.get("Old Data", "")).strip()
            })
        _set_data_cache("_records_cache", parsed)
        return list(parsed)
    except Exception as e:
        st.error(f"Load Error: {e}")
        return []


def save_all_records(records):
    export = []
    for r in records:
        export.append({
            "ID": int(r.get("id", 0)),
            "Employee Name": str(r.get("emp_name", "")),
            "Department": str(r.get("dept", "")),
            "Transaction Type": str(r.get("type", "")),
            "Category Reason": str(r.get("category", "")),
            "Date": str(r.get("date", "")),
            "Amount (£)": float(r.get("amount", 0.0)),
            "Line Manager": str(r.get("manager", "")),
            "Description": str(r.get("desc", "")),
            "Attachment Name": str(r.get("attachment_name", "None")),
            "Status": str(r.get("status", "pending")).lower(),
            "Director Comments": str(r.get("director_comments", "")),
            "Decision Date": str(r.get("decision_date", "")),
            "Decision By": str(r.get("decision_by", "")),
            "Submitted By": str(r.get("submitted_by", "")),
            "PDF File Path": str(r.get("pdf_path", "")),
            "Edited From ID": str(r.get("edited_from_id", "")),
            "Old Data": str(r.get("old_data", ""))
        })
    pd.DataFrame(export, columns=EXCEL_COLUMNS).to_excel(EXCEL_PATH, index=False, engine="openpyxl")
    _set_data_cache("_records_cache", list(records))
    sync_saved_file_to_drive(EXCEL_PATH)


def save_record_to_excel(new_record):
    current = load_records_from_excel()
    current.append(new_record)
    save_all_records(current)


def get_submitted_by(req):
    value = str(req.get("submitted_by", "") or "").strip()
    if value and value.lower() not in ("nan", "none", "-"):
        return value
    try:
        req_id = str(req.get("id", "")).strip()
        if req_id:
            for entry in reversed(load_audit_log()):
                action = str(entry.get("Action", "")).strip().upper()
                audit_req_id = str(entry.get("Request_ID", "")).strip()
                if action == "CREATED" and audit_req_id == req_id:
                    creator = str(entry.get("User_Name", "")).strip()
                    if creator and creator.lower() not in ("nan", "none", "-"):
                        return creator
    except Exception:
        pass
    return ""


def get_completed_by(req):
    for key in ("completed_by", "decision_by", "approved_by"):
        value = str(req.get(key, "") or "").strip()
        if value and value.lower() not in ("nan", "none", "-", "director"):
            return value
    return ""


# ============================================================
# WORK ORDER DATA LAYER
# ============================================================
def load_work_orders(force=False):
    if not force and "_work_orders_cache" in st.session_state:
        return list(st.session_state["_work_orders_cache"])
    safe_init_excel(WORK_ORDERS_PATH, WORK_ORDER_COLUMNS)
    try:
        df = _read_excel_records(WORK_ORDERS_PATH)
        records = []
        for r in df.to_dict(orient="records"):
            try:
                amount = float(r.get("Amount (£)", 0) or 0)
            except Exception:
                amount = 0.0
            try:
                hours = float(r.get("Hours", 0) or 0)
            except Exception:
                hours = 0.0
            records.append({
                "id": str(r.get("Work Order ID", "")).strip(),
                "manual_work_order_no": str(r.get("Manual Work Order No.", "")).strip() or str(r.get("Work Order ID", "")).strip(),
                "emp_name": str(r.get("Employee Name", "")).strip(),
                "dept": str(r.get("Department", "")).strip(),
                "work_date": str(r.get("Work Date", "")).strip(),
                "hours": hours, "amount": amount,
                "manager": str(r.get("Manager", "")).strip(),
                "desc": str(r.get("Description", "")).strip(),
                "attachment_name": str(r.get("Attachment Name", "None")).strip(),
                "status": str(r.get("Status", "pending_manager")).strip().lower(),
                "site_address": str(r.get("Site Address", "")).strip(),
                "customer_job_no": str(r.get("Customer Job No.", "")).strip(),
                "manager_comments": str(r.get("Manager Comments", "")).strip(),
                "manager_decision_date": str(r.get("Manager Decision Date", "")).strip(),
                "manager_decision_by": str(r.get("Manager Decision By", "")).strip(),
                "director_comments": str(r.get("Director Comments", "")).strip(),
                "director_decision_date": str(r.get("Director Decision Date", "")).strip(),
                "director_decision_by": str(r.get("Director Decision By", "")).strip(),
                "submitted_by": str(r.get("Submitted By", "")).strip(),
                "submitted_date": str(r.get("Submitted Date", "")).strip(),
                "payroll_status": str(r.get("Payroll Status", "Pending")).strip(),
                "payroll_date": str(r.get("Payroll Date", "")).strip(),
                "payroll_by": str(r.get("Payroll By", "")).strip(),
                "pdf_path": str(r.get("PDF File Path", "")).strip(),
            })
        _set_data_cache("_work_orders_cache", records)
        return list(records)
    except Exception as e:
        st.error(f"Work Order Load Error: {e}")
        return []


def save_all_work_orders(records, sync=True):
    rows = []
    for r in records:
        rows.append({
            "Work Order ID": str(r.get("id", "")),
            "Manual Work Order No.": str(r.get("manual_work_order_no", "")).strip() or str(r.get("id", "")),
            "Employee Name": str(r.get("emp_name", "")),
            "Department": str(r.get("dept", "")),
            "Work Date": str(r.get("work_date", "")),
            "Hours": float(r.get("hours", 0)),
            "Amount (£)": float(r.get("amount", 0)),
            "Manager": str(r.get("manager", "")),
            "Description": str(r.get("desc", "")),
            "Attachment Name": str(r.get("attachment_name", "None")),
            "Status": str(r.get("status", "pending_manager")),
            "Site Address": str(r.get("site_address", "")),
            "Customer Job No.": str(r.get("customer_job_no", "")),
            "Manager Comments": str(r.get("manager_comments", "")),
            "Manager Decision Date": str(r.get("manager_decision_date", "")),
            "Manager Decision By": str(r.get("manager_decision_by", "")),
            "Director Comments": str(r.get("director_comments", "")),
            "Director Decision Date": str(r.get("director_decision_date", "")),
            "Director Decision By": str(r.get("director_decision_by", "")),
            "Submitted By": str(r.get("submitted_by", "")),
            "Submitted Date": str(r.get("submitted_date", "")),
            "Payroll Status": str(r.get("payroll_status", "Pending")),
            "Payroll Date": str(r.get("payroll_date", "")),
            "Payroll By": str(r.get("payroll_by", "")),
            "PDF File Path": str(r.get("pdf_path", "")),
        })
    pd.DataFrame(rows, columns=WORK_ORDER_COLUMNS).to_excel(WORK_ORDERS_PATH, index=False, engine="openpyxl")
    _set_data_cache("_work_orders_cache", list(records))
    if sync:
        sync_saved_file_to_drive(WORK_ORDERS_PATH)


def get_next_work_order_id(records):
    nums = []
    for r in records:
        raw = str(r.get("id", ""))
        try:
            nums.append(int(raw.replace("WO-", "")))
        except Exception:
            pass
    return f"WO-{max(nums) + 1 if nums else 1:05d}"


def get_work_order_number(req):
    manual = str(req.get("manual_work_order_no", "") or "").strip()
    return manual or str(req.get("id", "") or "").strip()


def _manager_options_for_department(department):
    users = load_users()
    names = []
    for u in users.values():
        if str(u.get("role", "")).lower() == "manager" and str(u.get("dept", "")) == str(department):
            names.append(u.get("full_name", ""))
    return sorted([n for n in names if n])


def _pdf_font_paths():
    candidates = [
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf", "/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf")
    ]
    for regular, bold in candidates:
        if os.path.exists(regular) and os.path.exists(bold):
            return regular, bold
    return None, None


def _pdf_text(value):
    if value is None:
        return ""
    return str(value).replace("\x00", "")


# ============================================================
# INSPECTOR BONUS DATA LAYER
# ============================================================
def load_inspector_bonus(force=False):
    if not force and "_inspector_bonus_cache" in st.session_state:
        return list(st.session_state["_inspector_bonus_cache"])
    safe_init_excel(INSPECTOR_BONUS_PATH, INSPECTOR_BONUS_COLUMNS)
    try:
        df = _read_excel_records(INSPECTOR_BONUS_PATH)
        records = []
        for r in df.to_dict(orient="records"):
            try:
                amount = float(r.get("Bonus Amount (£)", 0) or 0)
            except Exception:
                amount = 0.0
            try:
                jobs = float(r.get("Total Jobs Completed", 0) or 0)
            except Exception:
                jobs = 0.0
            records.append({
                "id": str(r.get("ID", "")).strip(),
                "inspector_name": str(r.get("Inspector Name", "")).strip(),
                "month_year": str(r.get("Month & Year", "")).strip(),
                "days_absent": str(r.get("Days Absent", "")).strip(),
                "reasons": str(r.get("Reasons for Absence", "")).strip(),
                "total_jobs": jobs, "bonus_amount": amount,
                "status": str(r.get("Status", "pending_director")).strip().lower(),
                "director_comments": str(r.get("Director Comments", "")).strip(),
                "director_decision_date": str(r.get("Director Decision Date", "")).strip(),
                "director_decision_by": str(r.get("Director Decision By", "")).strip(),
                "submitted_by": str(r.get("Submitted By", "")).strip(),
                "submitted_date": str(r.get("Submitted Date", "")).strip(),
                "pdf_path": str(r.get("PDF File Path", "")).strip(),
            })
        _set_data_cache("_inspector_bonus_cache", records)
        return list(records)
    except Exception as e:
        st.error(f"Inspector Bonus Load Error: {e}")
        return []


def save_all_inspector_bonus(records, sync=True):
    rows = []
    for r in records:
        rows.append({
            "ID": str(r.get("id", "")),
            "Inspector Name": str(r.get("inspector_name", "")),
            "Month & Year": str(r.get("month_year", "")),
            "Days Absent": str(r.get("days_absent", "")),
            "Reasons for Absence": str(r.get("reasons", "")),
            "Total Jobs Completed": float(r.get("total_jobs", 0)),
            "Bonus Amount (£)": float(r.get("bonus_amount", 0)),
            "Status": str(r.get("status", "pending_director")),
            "Director Comments": str(r.get("director_comments", "")),
            "Director Decision Date": str(r.get("director_decision_date", "")),
            "Director Decision By": str(r.get("director_decision_by", "")),
            "Submitted By": str(r.get("submitted_by", "")),
            "Submitted Date": str(r.get("submitted_date", "")),
            "PDF File Path": str(r.get("pdf_path", "")),
        })
    pd.DataFrame(rows, columns=INSPECTOR_BONUS_COLUMNS).to_excel(INSPECTOR_BONUS_PATH, index=False, engine="openpyxl")
    _set_data_cache("_inspector_bonus_cache", list(records))
    if sync:
        sync_saved_file_to_drive(INSPECTOR_BONUS_PATH)


def get_next_inspector_bonus_id(records):
    nums = []
    for r in records:
        raw = str(r.get("id", ""))
        try:
            nums.append(int(raw.replace("IB-", "")))
        except Exception:
            pass
    return f"IB-{max(nums) + 1 if nums else 1:04d}"


# ============================================================
# LOGIN AND MAIN UI
# ============================================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_info" not in st.session_state:
    st.session_state.user_info = {}


def display_company_header():
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=300)
        else:
            st.title("⚡ ACOOLE ELECTRICAL LTD")
        st.caption("Acoole Operations & Authorisation Portal")
        st.divider()


def change_my_password_form():
    if not st.session_state.get("logged_in") or not st.session_state.get("user_info"):
        return
    render_google_drive_status()

    with st.sidebar.expander("🔑 Change My Password", expanded=False):
        USERS = load_users()
        current_username = st.session_state.user_info.get("username")
        if not current_username or current_username not in USERS:
            st.warning("⚠️ Please log in before changing password.")
            return
        with st.form("change_my_password", clear_on_submit=True):
            old_pass = st.text_input("Current Password", type="password")
            new_pass1 = st.text_input("New Password", type="password")
            new_pass2 = st.text_input("Confirm New Password", type="password")
            if st.form_submit_button("✅ Update Password", type="primary"):
                if USERS[current_username]["password"] != old_pass:
                    st.error("❌ Current password is NOT correct!")
                    return
                if new_pass1 != new_pass2:
                    st.error("❌ New passwords do NOT match!")
                    return
                if len(new_pass1) < 4:
                    st.error("❌ New password must be at least 4 characters!")
                    return
                USERS[current_username]["password"] = new_pass1
                save_users(USERS)
                st.session_state.user_info["password"] = new_pass1
                st.success("✅ Password changed successfully!")
                st.balloons()
                st.rerun()


def refresh_authenticated_user_session():
    if not st.session_state.get("logged_in"):
        return
    current_username = str(st.session_state.get("user_info", {}).get("username", "")).strip().lower()
    if not current_username:
        return
    try:
        users = load_users(force=False)
        current = users.get(current_username)
        if not current:
            st.session_state.clear()
            st.rerun()
            return
        if not current.get("is_active", True):
            st.session_state.clear()
            st.rerun()
            return
        st.session_state.user_info = {**current, "username": current_username}
    except Exception as e:
        print(f"Authenticated user permission refresh failed: {e}")


if not st.session_state.logged_in:
    display_company_header()
    with st.form("login_form", border=True):
        st.markdown("### 🔒 Secure Gateway Login")
        st.caption("Enter your credentials to access the system")
        st.divider()
        username = st.text_input("🔐 Username", placeholder="e.g. andy, payroll, wais").lower().strip()
        password = st.text_input("🔑 Password", type="password", placeholder="Enter your password")
        if st.form_submit_button("🔐 Authenticate Portal", type="primary", width="stretch"):
            USERS = load_users()
            if username in USERS and USERS[username]["password"] == password:
                if not USERS[username].get("is_active", True):
                    st.error("❌ Your account has been deactivated. Please contact your Super Admin.")
                else:
                    st.session_state.logged_in = True
                    st.session_state.user_info = {**USERS[username], "username": username}
                    st.rerun()
            else:
                st.error("❌ Invalid Username or Password. Please try again.")
    st.stop()

refresh_authenticated_user_session()

col_left, col_right = st.columns([4, 1])
with col_left:
    refresh_data_button()
with col_right:
    if st.button("🔒 Secure Logout", type="secondary", key="top_right_logout"):
        st.session_state.clear()
        st.rerun()

display_company_header()

user_info = st.session_state.get("user_info", {})
full_name = user_info.get("full_name", user_info.get("username", "User"))
dept = user_info.get("dept", "")
role = user_info.get("role", "")

st.info(f"👤 Welcome: {full_name} | {dept} | {role}")

change_my_password_form()

# Load live requests on demand (no Drive call)
all_live_requests = load_records_from_excel(force=False)
CATEGORIES = load_categories()

st.divider()


# ============================================================
# Minimal role-based portal (placeholder — full portals can be added below)
# ============================================================
if role == "Super Admin":
    st.subheader("🛡️ Super Admin Portal")
    st.success("✅ Logged in as Super Admin. Use the buttons and tools below.")
    st.caption("To add your Super Admin tools (System Management, Data Control, etc.), "
               "extend this section with the same code you previously had. "
               "The page will now LOAD correctly because Drive is no longer blocking startup.")
    if st.button("🔄 Manually Sync All Data to Google Drive", key="sa_sync_btn"):
        with st.spinner("Uploading all workbooks..."):
            _ensure_drive_service()
            if drive_service is None:
                st.error(f"Google Drive not connected: {DRIVE_CONNECTION_ERROR}")
            else:
                results = _force_sync_all_drive_files()
                for label, ok in results:
                    st.write(("✅ " if ok else "❌ ") + label)

elif role == "Director":
    st.subheader("🎛️ Director Portal")
    st.success(f"✅ Logged in as Director: {full_name}")
    st.caption("Approvals will appear here once you add the Director tabs back.")

elif role == "Payroll":
    st.subheader("🧾 Payroll Portal")
    st.success(f"✅ Logged in as Payroll: {full_name}")

elif role in ["Manager", "Staff", "Team Member", "Work Order Manager", "Work Order Employee"]:
    st.subheader(f"👤 {role} Portal")
    st.success(f"✅ Logged in as {role}: {full_name} | {dept}")

elif role == "Employee":
    st.subheader("👤 Employee Portal")
    st.success(f"✅ Logged in as Employee: {full_name}")

else:
    st.subheader("🔐 Access Restricted")
    st.error("❌ Your role does not have a defined portal. Please contact Super Admin.")

st.divider()
st.caption("Acoole Portal v4.31 — Google Drive is now lazy-loaded to prevent startup blocking.")
