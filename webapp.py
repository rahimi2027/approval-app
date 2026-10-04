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
    "Employee ID", "Full Name", "Start Date", "Position / Job Title", "Department",
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
_DRIVE_SYNC_QUEUE = None

# Local persistence / recovery. Google Drive is a backup layer; these local files
# are the durable working copy and RAM is never treated as the only copy.
LOCAL_RECOVERY_DIR = os.path.join(APP_FOLDER, "local_recovery")
DRIVE_SYNC_STATE_PATH = os.path.join(APP_FOLDER, "drive_sync_state.json")
DRIVE_AUTO_SYNC_INTERVAL_SECONDS = 3600
_DRIVE_WORKER_STARTED = False
_DRIVE_WORKER_WAKE = threading.Event()
_DRIVE_STATE_LOCK = threading.RLock()
os.makedirs(LOCAL_RECOVERY_DIR, exist_ok=True)

# ============================================================
# GOOGLE DRIVE CONNECTION — SERVICE ACCOUNT (BASE64 METHOD)
# ============================================================
try:
    gdrive = st.secrets["gdrive"]
    b64_string = str(gdrive["key_b64"]).replace("\n", "").replace("\r", "").replace(" ", "").replace("\t", "")
    padding_needed = (4 - len(b64_string) % 4) % 4
    if padding_needed:
        b64_string += "=" * padding_needed
    decoded_json = base64.b64decode(b64_string).decode("utf-8")
    creds_dict = json.loads(decoded_json)
    credentials = service_account.Credentials.from_service_account_info(
        creds_dict, scopes=SCOPES
    )
    drive_service = build("drive", "v3", credentials=credentials, cache_discovery=False)
    folder_check = drive_service.files().get(
        fileId=GOOGLE_DRIVE_FOLDER_ID,
        fields="id,name,mimeType,parents",
        supportsAllDrives=True,
    ).execute()
    DRIVE_CONNECTION_ERROR = ""
    print(f"Google Drive connected: {folder_check.get('name', GOOGLE_DRIVE_FOLDER_ID)}")
except Exception as e:
    drive_service = None
    DRIVE_CONNECTION_ERROR = f"{type(e).__name__}: {e}"
    print(f"Google Drive initialisation failed; local storage will be used: {DRIVE_CONNECTION_ERROR}")

def upload_to_google_drive(local_file_path, display_filename):
    if drive_service is None or not os.path.exists(local_file_path): return None
    cache_key = f"{GOOGLE_DRIVE_FOLDER_ID}::{display_filename}"
    def _do(file_id=None):
        media = MediaFileUpload(local_file_path, resumable=False)
        if file_id:
            return drive_service.files().update(fileId=file_id, media_body=media, fields="id,name,parents", supportsAllDrives=True).execute()
        metadata = {"name": display_filename, "parents": [GOOGLE_DRIVE_FOLDER_ID]}
        return drive_service.files().create(body=metadata, media_body=media, fields="id,name,parents", supportsAllDrives=True).execute()
    try:
        cached_id = _DRIVE_ID_CACHE.get(cache_key)
        if cached_id:
            try: return _do(cached_id).get("id")
            except Exception: _DRIVE_ID_CACHE.pop(cache_key, None)
        existing = _drive_find_file(display_filename)
        if existing:
            _DRIVE_ID_CACHE[cache_key] = existing["id"]
            return _do(existing["id"]).get("id")
        created = _do(None)
        _DRIVE_ID_CACHE[cache_key] = created.get("id")
        return created.get("id")
    except Exception as e:
        DRIVE_LAST_SYNC_ERROR[display_filename] = f"{type(e).__name__}: {e}"
        print(f"Google Drive upload failed for {display_filename}: {e}")
        return None

_DRIVE_FIND_CACHE = {}
_DRIVE_FIND_CACHE_TTL = 60.0
_DRIVE_ID_CACHE = {}

def _drive_find_file(filename, parent_id=GOOGLE_DRIVE_FOLDER_ID):
    if drive_service is None:
        return None
    import time
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
    if drive_service is None or not os.path.exists(local_path): return None
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
            try: return _do(cached_id).get("id")
            except Exception: _DRIVE_ID_CACHE.pop(cache_key, None)
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

def _drive_remote_workbook_has_rows(file_id):
    if drive_service is None or not file_id:
        return False
    tmp_path = os.path.join(APP_FOLDER, f".__drive_check_{file_id}.xlsx")
    try:
        if not _drive_download_file(file_id, tmp_path):
            return False
        df = pd.read_excel(tmp_path, engine="openpyxl")
        return not df.empty
    except Exception as e:
        print(f"Drive workbook row check failed: {e}")
        return False
    finally:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass


def sync_persistent_file(local_path, columns=None):
    if drive_service is None:
        if not os.path.exists(local_path) and columns is not None:
            pd.DataFrame(columns=columns).to_excel(local_path, index=False, engine="openpyxl")
        return

    with _DRIVE_SYNC_LOCK:
        filename = os.path.basename(local_path)
        backup_name = DRIVE_BACKUP_FILENAMES.get(local_path, "")
        remote = _drive_find_file(filename)

        if remote:
            if backup_name:
                backup_remote = _drive_find_file(backup_name)
                if backup_remote and not _drive_remote_workbook_has_rows(remote["id"]):
                    if _drive_remote_workbook_has_rows(backup_remote["id"]):
                        if _drive_download_file(backup_remote["id"], local_path):
                            print(f"Recovered {filename} from non-empty Drive backup {backup_name}")
                            return

            if _drive_download_file(remote["id"], local_path):
                return
            print(f"Using local copy of {filename} because Drive download failed")
            if os.path.exists(local_path):
                return

        if backup_name:
            backup_remote = _drive_find_file(backup_name)
            if backup_remote and _drive_remote_workbook_has_rows(backup_remote["id"]):
                if _drive_download_file(backup_remote["id"], local_path):
                    print(f"Recovered missing {filename} from Drive backup {backup_name}")
                    return

        if not os.path.exists(local_path) and columns is not None:
            pd.DataFrame(columns=columns).to_excel(local_path, index=False, engine="openpyxl")
            _drive_upload_path(local_path, filename)

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

        # Keep the latest 5 snapshots per workbook.
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
    """Queue a local save for the next hourly Drive upload without blocking the user."""
    if not local_path or not os.path.exists(local_path):
        return False
    with _DRIVE_SYNC_LOCK:
        if make_recovery_snapshot:
            _create_local_recovery_snapshot(local_path)
        DRIVE_PENDING_SYNC.add(os.path.abspath(local_path))
        _save_drive_sync_state()
    _DRIVE_WORKER_WAKE.set()
    return True


def sync_backup_file_to_drive(local_path):
    """Immediate backup helper used only by the explicit/manual backup path."""
    if drive_service is None or not os.path.exists(local_path):
        return None
    backup_name = DRIVE_BACKUP_FILENAMES.get(local_path)
    if not backup_name:
        return None
    try:
        return _drive_upload_path(local_path, backup_name)
    except Exception as e:
        print(f"Drive backup sync failed for {backup_name}: {e}")
        return None


def _sync_saved_file_to_drive_now(local_path):
    """Actually upload one file. This is deliberately called only by the background worker or manual force-sync."""
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
                live_error = DRIVE_LAST_SYNC_ERROR.get(
                    filename, "Live Google Drive upload returned no file ID."
                )
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
                    DRIVE_LAST_SYNC_ERROR[backup_name] = DRIVE_LAST_SYNC_ERROR.get(
                        backup_name, "Backup Google Drive upload returned no file ID."
                    )
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
            # Only remove the pending flag after at least one Drive copy succeeded.
            DRIVE_PENDING_SYNC.discard(local_path)
            _save_drive_sync_state()
            return True

        DRIVE_PENDING_SYNC.add(local_path)
        _DRIVE_SYNC_FINGERPRINTS.pop(local_path, None)
        _save_drive_sync_state()
        return False


_DRIVE_STORAGE_PROCESS_READY = False


def sync_saved_file_to_drive(local_path):
    """Save hook for persistent data. Super Admin writes sync immediately; other users queue for hourly backup."""
    filename = os.path.basename(local_path)
    if not os.path.exists(local_path):
        DRIVE_LAST_SYNC_ERROR[filename] = "Local file does not exist."
        return False

    role = str(st.session_state.get("user_info", {}).get("role", "")).strip() if hasattr(st, "session_state") else ""
    if role == "Super Admin" and drive_service is not None:
        # Super Admin changes are business-critical: upload the live workbook and its
        # rolling backup immediately, rather than waiting for the hourly worker.
        return _sync_saved_file_to_drive_now(local_path)

    return _mark_drive_file_dirty(local_path)


def ensure_all_drive_backups():
    """Queue all local workbooks; actual uploads are performed by the hourly worker."""
    for local_path in DRIVE_BACKUP_FILENAMES:
        if os.path.exists(local_path):
            _mark_drive_file_dirty(local_path, make_recovery_snapshot=False)


def _upload_to_drive_bg(local_path, filename):
    # Kept for compatibility with existing code. It now means queue for background sync.
    _mark_drive_file_dirty(local_path, make_recovery_snapshot=False)


def _drive_background_worker():
    """Process-local daemon. It never touches Streamlit session state/UI."""
    while True:
        try:
            # Wake on a new save, but do not upload immediately. The hourly interval is the
            # normal schedule; a wake simply lets the worker notice pending work.
            _DRIVE_WORKER_WAKE.wait(timeout=60)
            _DRIVE_WORKER_WAKE.clear()
            if drive_service is None:
                continue
            last_run = getattr(_drive_background_worker, "last_run", 0.0)
            now = time.time()
            if now - last_run < DRIVE_AUTO_SYNC_INTERVAL_SECONDS:
                continue
            _drive_background_worker.last_run = now
            # Every hour back up ALL existing persistent workbooks, not only files
            # that happened to be changed. This keeps the Drive copy current even
            # after a quiet hour. Pending files are naturally included in this pass.
            targets = [path for path in DRIVE_BACKUP_FILENAMES if os.path.exists(path)]
            targets = sorted(set(targets) | set(DRIVE_PENDING_SYNC))
            if not targets:
                _save_drive_sync_state()
                continue

            print(f"Automatic Drive backup starting for {len(targets)} file(s).")
            for local_path in targets:
                try:
                    if os.path.exists(local_path):
                        _sync_saved_file_to_drive_now(local_path)
                    else:
                        DRIVE_PENDING_SYNC.discard(local_path)
                except Exception as e:
                    DRIVE_LAST_SYNC_ERROR[os.path.basename(local_path)] = f"{type(e).__name__}: {e}"
                    print(f"Automatic Drive backup failed for {local_path}: {e}")
            _save_drive_sync_state()
        except Exception as e:
            print(f"Drive background worker error: {type(e).__name__}: {e}")
            time.sleep(10)


def _start_drive_background_worker():
    global _DRIVE_WORKER_STARTED
    if _DRIVE_WORKER_STARTED:
        return
    if drive_service is None:
        return
    _restore_drive_sync_state()
    _DRIVE_WORKER_STARTED = True
    # Existing unsynced work from a previous process should be backed up promptly.
    # If there is no pending work, start the normal one-hour clock now.
    _drive_background_worker.last_run = 0.0 if DRIVE_PENDING_SYNC else time.time()
    worker = threading.Thread(
        target=_drive_background_worker,
        name="acoole-drive-backup-worker",
        daemon=True,
    )
    worker.start()
    print("Automatic Google Drive backup worker started (hourly).")


def _force_sync_all_drive_files():
    """Explicit Super Admin/Director action: upload every existing workbook immediately."""
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
            st.error("Google Drive is NOT connected. Local data is still being saved.")
            st.code(DRIVE_CONNECTION_ERROR or "No connection error was captured.")
            st.caption(f"Folder ID: {GOOGLE_DRIVE_FOLDER_ID}")
        else:
            st.success("Google Drive backup is connected.")
            st.caption("Normal user actions use local persistent files. Google Drive is updated automatically every hour.")

        if st.button("🔄 Save All Data to Google Drive Now", key="force_drive_sync_all"):
            if drive_service is None:
                st.error("Cannot sync because Google Drive is not connected. Local data remains safe on disk.")
            else:
                results = []
                with st.spinner("Saving all local data to Google Drive..."):
                    results = _force_sync_all_drive_files()
                for label, ok in results:
                    st.write(("✅ " if ok else "❌ ") + label)
                st.rerun()

        rows = _drive_status_rows()
        st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
        pending = len(DRIVE_PENDING_SYNC)
        st.caption(
            f"Pending Drive uploads: {pending}. Local saves never wait for Google Drive. "
            f"Automatic backup interval: {DRIVE_AUTO_SYNC_INTERVAL_SECONDS // 60} minutes."
        )

@st.cache_resource(show_spinner=False)
def _initialise_drive_storage_once():
    if drive_service is None:
        return False
    os.makedirs(APP_FOLDER, exist_ok=True)
    critical = [
        (EXCEL_PATH, EXCEL_COLUMNS),
        (USER_DB_PATH, USER_DB_COLUMNS),
        (SETTINGS_PATH, ["setting", "value"]),
        (HR_EMPLOYEES_PATH, HR_EMPLOYEE_COLUMNS),
        (HR_PORTAL_LEAVE_PATH, HR_PORTAL_LEAVE_COLUMNS),
    ]
    for path, columns in critical:
        try:
            if not os.path.exists(path) or os.path.getsize(path) == 0:
                sync_persistent_file(path, columns)
        except Exception as e:
            print(f"Critical Drive initialisation failed for {os.path.basename(path)}: {e}")
    print("Critical Drive workbooks initialised once for this Streamlit server process.")
    _restore_drive_sync_state()
    # Existing local files are authoritative during the current working session.
    # Only missing/empty files are restored from Drive here, preventing an older Drive
    # copy from overwriting local work that has not yet been backed up.
    _start_drive_background_worker()
    return True

def initialise_drive_storage():
    return _initialise_drive_storage_once()

def get_onedrive_token():
    if not ONEDRIVE_CLIENT_ID or not ONEDRIVE_CLIENT_SECRET: return None
    try:
        url = f"https://login.microsoftonline.com/{ONEDRIVE_TENANT_ID}/oauth2/v2.0/token"
        data = {"grant_type": "client_credentials", "client_id": ONEDRIVE_CLIENT_ID, "client_secret": ONEDRIVE_CLIENT_SECRET, "scope": "https://graph.microsoft.com/.default"}
        res = requests.post(url, data=data, timeout=30)
        if res.status_code == 200: return res.json().get("access_token")
    except Exception as e: st.warning(f"⚠️ OneDrive connection: {e}")
    return None

def upload_to_onedrive(local_file_path, remote_filename=None):
    if not USE_ONEDRIVE: return False
    token = get_onedrive_token()
    if not token: return False
    filename = remote_filename or os.path.basename(local_file_path)
    remote_path = f"{ONEDRIVE_FOLDER}{filename}"
    try:
        with open(local_file_path, 'rb') as f: file_content = f.read()
        url = f"https://graph.microsoft.com/v1.0/drives/me/items/root:/{remote_path}:/content"
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/octet-stream"}
        res = requests.put(url, data=file_content, headers=headers, timeout=60)
        if res.status_code in (200, 201): return True
    except Exception: pass
    return False

DEFAULT_CATEGORIES = ["Food Allowance", "Others", "Parking", "Parking Fine", "GYM Membership", "Item Not Returned", "Item Missing"]
DEFAULT_ROLES = ["Manager", "Staff", "Team Member", "Employee", "Work Order Employee", "Work Order Manager", "Director", "Payroll", "Super Admin"]
DEFAULT_DEPARTMENTS = ["National Grid", "Isolator", "Project", "Accounts", "Payroll Department", "ACoole Electrical Ltd", "Store", "HR"]
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
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    if not os.path.exists(path):
        tmp_path = f"{path}.init.tmp.xlsx"
        try:
            pd.DataFrame(columns=columns).to_excel(tmp_path, index=False, engine="openpyxl")
            os.replace(tmp_path, path)
            return True
        finally:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass
    try:
        df = pd.read_excel(path, engine="openpyxl").fillna("")
        changed = False
        for col in columns:
            if col not in df.columns:
                df[col] = ""
                changed = True
        if changed:
            tmp_path = f"{path}.repair.tmp.xlsx"
            try:
                df.to_excel(tmp_path, index=False, engine="openpyxl")
                os.replace(tmp_path, path)
            finally:
                try:
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)
                except Exception:
                    pass
        return True
    except Exception as e:
        print(f"Excel initialisation/recovery for {path} failed: {e}")
        if os.path.exists(path):
            return False
        tmp_path = f"{path}.init.tmp.xlsx"
        try:
            pd.DataFrame(columns=columns).to_excel(tmp_path, index=False, engine="openpyxl")
            os.replace(tmp_path, path)
            return True
        finally:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass

def _invalidate_data_cache(*names):
    for name in names: st.session_state.pop(name, None)

def _set_data_cache(name, value):
    st.session_state[name] = value
    return value

def _read_excel_records(path):
    if not os.path.exists(path): return pd.DataFrame()
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
    except Exception as e: print(f"⚠️ Audit lookup failed: {e}")
    return dept, amount, decision_by, decision_date

def _audit_json(value):
    try:
        return json.dumps(value, ensure_ascii=False, default=str)[:300] if value else "-"
    except Exception:
        return str(value)[:300] if value is not None else "-"

def log_action(action, req_id="-", old_data=None, new_data=None, fields_changed=None, decision_by=None, decision_date=None):
    if not st.session_state.get("logged_in"): return
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
    st.info("🔒 Super Admin Only — The audit history can be cleared from the Danger Zone below."); st.divider()
    logs = load_audit_log()
    if not logs: st.info("📋 No activity recorded yet."); return
    c1, c2, c3, c4 = st.columns(4)
    with c1: filter_user = st.multiselect("👤 Filter by User", sorted(set([l["User_Name"] for l in logs])))
    with c2: dept_list = sorted(set([l.get("Department", "") for l in logs if l.get("Department") != "-"]))
    filter_dept = st.multiselect("🏢 Filter by Department", dept_list)
    with c3: filter_action = st.multiselect("🔧 Filter by Action", sorted(set([l["Action"] for l in logs])))
    with c4: req_list = sorted(set([str(l["Request_ID"]) for l in logs if str(l["Request_ID"]) != "-"]))
    filter_req = st.multiselect("🆔 Filter by Request ID", req_list)
    filtered = logs
    if filter_user: filtered = [l for l in filtered if l["User_Name"] in filter_user]
    if filter_dept: filtered = [l for l in filtered if l.get("Department", "") in filter_dept]
    if filter_action: filtered = [l for l in filtered if l["Action"] in filter_action]
    if filter_req: filtered = [l for l in filtered if str(l["Request_ID"]) in filter_req]
    st.metric(f"📄 Total Entries", len(filtered)); st.divider()
    for entry in reversed(filtered):
        aid, ts, user, role, action, req_id = entry["AuditID"], entry["Timestamp"], entry["User_Name"], entry["User_Role"], entry["Action"], entry["Request_ID"]
        dept, amount, dec_by, dec_date = entry.get("Department", "-"), entry.get("Amount", "-"), entry.get("Decision_By", "-"), entry.get("Decision_Date", "-")
        field, old_val, new_val = entry["Field_Changed"], entry["Old_Value"], entry["New_Value"]
        icon = {"CREATED": "➕", "EDITED": "✏️", "APPROVED": "✅", "REJECTED": "❌", "DELETED": "🗑️", "STATUS_CHANGED": "🔄"}.get(action, "ℹ️")
        title = f"{icon} {action}" + (f" — Request #{req_id}" if str(req_id) != "-" else "") + f" | {user} ({role}) | {ts}"
        with st.expander(title):
            st.write(f"**🕐 Time:** {ts}")
            st.write(f"**👤 User:** {user} — *{role}*")
            if str(req_id) != "-": st.write(f"**🆔 Request ID:** #{req_id}")
            if dept != "-": st.write(f"**🏢 Department:** {dept}")
            if amount != "-": st.write(f"**💷 Amount:** {amount}")
            if action in ["APPROVED", "REJECTED", "STATUS_CHANGED"]:
                st.write(f"**🎯 Decision By:** {dec_by}")
                st.write(f"**📅 Decision Date:** {dec_date}")
            if field and field != "-" and field != "No Changes":
                st.write(f"**📝 Field Changed:** {field}")
                if old_val and old_val != "-": st.markdown(f"**⬅️ Old:** `{old_val}`")
                if new_val and new_val != "-": st.markdown(f"**➡️ New:** `{new_val}`")
            else:
                st.write(f"**📋 Details:** {new_val}")
    st.divider()
    df_export = pd.DataFrame(filtered)
    st.download_button("📥 Download Full Audit Log (CSV)", df_export.to_csv(index=False).encode("utf-8"), "Acoole_Audit_Log.csv", type="primary")

def format_date(d):
    if not d or str(d).strip().lower() in ["", "none", "nan"]: return "-"
    return str(d).strip()

def get_next_id(all_records):
    if not all_records: return 1
    return max(int(r.get("id", 0)) for r in all_records) + 1

def show_old_new_comparison(old_json, new_rec):
    try: old = json.loads(old_json) if old_json and old_json != "{}" else {}
    except: old = {}
    if not old: st.info("📋 New request — no previous version."); return
    st.markdown("#### 🔄 Changes (Previous → New)")
    fields = [("emp_name", "Employee Name"), ("dept", "Department"), ("type", "Transaction Type"), ("category", "Category"), ("date", "Date"), ("amount", "Amount (£)"), ("manager", "Line Manager"), ("desc", "Description")]
    changed = False
    for key, label in fields:
        o, n = str(old.get(key, "")).strip(), str(new_rec.get(key, "")).strip()
        if o != n:
            changed = True
            st.markdown(f"**{label}**: ~~`{o}`~~ → **`{n}`**")
    if not changed: st.info("✅ No changes detected.")

def refresh_data_button():
    if st.button("🔄 Refresh Data", type="secondary", key="refresh_data_btn"):
        with st.spinner("Refreshing from Google Drive..."):
            if drive_service is not None:
                for path in (EXCEL_PATH, USER_DB_PATH, SETTINGS_PATH, AUDIT_LOG_PATH, INSPECTOR_BONUS_PATH, WORK_ORDERS_PATH, HR_LEAVE_PATH, HR_DAILY_RATES_PATH, STORE_DEDUCTION_PATH, STORE_ITEMS_PATH, HR_EMPLOYEES_PATH, HR_PORTAL_LEAVE_PATH, EMPLOYEE_ITEMS_PATH, ITEM_CHECKIN_PATH, LEAVER_CLEARANCE_PATH):
                    with _DRIVE_SYNC_LOCK:
                        pending = path in DRIVE_PENDING_SYNC
                    if pending:
                        continue
                    remote = _drive_find_file(os.path.basename(path))
                    if remote:
                        _drive_download_file(remote["id"], path)
            _invalidate_data_cache("_records_cache", "_users_cache", "_settings_cache", "_audit_log_cache", "_work_orders_cache", "_inspector_bonus_cache", "_audit_log_count", "_hr_leave_cache", "_hr_daily_rates_cache", "_store_deduction_cache", "_store_items_cache", "_employee_items_cache", "_item_checkins_cache")
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

def load_departments(): return _load_setting_value("departments", DEFAULT_DEPARTMENTS)

def save_departments(dept_list):
    init_settings()
    df = _read_excel_records(SETTINGS_PATH)
    found = False
    for idx, r in df.iterrows():
        if r["setting"] == "departments":
            df.at[idx, "value"] = "|".join(dept_list); found = True
    if not found:
        df = pd.concat([df, pd.DataFrame([{"setting": "departments", "value": "|".join(dept_list)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False, engine="openpyxl")
    _invalidate_data_cache("_settings_cache")
    sync_saved_file_to_drive(SETTINGS_PATH)

def load_categories(): return _load_setting_value("categories", DEFAULT_CATEGORIES)

def save_categories(cat_list):
    init_settings()
    df = _read_excel_records(SETTINGS_PATH)
    found = False
    for idx, r in df.iterrows():
        if r["setting"] == "categories":
            df.at[idx, "value"] = "|".join(cat_list); found = True
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
            df.at[idx, "value"] = "|".join(roles_list); found = True
    if not found:
        df = pd.concat([df, pd.DataFrame([{"setting": "roles", "value": "|".join(roles_list)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False, engine="openpyxl")
    _invalidate_data_cache("_settings_cache")
    sync_saved_file_to_drive(SETTINGS_PATH)

def load_hr_categories(): return _load_setting_value("hr_categories", DEFAULT_HR_CATEGORIES)

def save_hr_categories(cats):
    init_settings()
    df = _read_excel_records(SETTINGS_PATH)
    found = False
    for idx, r in df.iterrows():
        if r["setting"] == "hr_categories":
            df.at[idx, "value"] = "|".join(cats); found = True
    if not found:
        df = pd.concat([df, pd.DataFrame([{"setting": "hr_categories", "value": "|".join(cats)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False, engine="openpyxl")
    _invalidate_data_cache("_settings_cache")
    sync_saved_file_to_drive(SETTINGS_PATH)

try:
    initialise_drive_storage()
except Exception as e:
    print(f"Drive storage initialisation skipped: {e}")

def init_user_db():
    safe_init_excel(USER_DB_PATH, USER_DB_COLUMNS)
    try:
        df = pd.read_excel(USER_DB_PATH, engine="openpyxl")
        if df.empty:
            pd.DataFrame(DEFAULT_USERS).to_excel(USER_DB_PATH, index=False, engine="openpyxl")
            sync_saved_file_to_drive(USER_DB_PATH)
    except:
        pd.DataFrame(DEFAULT_USERS).to_excel(USER_DB_PATH, index=False, engine="openpyxl")
        sync_saved_file_to_drive(USER_DB_PATH)

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
            if not username: continue
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

        # Auto-grant Store Department Deduction permission to any user whose
        # department is "Store". They run the item issuing and leaver check-in
        # workflow for all employees.
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
                try: os.remove(tmp_path)
                except OSError: pass
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
        st.session_state.hrp_adjustment_notes = {e["emp_id"]: e.get("adjustment_note", "") for e in st.session_state.hrp_employees if e.get("adjustment_note") }
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
    if end_date < start_date: return 0
    total = 0
    current = start_date
    while current <= end_date:
        if current.weekday() < 5 and current not in _hrp_non_working_dates(current.year):
            total += 1
        current += timedelta(days=1)
    return total

def _hrp_calculate_leave_days(leave_type, start_date, end_date, half_day=False):
    if end_date < start_date: return 0.0
    if leave_type == "Half Day Holiday" or half_day:
        return 0.5 if _hrp_get_working_days(start_date, end_date) > 0 else 0.0
    return float(_hrp_get_working_days(start_date, end_date))

def _hrp_calculate_service_years(start_date):
    days = (date.today() - start_date).days
    if days < 0: return 0.0
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
        if years: parts.append(f"{years} year" + ("s" if years != 1 else ""))
        if months: parts.append(f"{months} month" + ("s" if months != 1 else ""))
        if days or not parts: parts.append(f"{days} day" + ("s" if days != 1 else ""))
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
    return {"leave_records": leave_records, "settlements": settlements, "holiday_taken": round(holiday_taken,1), "types": types}


def _hrp_leaver_history_pdf(employee, history):
    if not PDF_AVAILABLE:
        return None
    try:
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=14)
        pdf.add_page()
        regular_font, bold_font = _pdf_font_paths()
        if regular_font and bold_font:
            pdf.add_font("DejaVu", "", regular_font)
            pdf.add_font("DejaVu", "B", bold_font)
            family = "DejaVu"
        else:
            family = "Helvetica"

        def safe(v):
            txt = _pdf_text(v)
            return txt if family == "DejaVu" else txt.encode("latin-1", "replace").decode("latin-1")

        def hard_wrap(v, width=80):
            txt = safe(v)
            return "\n".join(
                "\n".join(textwrap.wrap(part, width=width, break_long_words=True, break_on_hyphens=False) or [""])
                for part in txt.splitlines()
            ) or "-"

        def full_line(text, h=6, bold=False, size=9):
            pdf.set_x(pdf.l_margin)
            pdf.set_font(family, "B" if bold else "", size)
            pdf.multi_cell(pdf.epw, h, hard_wrap(text), new_x="LMARGIN", new_y="NEXT")

        if os.path.exists(LOGO_PATH):
            try:
                pdf.image(LOGO_PATH, x=75, y=10, w=60)
                pdf.ln(28)
            except Exception:
                pdf.ln(5)

        full_line("EMPLOYEE LEAVER COMPLETE HISTORY", 10, True, 16)
        pdf.ln(3)
        full_line("EMPLOYEE DETAILS", 7, True, 11)

        details = [
            ("Employee ID", employee.get("emp_id")), ("Full Name", employee.get("name")),
            ("Department", employee.get("department")), ("Position", employee.get("job_title")),
            ("Start Date", employee.get("start_date")), ("End / Leaving Date", employee.get("leaving_date")),
            ("Duration of Work", _hrp_work_duration(employee.get("start_date"), employee.get("leaving_date"))),
            ("Agreement", employee.get("agreement_type")), ("Working Pattern", employee.get("working_pattern")),
            ("Days Per Week", employee.get("days_per_week")), ("Reason for Leaving", employee.get("leaving_reason")),
        ]
        label_w = 48
        value_w = max(20, pdf.epw - label_w)
        for label, value in details:
            pdf.set_x(pdf.l_margin)
            pdf.set_font(family, "B", 9)
            pdf.cell(label_w, 6, safe(f"{label}:"))
            pdf.set_font(family, "", 9)
            pdf.multi_cell(value_w, 6, hard_wrap(value if value not in (None, "") else "-"), new_x="LMARGIN", new_y="NEXT")

        pdf.ln(3)
        full_line("HOLIDAY / LEAVE HISTORY", 7, True, 11)
        if history.get("types"):
            for typ, days in history["types"].items():
                full_line(f"{typ}: {days:.1f} day(s)", 6, False, 9)
        else:
            full_line("No approved leave records found.", 6, False, 9)

        pdf.ln(2)
        full_line(f"Approved annual holiday taken: {history.get('holiday_taken', 0.0):.1f} days", 6, True, 10)
        pdf.ln(3)
        full_line("FINAL HOLIDAY SETTLEMENTS", 7, True, 11)
        if history.get("settlements"):
            for r in history["settlements"]:
                try:
                    days = float(r.get("days", 0) or 0)
                except Exception:
                    days = 0.0
                try:
                    amount = float(r.get("amount", 0) or 0)
                except Exception:
                    amount = 0.0
                full_line(
                    f"Settlement #{r.get('id')} | {r.get('type','')} | {days:.1f} days | £{amount:.2f} | {str(r.get('status','')).title()} | {r.get('date','')}",
                    6, False, 9
                )
                if r.get("director_comments"):
                    full_line(f"Director comments: {r.get('director_comments')}", 5, False, 8)
                if r.get("rejection_reason"):
                    full_line(f"Rejection reason: {r.get('rejection_reason')}", 5, False, 8)
        else:
            full_line("No final holiday settlement records found.", 6, False, 9)

        pdf.ln(3)
        full_line("LEAVE RECORDS", 7, True, 11)
        if history.get("leave_records"):
            for r in history["leave_records"]:
                try:
                    days = float(r.get("days", 0) or 0)
                except Exception:
                    days = 0.0
                line = (
                    f"{r.get('date_from','')} → {r.get('date_to','')} | "
                    f"{r.get('type','')} | {days:.1f} days | {r.get('status','')} | {r.get('notes','')}"
                )
                full_line(line, 5, False, 8)
        else:
            full_line("No leave records found.", 5, False, 8)

        os.makedirs(PDF_DIR, exist_ok=True)
        safe_id = re.sub(r"[^A-Za-z0-9_-]+", "_", str(employee.get("emp_id", "leaver")))
        path = os.path.join(PDF_DIR, f"Leaver_History_{safe_id}.pdf")
        pdf.output(path)
        _upload_to_drive_bg(path, os.path.basename(path))
        return path
    except Exception as e:
        st.error(f"Leaver history PDF error: {type(e).__name__}: {e}")
        return None


def _hrp_render_leavers_tab():
    st.subheader("📚 Leavers")
    st.caption("Complete history for employees who have left the company. Leavers are removed from the live Holiday Calculator.")
    leavers = [e for e in st.session_state.get("hrp_employees", []) if str(e.get("status", "")).strip().casefold() == "left"]
    if not leavers:
        st.info("No employees have been recorded as Left.")
        return
    st.markdown("### Leaver List")
    list_rows = [{"Employee ID":e.get("emp_id"),"Full Name":e.get("name"),"Department":e.get("department"),"Start Date":e.get("start_date"),"Left Date":e.get("leaving_date"),"Reason":e.get("leaving_reason")} for e in leavers]
    st.dataframe(pd.DataFrame(list_rows), width="stretch", hide_index=True)
    options = [f"{e.get('name','')} — {e.get('emp_id','')}" for e in leavers]
    by_label = dict(zip(options, leavers))
    search = st.text_input("🔎 Search Leaver", placeholder="Search by name or ACE-ID / Employee ID...", key="hrp_leaver_search")
    filtered = [e for e in leavers if not search.strip() or search.casefold().strip() in (str(e.get('name',''))+' '+str(e.get('emp_id',''))).casefold()]
    if not filtered:
        st.warning("No leaver matches your search.")
        return
    labels = [f"{e.get('name','')} — {e.get('emp_id','')}" for e in filtered]
    selected_label = st.selectbox("Select Leaver", labels, key="hrp_selected_leaver")
    employee = next(e for e in filtered if f"{e.get('name','')} — {e.get('emp_id','')}" == selected_label)
    history = _hrp_leaver_history(employee)
    st.divider(); st.subheader("👤 Complete Employee History")
    c1,c2,c3 = st.columns(3)
    c1.write(f"**Employee ID:** {employee.get('emp_id','')}")
    c1.write(f"**Full Name:** {employee.get('name','')}")
    c1.write(f"**Department:** {employee.get('department','')}")
    c2.write(f"**Start Date:** {employee.get('start_date','')}")
    c2.write(f"**End Date:** {employee.get('leaving_date','')}")
    c2.write(f"**Duration:** {_hrp_work_duration(employee.get('start_date'), employee.get('leaving_date'))}")
    c3.write(f"**Position:** {employee.get('job_title','')}")
    c3.write(f"**Reason for Leaving:** {employee.get('leaving_reason','') or '-'}")
    c3.write(f"**Agreement:** {employee.get('agreement_type','')}")
    st.markdown("### Holidays / Leave Taken")
    if history["types"]:
        st.dataframe(pd.DataFrame([{"Leave Type":k,"Approved Days":v} for k,v in history["types"].items()]), width="stretch", hide_index=True)
    else: st.info("No approved leave records found.")
    st.markdown("### Final Holiday Settlement")
    try:
        leave_end = pd.to_datetime(employee.get("leaving_date")).date()
        leave_calc = _hrp_get_leaving_entitlement(employee, leave_end)
        leave_used = _hrp_get_approved_holiday_days_to_date(employee.get("emp_id"), leave_end)
        start_dt = pd.to_datetime(employee.get("start_date")).date()
        closure_days = _hrp_get_company_closure_holiday_days(employee, start_dt, leave_end)
        st.dataframe(pd.DataFrame([{
            "Pure Holiday Entitlement": leave_calc.get("gross", 0.0),
            "Bank Holidays": leave_calc.get("bank_holidays", 0.0),
            "Holiday Taken": round(leave_used + closure_days, 1),
            "Final Settlement Days": leave_calc.get("net", 0.0) - leave_used - closure_days - leave_calc.get("bank_holidays", 0.0),
        }]), width="stretch", hide_index=True)
    except Exception:
        pass
    if history["settlements"]:
        st.dataframe(pd.DataFrame([{"Settlement ID":r.get('id'),"Type":r.get('type'),"Days":float(r.get('days',0) or 0),"Amount (£)":float(r.get('amount',0) or 0),"Status":r.get('status','').title(),"Date":r.get('date')} for r in history["settlements"]]), width="stretch", hide_index=True)
    else: st.info("No final holiday settlement found.")
    if st.button("📄 Generate Complete Leaver History PDF", key=f"gen_leaver_pdf_{employee.get('emp_id')}", type="primary"):
        with st.spinner("Generating leaver history PDF..."):
            pdf_path = _hrp_leaver_history_pdf(employee, history)
        if pdf_path and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f: st.download_button("⬇️ Download Complete Leaver History PDF", f.read(), file_name=os.path.basename(pdf_path), mime="application/pdf", type="primary", key=f"leaver_pdf_{employee.get('emp_id')}")
        else:
            st.error("Could not generate the leaver history PDF.")


def _hrp_report_date(value):
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        parsed = pd.to_datetime(value, errors="coerce")
        return "" if pd.isna(parsed) else parsed.date()
    except Exception:
        return str(value)


def _hrp_report_cell(value):
    if value is None:
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return _hrp_report_date(value)
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return value


def _hrp_report_workbook_base(title, subtitle):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    ws.sheet_view.showGridLines = False

    navy = "1F4E78"
    blue = "D9EAF7"
    light = "F3F6F9"
    green = "E2F0D9"
    gold = "FFF2CC"
    white = "FFFFFF"
    grey = "666666"
    thin = Side(style="thin", color="D9E1F2")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:H1")
    ws["A1"] = title
    ws["A1"].font = Font(size=18, bold=True, color=white)
    ws["A1"].fill = PatternFill("solid", fgColor=navy)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:H2")
    ws["A2"] = subtitle
    ws["A2"].font = Font(size=11, italic=True, color=grey)
    ws["A2"].alignment = Alignment(horizontal="center")

    ws["A4"] = "Report Generated"
    ws["B4"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    ws["A5"] = "Prepared By"
    ws["B5"] = str(st.session_state.get("full_name", st.session_state.get("username", "HR Manager")))
    for cell in (ws["A4"], ws["A5"]):
        cell.font = Font(bold=True, color=navy)

    return wb, ws, {
        "navy": navy, "blue": blue, "light": light, "green": green,
        "gold": gold, "white": white, "grey": grey, "border": border,
        "thin": thin, "get_column_letter": get_column_letter,
    }


def _hrp_style_report_table(ws, start_row, start_col, headers, rows, styles, total_row=None):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    for offset, header in enumerate(headers):
        cell = ws.cell(start_row, start_col + offset, header)
        cell.font = Font(bold=True, color=styles["white"])
        cell.fill = PatternFill("solid", fgColor=styles["navy"])
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = styles["border"]

    for r_idx, row in enumerate(rows, start_row + 1):
        for c_idx, value in enumerate(row, start_col):
            cell = ws.cell(r_idx, c_idx, _hrp_report_cell(value))
            cell.border = styles["border"]
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if r_idx % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=styles["light"])

    if total_row is not None:
        r = start_row + 1 + len(rows)
        for c_idx, value in enumerate(total_row, start_col):
            cell = ws.cell(r, c_idx, _hrp_report_cell(value))
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor=styles["green"])
            cell.border = styles["border"]
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    ws.freeze_panes = ws.cell(start_row + 1, start_col)
    return start_row + 1 + len(rows) + (1 if total_row is not None else 0)


def _hrp_autofit_report(ws, min_width=10, max_width=28):
    from openpyxl.utils import get_column_letter
    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for cell in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=1, max_row=min(ws.max_row, 200)):
            for c in cell:
                value = "" if c.value is None else str(c.value)
                max_len = max(max_len, max((len(line) for line in value.split("\n")), default=0))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, min_width), max_width)


def _hrp_build_employee_report_xlsx():
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    import io

    _hr_portal_init()
    employees = _hrp_normalize_employee_records(st.session_state.get("hrp_employees", []))
    employees = sorted(employees, key=lambda e: (str(e.get("department", "")).casefold(), str(e.get("name", "")).casefold()))

    wb, ws, styles = _hrp_report_workbook_base(
        "COMPANY EMPLOYEE REPORT",
        "Department-wise Employee Master Report",
    )

    ws["A7"] = "EMPLOYEE SUMMARY"
    ws["A7"].font = Font(size=13, bold=True, color=styles["navy"])
    summary = [
        ("Total Employees", len(employees)),
        ("Active Employees", sum(str(e.get("status", "")).casefold() == "active" for e in employees)),
        ("Left Employees", sum(str(e.get("status", "")).casefold() == "left" for e in employees)),
        ("Departments", len({str(e.get("department", "")).strip() for e in employees if str(e.get("department", "")).strip()})),
    ]
    for i, (label, value) in enumerate(summary, 8):
        ws.cell(i, 1, label).font = Font(bold=True)
        ws.cell(i, 2, value)
        ws.cell(i, 1).fill = PatternFill("solid", fgColor=styles["blue"])
        ws.cell(i, 2).fill = PatternFill("solid", fgColor=styles["light"])
        ws.cell(i, 1).border = styles["border"]
        ws.cell(i, 2).border = styles["border"]

    headers = list(HR_EMPLOYEE_COLUMNS) + [
        "Calculated Holiday Entitlement", "Approved Holiday Used", "Holiday Balance"
    ]
    row = 14
    departments = []
    seen = set()
    for e in employees:
        dept = str(e.get("department", "")).strip() or "Unassigned"
        if dept not in seen:
            departments.append(dept); seen.add(dept)

    for dept in departments:
        dept_employees = [e for e in employees if (str(e.get("department", "")).strip() or "Unassigned") == dept]
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers))
        ws.cell(row, 1, f"DEPARTMENT: {dept}")
        ws.cell(row, 1).font = Font(size=13, bold=True, color=styles["white"])
        ws.cell(row, 1).fill = PatternFill("solid", fgColor=styles["navy"])
        ws.cell(row, 1).alignment = Alignment(vertical="center")
        row += 1

        rows = []
        for e in dept_employees:
            pos = _hrp_get_holiday_position(e["emp_id"])
            values = []
            source_map = {
                "Employee ID": e.get("emp_id", ""), "Full Name": e.get("name", ""),
                "Start Date": e.get("start_date", ""), "Position / Job Title": e.get("job_title", ""),
                "Department": e.get("department", ""), "Agreement Type": e.get("agreement_type", ""),
                "Status": e.get("status", ""), "Working Pattern": e.get("working_pattern", ""),
                "Days Worked Per Week": e.get("days_per_week", 5),
                "Holiday Entitlement Override": e.get("entitlement_override", ""),
                "Entitlement Adjustment Note": e.get("adjustment_note", ""),
                "Leaving Date": e.get("leaving_date", ""), "Leaving Reason": e.get("leaving_reason", ""),
            }
            for col in HR_EMPLOYEE_COLUMNS:
                values.append(source_map.get(col, ""))
            values.extend([pos.get("entitlement", 0), pos.get("used", 0), pos.get("balance", 0)])
            rows.append(values)

        total_used = round(sum(float(_hrp_get_holiday_position(e["emp_id"]).get("used", 0) or 0) for e in dept_employees), 1)
        total_balance = round(sum(float(_hrp_get_holiday_position(e["emp_id"]).get("balance", 0) or 0) for e in dept_employees), 1)
        total_row = ["Department Total", len(dept_employees)] + [""] * (len(headers) - 5) + ["", total_used, total_balance]
        row = _hrp_style_report_table(ws, row, 1, headers, rows, styles, total_row=total_row)
        row += 2

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers))
    ws.cell(row, 1, "COMPANY TOTAL")
    ws.cell(row, 1).font = Font(size=14, bold=True, color=styles["white"])
    ws.cell(row, 1).fill = PatternFill("solid", fgColor=styles["navy"])
    row += 1
    total_used = round(sum(float(_hrp_get_holiday_position(e["emp_id"]).get("used", 0) or 0) for e in employees), 1)
    total_balance = round(sum(float(_hrp_get_holiday_position(e["emp_id"]).get("balance", 0) or 0) for e in employees), 1)
    company_total = ["TOTAL EMPLOYEES", len(employees)] + [""] * (len(headers) - 5) + ["", total_used, total_balance]
    _hrp_style_report_table(ws, row, 1, ["Summary", "Total", "", "", "", "", ""], [["Total Employees", len(employees), "", "", "", "", ""]], styles)

    ws.sheet_view.zoomScale = 85
    _hrp_autofit_report(ws, min_width=10, max_width=32)
    ws.freeze_panes = "A15"
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def _hrp_build_holiday_report_xlsx(report_year):
    from openpyxl.styles import Font, PatternFill, Alignment
    import io

    _hr_portal_init()
    employees = _hrp_normalize_employee_records(st.session_state.get("hrp_employees", []))
    leave_records = _hrp_load_leave_records()
    employees = sorted(employees, key=lambda e: (str(e.get("department", "")).casefold(), str(e.get("name", "")).casefold()))

    wb, ws, styles = _hrp_report_workbook_base(
        "EMPLOYEE HOLIDAY REPORT",
        f"Department-wise Employee Holiday Report — {report_year}",
    )
    ws["A7"] = "HOLIDAY SUMMARY"
    ws["A7"].font = Font(size=13, bold=True, color=styles["navy"])

    holiday_types = set(HR_PORTAL_HOLIDAY_LEAVE_TYPES)
    approved_holiday_records = [
        r for r in leave_records
        if str(r.get("type", "")).strip() in holiday_types
        and str(r.get("status", "")).strip().casefold() == "approved"
        and _hrp_report_date(r.get("date_from")) != ""
        and _hrp_report_date(r.get("date_from")).year == int(report_year)
    ]
    total_holiday_days = round(sum(float(r.get("days", 0) or 0) for r in approved_holiday_records), 1)
    summary = [
        ("Total Employees", len(employees)),
        ("Employees With Approved Holiday", len({str(r.get("employee_id", "")).strip().casefold() for r in approved_holiday_records})),
        ("Approved Holiday Days", total_holiday_days),
        ("Departments", len({str(e.get("department", "")).strip() for e in employees if str(e.get("department", "")).strip()})),
    ]
    for i, (label, value) in enumerate(summary, 8):
        ws.cell(i, 1, label).font = Font(bold=True)
        ws.cell(i, 2, value)
        ws.cell(i, 1).fill = PatternFill("solid", fgColor=styles["blue"])
        ws.cell(i, 2).fill = PatternFill("solid", fgColor=styles["light"])
        ws.cell(i, 1).border = styles["border"]
        ws.cell(i, 2).border = styles["border"]

    row = 14
    departments = []
    seen = set()
    for e in employees:
        dept = str(e.get("department", "")).strip() or "Unassigned"
        if dept not in seen:
            departments.append(dept); seen.add(dept)

    employee_headers = ["Employee ID", "Employee Name", "Job Title", "Status", "Annual Entitlement", "Approved Holiday Used", "Holiday Balance"]
    detail_headers = ["Leave ID", "Date From", "Date To", "Holiday Type", "Days", "Status", "Notes"]

    for dept in departments:
        dept_employees = [e for e in employees if (str(e.get("department", "")).strip() or "Unassigned") == dept]
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        ws.cell(row, 1, f"DEPARTMENT: {dept}")
        ws.cell(row, 1).font = Font(size=13, bold=True, color=styles["white"])
        ws.cell(row, 1).fill = PatternFill("solid", fgColor=styles["navy"])
        row += 1

        summary_rows = []
        dept_days = 0.0
        for e in dept_employees:
            pos = _hrp_get_holiday_position(e["emp_id"])
            year_records = [
                r for r in leave_records
                if str(r.get("employee_id", "")).strip().casefold() == str(e.get("emp_id", "")).strip().casefold()
                and str(r.get("type", "")).strip() in holiday_types
                and str(r.get("status", "")).strip().casefold() == "approved"
                and _hrp_report_date(r.get("date_from")) != ""
                and _hrp_report_date(r.get("date_from")).year == int(report_year)
            ]
            used_year = round(sum(float(r.get("days", 0) or 0) for r in year_records), 1)
            dept_days += used_year
            summary_rows.append([
                e.get("emp_id", ""), e.get("name", ""), e.get("job_title", ""), e.get("status", ""),
                pos.get("entitlement", 0), used_year, pos.get("balance", 0),
            ])

        row = _hrp_style_report_table(
            ws, row, 1, employee_headers, summary_rows, styles,
            total_row=["Department Total", len(dept_employees), "", "", "", round(dept_days, 1), ""],
        )
        row += 1

        for e in dept_employees:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
            ws.cell(row, 1, f"{e.get('name', '')} — {e.get('emp_id', '')}")
            ws.cell(row, 1).font = Font(bold=True, color=styles["navy"])
            ws.cell(row, 1).fill = PatternFill("solid", fgColor=styles["gold"])
            row += 1
            records = [
                r for r in leave_records
                if str(r.get("employee_id", "")).strip().casefold() == str(e.get("emp_id", "")).strip().casefold()
                and str(r.get("type", "")).strip() in holiday_types
                and _hrp_report_date(r.get("date_from")) != ""
                and _hrp_report_date(r.get("date_from")).year == int(report_year)
            ]
            records.sort(key=lambda r: (_hrp_report_date(r.get("date_from")) or date.min))
            detail_rows = [[
                r.get("leave_id", ""), _hrp_report_date(r.get("date_from")), _hrp_report_date(r.get("date_to")),
                r.get("type", ""), r.get("days", 0), r.get("status", ""), r.get("notes", ""),
            ] for r in records]
            if detail_rows:
                row = _hrp_style_report_table(ws, row, 1, detail_headers, detail_rows, styles)
            else:
                ws.cell(row, 1, "No holiday records for this employee in the selected year.")
                ws.cell(row, 1).font = Font(italic=True, color=styles["grey"])
                row += 1
            row += 1

        row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    ws.cell(row, 1, "COMPANY TOTAL")
    ws.cell(row, 1).font = Font(size=14, bold=True, color=styles["white"])
    ws.cell(row, 1).fill = PatternFill("solid", fgColor=styles["navy"])
    row += 1
    _hrp_style_report_table(
        ws, row, 1,
        ["Summary", "Total", "", "", "", "", ""],
        [["Total Employees", len(employees), "", "", "", "", ""], ["Approved Holiday Days", total_holiday_days, "", "", "", "", ""]],
        styles,
    )

    ws.sheet_view.zoomScale = 90
    _hrp_autofit_report(ws, min_width=11, max_width=32)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def render_hr_download_reports():
    st.subheader("📥 HR Reports")
    st.caption("Download professional Excel reports grouped by department.")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 👥 Employee Master Report")
        st.caption("All employees, grouped by department, with complete employee fields and company total.")
        if st.button("🔄 Prepare Employee Report", key="hr_prepare_employee_report", type="secondary"):
            try:
                st.session_state.hrp_employee_report_bytes = _hrp_build_employee_report_xlsx()
                st.success("Employee report prepared.")
            except Exception as exc:
                st.error(f"Could not create the employee report: {exc}")
        if st.session_state.get("hrp_employee_report_bytes"):
            st.download_button(
                "📥 Download Employee Report",
                st.session_state.hrp_employee_report_bytes,
                file_name=f"Employee_Master_Report_{date.today().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="hr_download_employee_report",
                type="primary",
            )

    with c2:
        st.markdown("### 📅 Employee Holiday Report")
        st.caption("Department-wise employees, holiday allowance, holiday used, balance and individual holiday records.")
        report_year = st.number_input(
            "Holiday Year", min_value=2020, max_value=2100, value=date.today().year,
            step=1, key="hr_download_holiday_year"
        )
        if st.button("🔄 Prepare Holiday Report", key="hr_prepare_holiday_report", type="secondary"):
            try:
                st.session_state.hrp_holiday_report_bytes = _hrp_build_holiday_report_xlsx(int(report_year))
                st.session_state.hrp_holiday_report_year = int(report_year)
                st.success(f"Holiday report for {int(report_year)} prepared.")
            except Exception as exc:
                st.error(f"Could not create the holiday report: {exc}")
        if st.session_state.get("hrp_holiday_report_bytes"):
            prepared_year = st.session_state.get("hrp_holiday_report_year", int(report_year))
            st.download_button(
                "📥 Download Holiday Report",
                st.session_state.hrp_holiday_report_bytes,
                file_name=f"Employee_Holiday_Report_{prepared_year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="hr_download_holiday_report",
                type="primary",
            )

def render_hr_portal(current_user_info=None):
    """HR Portal — HR Management, Holiday Calendar, Employee Details, Leave and Leave History."""
    _hr_portal_init()
    if "hrp_role" not in st.session_state:
        st.session_state.hrp_role = "hr"
    if "hrp_current_emp_id" not in st.session_state:
        st.session_state.hrp_current_emp_id = ""
    if "hrp_new_employee_form_version" not in st.session_state:
        st.session_state.hrp_new_employee_form_version = 0
    if "hrp_leave_employee_id" not in st.session_state:
        st.session_state.hrp_leave_employee_id = ""

    st.subheader("🏢 HR Portal — Employee Management")
    st.caption("HR-managed employee, holiday and leave management system")
    st.divider()

    col_role1, col_role2 = st.columns([3, 1])
    with col_role2:
        selected_role = st.selectbox(
            "View As",
            ["hr", "emp"],
            index=0 if st.session_state.hrp_role == "hr" else 1,
            format_func=lambda v: "🔐 HR — Full Access" if v == "hr" else "👤 Employee — View Only",
            key="hrp_role_selector",
        )
        if selected_role != st.session_state.hrp_role:
            st.session_state.hrp_role = selected_role
            st.rerun()

    is_hr = st.session_state.hrp_role == "hr"
    current_role_name = str((current_user_info or {}).get("role", "")).strip().casefold() if isinstance(current_user_info, dict) else ""
    current_dept_name = str((current_user_info or {}).get("dept", "")).strip().casefold() if isinstance(current_user_info, dict) else ""
    is_hr_manager = current_dept_name == "hr" or current_role_name == "hr manager"

    active_employee_ids = [e["emp_id"] for e in st.session_state.hrp_employees if e.get("status") == "Active"]
    all_employee_ids = [e["emp_id"] for e in st.session_state.hrp_employees]
    if is_hr:
        if st.session_state.hrp_current_emp_id not in all_employee_ids:
            st.session_state.hrp_current_emp_id = active_employee_ids[0] if active_employee_ids else (all_employee_ids[0] if all_employee_ids else "")
    else:
        if not st.session_state.hrp_current_emp_id:
            st.session_state.hrp_current_emp_id = all_employee_ids[0] if all_employee_ids else ""

    selected_employee_id = st.session_state.hrp_current_emp_id
    emp = _hrp_get_employee(selected_employee_id) if selected_employee_id else None

    if is_hr:
        tab_hr, tab_calendar, tab_details, tab_leave, tab_history = st.tabs([
            "🧑‍💼 HR Management",
            "📅 Holiday Calendar",
            "👤 Employee Details",
            "✏️ Leave",
            "📋 Leave History",
        ])
    else:
        tab_details, tab_leave, tab_history = st.tabs([
            "👤 Employee Details",
            "✏️ Leave",
            "📋 Leave History",
        ])
        tab_hr = None
        tab_calendar = None

    if is_hr and tab_hr is not None:
        with tab_hr:
            st.subheader("🧑‍💼 HR Management")
            if is_hr_manager:
                st.caption("Employee Overview, Leave Approvals, Holiday Calculator, Employee Leaving and Employee Directory")
                hr_employee_tab, employee_overview_tab, hr_approval_tab, hr_holiday_calc_tab, hr_leaving_tab, hr_clearance_tab, hr_leavers_tab, employee_edit_tab, hr_reports_tab = st.tabs([
                    "👤 Employee Directory", "📊 Employee Overview", "✅ Leave Approvals", "📊 Holiday Calculator", "🚪 Employee Leaving", "🔗 Leaver Clearance", "📚 Leavers", "✏️ Edit / Deactivate Employee", "📥 HR Reports"
                ])
            else:
                st.caption("Employee Overview, Holiday Calculator, Employee Leaving and Employee Directory")
                hr_employee_tab, employee_overview_tab, hr_holiday_calc_tab, hr_leaving_tab, hr_clearance_tab, hr_leavers_tab, employee_edit_tab, hr_reports_tab = st.tabs([
                    "👤 Employee Directory", "📊 Employee Overview", "📊 Holiday Calculator", "🚪 Employee Leaving", "🔗 Leaver Clearance", "📚 Leavers", "✏️ Edit / Deactivate Employee", "📥 HR Reports"
                ])
                hr_approval_tab = None

            with employee_overview_tab:
                st.subheader("📊 Employee Overview — All Employees")
                if st.session_state.hrp_employees:
                    overview_rows = []
                    for e in st.session_state.hrp_employees:
                        pos = _hrp_get_holiday_position(e["emp_id"])
                        summary = _hrp_get_leave_summary(e["emp_id"])
                        overview_rows.append({
                            "Employee ID": e["emp_id"], "Name": e["name"], "Status": e.get("status", "Active"),
                            "Department": e.get("department", ""), "Position": e.get("job_title", ""),
                            "Start Date": e.get("start_date", ""), "Agreement": e.get("agreement_type", ""),
                            "Working Pattern": e.get("working_pattern", "Regular hours"), "Days/Week": e.get("days_per_week", 5),
                            "Holiday Entitlement": pos["entitlement"], "Bank Holidays (Separate)": _hrp_get_bank_holiday_days(e, e.get("start_date"), date.today()), "Holiday Used": pos["used"], "Holiday Balance": pos["balance"],
                            "Company Owes": pos["company_owes_employee"], "Employee Owes": pos["employee_owes_company"],
                            "Sick Days": summary["sick"], "Family / Emergency": summary["family"],
                            "Unpaid Days": summary["unpaid"], "Other Absence": summary["other"],
                        })
                    st.dataframe(pd.DataFrame(overview_rows), width="stretch", hide_index=True)
                else:
                    st.info("No employee records yet. Register your first employee below.")

            with hr_employee_tab:
                employee_directory_tab = st.container()
                with employee_directory_tab:
                    st.divider()
                    st.subheader("➕ Add New Employee")
                    st.info("HR enters the company Employee ID manually. Any unique letters/numbers format used by your company is accepted.")

                    new_form_version = st.session_state.hrp_new_employee_form_version
                    with st.form(f"hrp_new_employee_form_{new_form_version}", clear_on_submit=False):
                        col1, col2 = st.columns(2)
                        with col1:
                            new_emp_id = st.text_input(
                                "Employee ID",
                                placeholder="e.g. 001, EMP-001, A102 or your company batch number",
                                key=f"hrp_new_emp_id_{new_form_version}",
                            )
                            new_name = st.text_input(
                                "Full Name",
                                placeholder="e.g. John Smith",
                                key=f"hrp_new_name_{new_form_version}",
                            )
                            new_start_date = st.date_input(
                                "Start Date",
                                value=date.today(),
                                key=f"hrp_new_start_{new_form_version}",
                            )
                            new_position = st.text_input(
                                "Position / Job Title",
                                placeholder="e.g. Electrician",
                                key=f"hrp_new_pos_{new_form_version}",
                            )
                        with col2:
                            hr_software_departments = load_departments()
                            new_department = st.selectbox(
                                "Department",
                                options=hr_software_departments,
                                key=f"hrp_new_dept_{new_form_version}",
                            )
                            new_agreement = st.selectbox(
                                "Agreement Type",
                                options=HR_PORTAL_AGREEMENT_TYPES,
                                key=f"hrp_new_agree_{new_form_version}",
                            )
                            new_pattern = st.selectbox(
                                "Working Pattern",
                                ["Regular hours", "Irregular / Part-Year"],
                                key=f"hrp_new_pattern_{new_form_version}",
                            )
                            new_days_per_week = st.number_input(
                                "Contracted Days Per Week",
                                min_value=0.5,
                                max_value=7.0,
                                value=5.0,
                                step=0.5,
                                key=f"hrp_new_days_{new_form_version}",
                            )

                        create_employee = st.form_submit_button("➕ Create Employee", type="primary", width="stretch")

                    if create_employee:
                        valid, result = _hrp_validate_employee_id(new_emp_id)
                        if not valid:
                            st.error(result)
                        elif not new_name.strip():
                            st.error("Please enter the employee's full name.")
                        elif not new_position.strip():
                            st.error("Please enter the position / job title.")
                        else:
                            new_employee = {
                                "emp_id": result,
                                "name": new_name.strip(),
                                "start_date": new_start_date,
                                "department": new_department,
                                "job_title": new_position.strip(),
                                "agreement_type": new_agreement,
                                "status": "Active",
                                "working_pattern": new_pattern,
                                "days_per_week": new_days_per_week,
                                "entitlement_override": None,
                                "adjustment_note": "",
                            }
                            st.session_state.hrp_employees.append(new_employee)
                            _hrp_save_employees()
                            st.session_state.hrp_current_emp_id = result
                            log_action("HR_EMPLOYEE_ADDED", result, new_data=new_employee)
                            st.session_state.hrp_new_employee_form_version = new_form_version + 1
                            st.success(f"Employee {result} created successfully. The form is ready for the next employee.")
                            st.rerun()

                    st.subheader("Employee Directory")
                    st.caption("Edit employee details or permanently delete an employee record. Use Employee Leaving for all employee departures and final holiday settlement.")

                    employee_table = [
                        {
                            "Employee ID": e["emp_id"], "Name": e["name"], "Start Date": e["start_date"],
                            "Position": e["job_title"], "Department": e["department"],
                            "Agreement": e["agreement_type"], "Status": e["status"],
                        }
                        for e in st.session_state.hrp_employees
                    ]
                    if employee_table:
                        st.dataframe(pd.DataFrame(employee_table), width="stretch", hide_index=True)
                    else:
                        st.info("No employee records yet.")

            with employee_edit_tab:
                st.subheader("✏️ Edit / Deactivate Employee")
                st.caption("Edit employee details or set Active / Inactive status. Use Employee Leaving for employees who have left and their final holiday settlement.")
                edit_employee_ids = [e["emp_id"] for e in st.session_state.hrp_employees]
                if edit_employee_ids:
                    edit_emp_id = st.selectbox(
                        "Select Employee",
                        options=edit_employee_ids,
                        format_func=lambda eid: _hrp_employee_label(eid, include_status=True),
                        key="hrp_edit_emp_selector",
                    )
                    edit_emp = _hrp_get_employee(edit_emp_id)
                else:
                    edit_emp_id = ""
                    edit_emp = None

                if edit_emp:
                    with st.form(f"hrp_edit_employee_form_{edit_emp_id}"):
                        ec1, ec2 = st.columns(2)
                        with ec1:
                            edit_id = st.text_input("Employee ID", value=edit_emp["emp_id"], help="Must be ACE-ID followed by numbers, e.g. ACE-ID001.")
                            edit_name = st.text_input("Full Name", value=edit_emp["name"])
                            edit_start = st.date_input("Start Date", value=edit_emp["start_date"])
                            edit_position = st.text_input("Position / Job Title", value=edit_emp["job_title"])
                        with ec2:
                            hr_software_departments = load_departments()
                            edit_dept = st.selectbox(
                                "Department",
                                hr_software_departments,
                                index=hr_software_departments.index(edit_emp["department"]) if edit_emp["department"] in hr_software_departments else 0,
                            )
                            edit_agreement = st.selectbox(
                                "Agreement Type",
                                HR_PORTAL_AGREEMENT_TYPES,
                                index=HR_PORTAL_AGREEMENT_TYPES.index(edit_emp["agreement_type"]) if edit_emp["agreement_type"] in HR_PORTAL_AGREEMENT_TYPES else 0,
                            )
                            edit_pattern_options = ["Regular hours", "Irregular / Part-Year"]
                            edit_pattern = st.selectbox(
                                "Working Pattern",
                                edit_pattern_options,
                                index=edit_pattern_options.index(edit_emp.get("working_pattern", "Regular hours")) if edit_emp.get("working_pattern", "Regular hours") in edit_pattern_options else 0,
                            )
                            edit_days = st.number_input(
                                "Contracted Days Per Week",
                                min_value=0.5,
                                max_value=7.0,
                                value=float(edit_emp.get("days_per_week", 5) or 5),
                                step=0.5,
                            )
                            if edit_emp.get("status") == "Left":
                                edit_status = "Left"
                            else:
                                edit_status = st.selectbox(
                                    "Status",
                                    ["Active", "Inactive"],
                                    index=0 if edit_emp.get("status", "Active") == "Active" else 1,
                                )
                            edit_leaving_date = edit_emp.get("leaving_date")
                            edit_leaving_reason = edit_emp.get("leaving_reason", "")
                            if edit_emp.get("status") == "Left":
                                st.caption("This employee is recorded as Left. Use Employee Leaving to manage the leaving date and final holiday settlement.")
                        if st.form_submit_button("💾 Save Employee Changes", type="primary"):
                            old_id = str(edit_emp.get("emp_id", "")).strip()
                            valid_id, validated_id = _hrp_validate_employee_id(edit_id, exclude_id=old_id)
                            if not valid_id:
                                st.error(validated_id)
                                st.stop()
                            old_data = dict(edit_emp)
                            edit_emp["emp_id"] = validated_id
                            edit_emp["name"] = edit_name.strip()
                            edit_emp["start_date"] = edit_start
                            edit_emp["job_title"] = edit_position.strip()
                            edit_emp["department"] = edit_dept
                            edit_emp["agreement_type"] = edit_agreement
                            edit_emp["working_pattern"] = edit_pattern
                            edit_emp["days_per_week"] = edit_days
                            if edit_emp.get("status") != "Left":
                                edit_emp["status"] = edit_status
                            if validated_id != old_id:
                                users = load_users()
                                for u in users.values():
                                    if str(u.get("employee_id", "")).strip().casefold() == old_id.casefold():
                                        u["employee_id"] = validated_id
                                save_users(users)
                                hr_leave = load_hr_leave(force=True)
                                for r in hr_leave:
                                    if str(r.get("employee_id", "")).strip().casefold() == old_id.casefold():
                                        r["employee_id"] = validated_id
                                save_all_hr_leave(hr_leave)
                                portal_leave = _hrp_load_leave_records()
                                for r in portal_leave:
                                    if str(r.get("employee_id", "")).strip().casefold() == old_id.casefold():
                                        r["employee_id"] = validated_id
                                st.session_state.hrp_leave_records = portal_leave
                                _hrp_save_leave_records()
                            _hrp_save_employees()
                            st.session_state.hrp_current_emp_id = validated_id
                            log_action("HR_EMPLOYEE_EDITED", validated_id, old_data=old_data, new_data=dict(edit_emp))
                            st.success(f"Employee {validated_id} updated successfully.")
                            st.rerun()

                st.markdown("### 🗑️ Employee Record")
                st.caption("Use Active / Inactive here for normal employee management. Use Employee Leaving when an employee has actually left the company.")
                if edit_emp:
                    if st.button("🗑️ Permanently Delete Employee", key=f"hrp_delete_{edit_emp_id}", width="stretch"):
                        st.session_state[f"hrp_confirm_delete_{edit_emp_id}"] = True
                if edit_emp and st.session_state.get(f"hrp_confirm_delete_{edit_emp_id}", False):
                    st.warning("This permanently removes the employee record and their HR portal leave records. Use Employee Leaving when an employee leaves the company so the final holiday settlement is processed correctly.")
                    cc1, cc2 = st.columns(2)
                    with cc1:
                        if st.button("⚠️ Confirm Permanent Delete", key=f"hrp_confirm_delete_yes_{edit_emp_id}", type="primary", width="stretch"):
                            deleted = _hrp_get_employee(edit_emp_id)
                            st.session_state.hrp_employees = [e for e in st.session_state.hrp_employees if e["emp_id"] != edit_emp_id]
                            st.session_state.hrp_leave_records = [r for r in st.session_state.hrp_leave_records if r["employee_id"] != edit_emp_id]
                            st.session_state.hrp_entitlement_overrides.pop(edit_emp_id, None)
                            st.session_state.hrp_adjustment_notes.pop(edit_emp_id, None)
                            _hrp_save_employees()
                            _hrp_save_leave_records()
                            log_action("HR_EMPLOYEE_DELETED", edit_emp_id, old_data=deleted)
                            st.session_state.pop(f"hrp_confirm_delete_{edit_emp_id}", None)
                            if st.session_state.get("hrp_current_emp_id") == edit_emp_id:
                                active_ids = [e["emp_id"] for e in st.session_state.hrp_employees if e.get("status") == "Active"]
                                st.session_state.hrp_current_emp_id = active_ids[0] if active_ids else ""
                            st.success(f"Employee {edit_emp_id} permanently deleted.")
                            st.rerun()
                    with cc2:
                        if st.button("Cancel Delete", key=f"hrp_confirm_delete_no_{edit_emp_id}", width="stretch"):
                            st.session_state.pop(f"hrp_confirm_delete_{edit_emp_id}", None)
                            st.rerun()


            if is_hr_manager:
                with hr_approval_tab:
                    render_hr_leave_approvals()

            with hr_leaving_tab:
                st.subheader("🚪 Employee Leaving")
                st.info("Select an employee to review their full employment details, enter the leaving date and reason, then record the employee as Left. After the employee is recorded as Left, the final holiday balance is calculated and you can submit an Addition or Deduction settlement for Director approval.")

                current_leaving_id = str(st.session_state.get("hrp_leaving_employee", "") or "").strip()
                leaving_ids = [
                    e["emp_id"] for e in st.session_state.hrp_employees
                    if (
                        str(e.get("status", "")).strip().casefold() == "active"
                        or str(e.get("emp_id", "")).strip().casefold() == current_leaving_id.casefold()
                    )
                    and not _hrp_has_active_final_settlement(e.get("emp_id", ""))
                ]
                if not leaving_ids:
                    st.success("✅ All employees currently recorded for leaving already have a pending or approved final settlement. No duplicate settlement can be raised.")
                else:
                    if current_leaving_id in leaving_ids:
                        leaving_index = leaving_ids.index(current_leaving_id)
                    else:
                        leaving_index = 0
                    leaving_id = st.selectbox(
                        "Employee",
                        leaving_ids,
                        index=leaving_index,
                        format_func=lambda eid: _hrp_employee_label(eid, include_status=True),
                        key="hrp_leaving_employee",
                    )
                    leaving_emp = _hrp_get_employee(leaving_id)

                    if leaving_emp:
                        st.markdown("### 👤 Employee Details")
                        d1, d2, d3 = st.columns(3)
                        d1.write(f"**Employee ID:** {leaving_emp.get('emp_id', '')}")
                        d1.write(f"**Full Name:** {leaving_emp.get('name', '')}")
                        d1.write(f"**Department:** {leaving_emp.get('department', '')}")
                        d2.write(f"**Position:** {leaving_emp.get('job_title', '')}")
                        d2.write(f"**Start Date:** {leaving_emp.get('start_date', '')}")
                        d2.write(f"**Agreement:** {leaving_emp.get('agreement_type', '')}")
                        d3.write(f"**Working Pattern:** {leaving_emp.get('working_pattern', '')}")
                        d3.write(f"**Days Per Week:** {float(leaving_emp.get('days_per_week', 5) or 5):g}")
                        d3.write(f"**Current Status:** {leaving_emp.get('status', 'Active')}")

                        st.divider()
                        st.markdown("### 📝 Leaving Details")
                        default_leave_date = leaving_emp.get("leaving_date") or date.today()
                        if not isinstance(default_leave_date, date):
                            try:
                                default_leave_date = pd.to_datetime(default_leave_date).date()
                            except Exception:
                                default_leave_date = date.today()

                        with st.form(f"hrp_employee_leaving_form_{leaving_id}", clear_on_submit=False):
                            lc1, lc2 = st.columns(2)
                            with lc1:
                                leaving_date = st.date_input(
                                    "Leaving Date",
                                    value=default_leave_date,
                                    key=f"hrp_leaving_date_{leaving_id}",
                                )
                            with lc2:
                                leaving_reason = st.text_input(
                                    "Leaving Reason",
                                    value=str(leaving_emp.get("leaving_reason", "") or ""),
                                    key=f"hrp_leaving_reason_{leaving_id}",
                                    placeholder="e.g. Resignation, redundancy, end of contract",
                                )
                            record_left = st.form_submit_button(
                                "💾 Record Employee as Left",
                                type="primary",
                                width="stretch",
                            )

                        if record_left:
                            start_for_leave = leaving_emp.get("start_date")
                            if not isinstance(start_for_leave, date):
                                try:
                                    start_for_leave = pd.to_datetime(start_for_leave).date()
                                except Exception:
                                    start_for_leave = leaving_date
                            if leaving_date < start_for_leave:
                                st.error("Leaving Date cannot be before the employee Start Date.")
                            elif not leaving_reason.strip():
                                st.error("Please enter a Leaving Reason.")
                            else:
                                old_status = leaving_emp.get("status", "Active")
                                leaving_emp["status"] = "Left"
                                leaving_emp["leaving_date"] = leaving_date
                                leaving_emp["leaving_reason"] = leaving_reason.strip()
                                _hrp_save_employees()
                                _ensure_leaver_clearance(
                                    leaving_emp,
                                    str((current_user_info or {}).get("full_name") or (current_user_info or {}).get("username") or "HR").strip() or "HR",
                                )
                                log_action(
                                    "HR_EMPLOYEE_LEFT",
                                    leaving_id,
                                    old_data={"status": old_status},
                                    new_data={"status": "Left", "leaving_date": str(leaving_date), "leaving_reason": leaving_reason.strip()},
                                )
                                st.success(f"{leaving_emp['name']} has been recorded as Left on {leaving_date:%d/%m/%Y}. The final holiday settlement is now ready for the next step.")
                                st.rerun()

                        is_recorded_left = (
                            str(leaving_emp.get("status", "")).strip().casefold() == "left"
                            and bool(leaving_emp.get("leaving_date"))
                        )
                        if is_recorded_left:
                            saved_leaving_date = leaving_emp.get("leaving_date")
                            if not isinstance(saved_leaving_date, date):
                                try:
                                    saved_leaving_date = pd.to_datetime(saved_leaving_date).date()
                                except Exception:
                                    saved_leaving_date = leaving_date

                            st.divider()
                            st.subheader("💷 Step 2 — Final Holiday Settlement")
                            st.info("The employee is now recorded as Left. The calculation below is frozen at the recorded leaving date. If the company owes holiday, submit an Addition; if the employee has overused holiday, submit a Deduction. The request goes to Director approval.")

                            calc = _hrp_get_leaving_entitlement(leaving_emp, saved_leaving_date)
                            recorded_used_to_leave = _hrp_get_approved_holiday_days_to_date(leaving_id, saved_leaving_date)
                            employee_start_for_leaving = leaving_emp.get("start_date")
                            if not isinstance(employee_start_for_leaving, date):
                                try:
                                    employee_start_for_leaving = pd.to_datetime(employee_start_for_leaving).date()
                                except Exception:
                                    employee_start_for_leaving = saved_leaving_date
                            closure_to_leave = _hrp_get_company_closure_holiday_days(
                                leaving_emp, employee_start_for_leaving, saved_leaving_date
                            )
                            used_to_leave = round(recorded_used_to_leave + closure_to_leave, 1)
                            holiday_available_after_bank = round(calc["net"] - calc["bank_holidays"], 1)
                            balance = round(holiday_available_after_bank - used_to_leave, 1)

                            c1, c2, c3, c4 = st.columns(4)
                            c1.metric("Pure holiday entitlement", f"{calc['gross']:.1f} days")
                            c2.metric("Bank holidays", f"{calc['bank_holidays']:.1f} days")
                            c3.metric("Holiday available", f"{holiday_available_after_bank:.1f} days")
                            c4.metric("Holiday taken / pre-booked", f"{used_to_leave:.1f} days")

                            if balance > 0:
                                st.success(f"🏢 Company owes employee: {balance:.1f} holiday day(s).")
                                settlement_direction = "Company Owes Employee"
                                settlement_type = "Addition"
                                settlement_category = "Unused Holiday Payout"
                            elif balance < 0:
                                st.warning(f"👤 Employee owes company: {abs(balance):.1f} holiday day(s).")
                                settlement_direction = "Employee Owes Company"
                                settlement_type = "Deduction"
                                settlement_category = "Overused Holiday Deduction"
                            else:
                                st.success("✅ Final holiday position is exactly balanced at 0.0 days. No settlement is required.")
                                settlement_direction = None
                                settlement_type = None
                                settlement_category = None

                            st.caption(
                                f"Leaving date: {saved_leaving_date:%d/%m/%Y} · Pure holiday entitlement: {calc['gross']:.1f} days · "
                                f"Bank holidays: {calc['bank_holidays']:.1f} days · Holiday available after bank holidays: {holiday_available_after_bank:.1f} days · "
                                f"Holiday taken / pre-booked: {used_to_leave:.1f} days · Final settlement: {balance:.1f} days."
                            )

                            if settlement_direction:
                                if st.button(
                                    f"🧾 Submit {settlement_type} Settlement — {abs(balance):.1f} days for Director Approval",
                                    key=f"hrp_create_leaving_settlement_{leaving_id}",
                                    width="stretch",
                                    type="primary",
                                ):
                                    hr_records = load_hr_leave()
                                    existing_final = [
                                        r for r in hr_records
                                        if r.get("final_holiday_settlement", False)
                                        and str(r.get("employee_id", "")).strip().casefold() == str(leaving_id).strip().casefold()
                                        and str(r.get("status", "")).strip().casefold() in {"pending", "approved"}
                                    ]
                                    if existing_final:
                                        existing = sorted(existing_final, key=lambda r: int(r.get("id", 0) or 0), reverse=True)[0]
                                        st.warning(
                                            f"⚠️ A final holiday settlement already exists for {leaving_emp.get('name', leaving_id)} "
                                            f"(Settlement #{existing.get('id')}, {str(existing.get('status', '')).title()}). "
                                            "A duplicate settlement cannot be created."
                                        )
                                        st.stop()
                                    new_id = get_next_hr_leave_id(hr_records)
                                    dept = leaving_emp.get("department", "")
                                    rate = get_hr_daily_rate(dept, settlement_type)
                                    amount = round(rate * abs(balance), 2) if rate else 0.01
                                    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    _creator_info = current_user_info if isinstance(current_user_info, dict) else (st.session_state.get("user_info", {}) or {})
                                    creator_name = str(
                                        _creator_info.get("full_name")
                                        or _creator_info.get("name")
                                        or _creator_info.get("username")
                                        or "HR"
                                    ).strip()
                                    if not creator_name:
                                        creator_name = "HR"
                                    rec = {
                                        "id": new_id, "employee_id": leaving_emp.get("emp_id", ""), "emp_name": leaving_emp.get("name", ""), "emp_dept": dept,
                                        "type": settlement_type, "category": settlement_category, "owe_owed": settlement_direction,
                                        "date": str(saved_leaving_date), "days": abs(balance), "amount": amount,
                                        "manager": creator_name,
                                        "desc": f"Final holiday settlement for employee leaving {saved_leaving_date:%d/%m/%Y}. Pure holiday entitlement {calc['gross']:.1f}; bank holidays {calc['bank_holidays']:.1f}; holiday available after bank holidays {holiday_available_after_bank:.1f}; holiday taken/pre-booked {used_to_leave:.1f}; final settlement {balance:.1f}.",
                                        "attachment_name": "None", "status": "pending", "director_comments": "", "rejection_reason": "",
                                        "final_holiday_settlement": True,
                                        "decision_date": "", "decision_by": "", "submitted_by": creator_name, "submitted_date": now, "pdf_path": "",
                                    }
                                    hr_records.append(rec)
                                    save_all_hr_leave(hr_records)
                                    log_action("HR_LEAVE_CREATED", new_id, new_data=rec)
                                    st.session_state["hr_leave_submission_notice"] = (
                                        f"Settlement #{new_id} created as {settlement_type} and sent to Director for approval. "
                                        f"It is now available in HR Leave Settlement → My Submitted Settlements and in the Director's HR Leave Settlement → Pending requests."
                                    )
                                    st.session_state["hr_leave_last_created_id"] = new_id
                                    st.rerun()

            st.divider()
            st.subheader("📋 Submitted Final Holiday Settlements")
            final_records = [r for r in load_hr_leave(force=False) if _hrp_is_final_holiday_settlement_record(r)]
            if final_records:
                for r in final_records[:20]:
                    status = str(r.get("status", "pending")).strip().casefold()
                    icon = "🟡" if status == "pending" else ("🟢" if status == "approved" else "🔴")
                    with st.expander(f"{icon} Settlement #{r.get('id')} | {r.get('emp_name')} | {r.get('type')} | {float(r.get('days', 0) or 0):.1f} days | {status.upper()}"):
                        st.write(f"👤 **Employee:** {r.get('emp_name')} | 🆔 {r.get('employee_id')} | 🏢 {r.get('emp_dept')}")
                        st.write(f"🔄 **Type:** {r.get('type')} | 🔢 **Days:** {float(r.get('days', 0) or 0):.1f} | 💷 **Amount:** £{float(r.get('amount', 0) or 0):.2f}")
                        st.write(f"📅 **Leaving / Settlement Date:** {r.get('date')} | 📝 **Submitted by:** {r.get('submitted_by')}")
                        if r.get('director_comments'): st.info(f"💬 Director: {r.get('director_comments')}")
                        if r.get('rejection_reason'): st.error(f"❌ Rejection: {r.get('rejection_reason')}")
                        if status == "pending":
                            edit_key = f"hrp_edit_final_{r.get('id')}"
                            if st.button("✏️ Edit Pending Settlement", key=edit_key):
                                st.session_state[f"hrp_edit_final_open_{r.get('id')}"] = True
                            if st.session_state.get(f"hrp_edit_final_open_{r.get('id')}"):
                                new_manager = st.text_input("Line Manager", value=str(r.get('manager', '') or ''), key=f"hrp_final_mgr_{r.get('id')}")
                                new_amount = st.number_input("Amount (£)", min_value=0.01, value=float(r.get('amount', 0.01) or 0.01), step=1.0, key=f"hrp_final_amt_{r.get('id')}")
                                new_desc = st.text_area("Description / Justification", value=str(r.get('desc', '') or ''), key=f"hrp_final_desc_{r.get('id')}")
                                ec1, ec2 = st.columns(2)
                                with ec1:
                                    if st.button("💾 Save Changes", key=f"hrp_save_final_{r.get('id')}", type="primary", width="stretch"):
                                        records = load_hr_leave(force=True)
                                        for rr in records:
                                            if int(rr.get('id', 0) or 0) == int(r.get('id', 0) or 0):
                                                rr['manager'] = new_manager.strip()
                                                rr['amount'] = float(new_amount)
                                                rr['desc'] = new_desc.strip()
                                                break
                                        save_all_hr_leave(records)
                                        st.session_state.pop(f"hrp_edit_final_open_{r.get('id')}", None)
                                        st.success(f"Settlement #{r.get('id')} updated.")
                                        st.rerun()
                                with ec2:
                                    if st.button("Cancel", key=f"hrp_cancel_final_{r.get('id')}", width="stretch"):
                                        st.session_state.pop(f"hrp_edit_final_open_{r.get('id')}", None)
                                        st.rerun()
            else:
                st.info("No final holiday settlements have been submitted yet.")

            with hr_clearance_tab:
                render_leaver_clearance_hr()

            with hr_leavers_tab:
                _hrp_render_leavers_tab()

            with hr_holiday_calc_tab:
                st.session_state.hrp_leave_records = _hrp_load_leave_records()
                st.subheader("📊 Holiday Calculator")
                st.info("Calculate an employee's current holiday position. Employees recorded as Left are closed and show 0.0 days in the live holiday position. Final leaving settlements are handled separately in Employee Leaving and approved by the Director.")
                settlement_ids = [e["emp_id"] for e in st.session_state.hrp_employees if str(e.get("status", "")).strip().casefold() == "active"]
                if settlement_ids:
                    settlement_employee_id = st.selectbox(
                        "Employee",
                        options=settlement_ids,
                        format_func=lambda eid: _hrp_employee_label(eid),
                        key="hrp_settlement_emp",
                    )
                    settlement_employee = _hrp_get_employee(settlement_employee_id)
                else:
                    settlement_employee = None
                if settlement_employee:
                    entitlement, entitlement_note, _, _ = _hrp_get_employee_entitlement(settlement_employee)
                    holiday_used = _hrp_get_approved_holiday_days(settlement_employee["emp_id"])
                    remaining = entitlement - holiday_used
                    holiday_position = _hrp_get_holiday_position(settlement_employee["emp_id"])
                    st.divider()
                    col1, col2, col3 = st.columns(3)
                    with col1: st.metric("Holiday Entitlement", f"{entitlement:.1f} days")
                    with col2: st.metric("Approved Holiday Used", f"{holiday_used:.1f} days")
                    with col3: st.metric("Balance", f"{holiday_position['balance']:.1f} days")
                    st.caption(entitlement_note)
                    owed1, owed2 = st.columns(2)
                    with owed1: st.metric("🏢 Company Owes Employee", f"{holiday_position['company_owes_employee']:.1f} days")
                    with owed2: st.metric("👤 Employee Owes Company", f"{holiday_position['employee_owes_company']:.1f} days")
                    st.divider()
                    summary = _hrp_get_leave_summary(settlement_employee["emp_id"])
                    st.subheader("Leave Settlement Summary")
                    settlement_data = [
                        {"Leave Category": "Annual Holiday", "Approved Days": summary["holiday"], "Affects Holiday Balance": "Yes"},
                        {"Leave Category": "Sick Leave", "Approved Days": summary["sick"], "Affects Holiday Balance": "No"},
                        {"Leave Category": "Unpaid Leave", "Approved Days": summary["unpaid"], "Affects Holiday Balance": "No"},
                        {"Leave Category": "Maternity Leave", "Approved Days": summary["maternity"], "Affects Holiday Balance": "No"},
                        {"Leave Category": "Paternity Leave", "Approved Days": summary["paternity"], "Affects Holiday Balance": "No"},
                    ]
                    st.dataframe(pd.DataFrame(settlement_data), width="stretch", hide_index=True)
                    st.divider()
                    st.subheader("Settlement Calculation")
                    st.write(f"**Employee:** {settlement_employee['name']}")
                    st.write(f"**Employee ID:** {settlement_employee['emp_id']}")
                    st.write(f"**Entitlement:** {entitlement:.1f} days")
                    st.write(f"**Approved Holiday Used:** {holiday_used:.1f} days")
                    st.write(f"**Live Holiday Balance:** {holiday_position['balance']:.1f} days")
                    if holiday_position['balance'] < 0:
                        st.warning(f"Employee is {abs(holiday_position['balance']):.1f} days over the current entitlement.")
                    else:
                        st.success(f"{holiday_position['balance']:.1f} days available.")
                else:
                    st.info("No employees are registered yet.")

    if is_hr and tab_calendar is not None:
        with tab_calendar:
            _hrp_render_holiday_calendar()

    with tab_details:
        st.subheader("Employee Details")

        if is_hr:
            if all_employee_ids:
                selector_options = active_employee_ids or all_employee_ids
                current_index = selector_options.index(st.session_state.hrp_current_emp_id) if st.session_state.hrp_current_emp_id in selector_options else 0
                selected_from_dropdown = st.selectbox(
                    "Current Employee",
                    options=selector_options,
                    index=current_index,
                    format_func=lambda eid: _hrp_employee_label(eid),
                    key="hrp_current_employee_details_selector",
                )
                if selected_from_dropdown != st.session_state.hrp_current_emp_id:
                    st.session_state.hrp_current_emp_id = selected_from_dropdown
                    st.rerun()
                emp = _hrp_get_employee(st.session_state.hrp_current_emp_id)
            else:
                emp = None
                st.info("No employees are registered yet. Go to **HR Management → Employee Management** to create the first employee.")
        else:
            emp = _hrp_get_employee(st.session_state.hrp_current_emp_id) if st.session_state.hrp_current_emp_id else None

        if emp:
            st.info("HR can update employee information." if is_hr else "Your employee information is view-only. Contact HR if any details need to be changed.")
            disabled = not is_hr
            col1, col2, col3 = st.columns(3)
            with col1:
                name = st.text_input("Full Name", value=emp["name"], disabled=disabled, key=f"hrp_emp_name_{emp['emp_id']}")
            with col2:
                start_date = st.date_input("Start Date", value=emp["start_date"], disabled=disabled, key=f"hrp_emp_start_{emp['emp_id']}")
            with col3:
                hr_software_departments = load_departments()
                department = st.selectbox(
                    "Department",
                    options=hr_software_departments,
                    index=hr_software_departments.index(emp["department"]) if emp["department"] in hr_software_departments else 0,
                    disabled=disabled,
                    key=f"hrp_emp_dept_{emp['emp_id']}",
                )
            col1, col2, col3 = st.columns(3)
            with col1:
                st.text_input("Employee ID", value=emp["emp_id"], disabled=True, key=f"hrp_emp_id_display_{emp['emp_id']}")
            with col2:
                job_title = st.text_input("Position / Job Title", value=emp["job_title"], disabled=disabled, key=f"hrp_emp_job_{emp['emp_id']}")
            with col3:
                agreement_type = st.selectbox(
                    "Agreement Type",
                    options=HR_PORTAL_AGREEMENT_TYPES,
                    index=HR_PORTAL_AGREEMENT_TYPES.index(emp["agreement_type"]) if emp["agreement_type"] in HR_PORTAL_AGREEMENT_TYPES else 0,
                    disabled=disabled,
                    key=f"hrp_emp_agreement_{emp['emp_id']}",
                )
            st.write("")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.text_input("Status", value=emp["status"], disabled=True, key=f"hrp_emp_status_display_{emp['emp_id']}")
            with col2:
                st.text_input("Working Pattern", value=emp.get("working_pattern", "Regular hours"), disabled=True, key=f"hrp_emp_pattern_display_{emp['emp_id']}")
            with col3:
                st.number_input("Contracted Days Per Week", min_value=0.5, max_value=7.0, value=float(emp.get("days_per_week", 5) or 5), step=0.5, disabled=True, key=f"hrp_emp_days_display_{emp['emp_id']}")

            if is_hr:
                if st.button("💾 Save Employee Details", type="primary", key=f"hrp_save_emp_{emp['emp_id']}"):
                    old_data = {
                        "name": emp["name"], "start_date": str(emp["start_date"]), "department": emp["department"],
                        "job_title": emp["job_title"], "agreement_type": emp["agreement_type"],
                    }
                    emp["name"] = name.strip()
                    emp["start_date"] = start_date
                    emp["department"] = department
                    emp["job_title"] = job_title.strip()
                    emp["agreement_type"] = agreement_type
                    new_data = {
                        "name": emp["name"], "start_date": str(emp["start_date"]), "department": emp["department"],
                        "job_title": emp["job_title"], "agreement_type": emp["agreement_type"],
                    }
                    _hrp_save_employees()
                    log_action("HR_EMPLOYEE_EDITED", emp["emp_id"], old_data=old_data, new_data=new_data)
                    st.success(f"Employee {emp['emp_id']} updated successfully.")
                    st.rerun()

            st.divider()
            st.subheader("📅 Holiday Allowance")
            entitlement, entitlement_note, service_years, calculated_base = _hrp_get_employee_entitlement(emp)
            holiday_position = _hrp_get_holiday_position(emp["emp_id"])
            holiday_used = holiday_position["used"]
            bank_holidays_passed = holiday_position.get("bank_holidays_passed", 0.0)
            upcoming_bank_holidays = holiday_position.get("upcoming_bank_holidays", 0.0)
            remaining = holiday_position["balance"]

            col1, col2, col3, col4, col5, col6 = st.columns(6)
            with col1: st.metric("Holiday Entitlement", f"{entitlement:.1f} days")
            with col2: st.metric("Holiday Used / Booked", f"{holiday_used:.1f} days")
            with col3: st.metric("Bank Holidays Passed", f"{bank_holidays_passed:.1f} days")
            with col4: st.metric("Upcoming Bank Holidays", f"{upcoming_bank_holidays:.1f} days")
            with col5: st.metric("Remaining", f"{remaining:.1f} days")
            with col6: st.metric("Service", f"{service_years:.1f} years")

            st.info(entitlement_note)

            today = date.today()
            try:
                emp_start = emp.get("start_date")
                if not isinstance(emp_start, date):
                    emp_start = pd.to_datetime(emp_start).date()
            except Exception:
                emp_start = today
            leave_year_end = date(today.year, 12, 31)
            if emp.get("leaving_date"):
                try:
                    leaving = emp.get("leaving_date")
                    if not isinstance(leaving, date):
                        leaving = pd.to_datetime(leaving).date()
                    leave_year_end = min(leave_year_end, leaving)
                except Exception:
                    pass
            upcoming_bank_list = []
            if leave_year_end >= today + timedelta(days=1):
                for bank_date, bank_name in _hrp_get_upcoming_bank_holidays(today + timedelta(days=1), limit=20):
                    if bank_date <= leave_year_end and bank_date >= max(today + timedelta(days=1), emp_start):
                        upcoming_bank_list.append((bank_date, bank_name))

            st.caption(
                f"Holiday balance calculation: {holiday_used:.1f} booked/pre-booked holiday day(s) "
                f"+ {bank_holidays_passed:.1f} bank holiday day(s) passed "
                f"+ {upcoming_bank_holidays:.1f} upcoming bank holiday day(s) "
                f"= {remaining:.1f} day(s) remaining from {entitlement:.1f} day(s) entitlement."
            )
            if upcoming_bank_list:
                with st.expander("📅 Upcoming Bank Holidays included in Remaining balance", expanded=False):
                    st.dataframe(
                        pd.DataFrame([
                            {"Date": bank_date.strftime("%d/%m/%Y"), "Bank Holiday": bank_name, "Days Reserved": 1.0}
                            for bank_date, bank_name in upcoming_bank_list
                        ]),
                        width="stretch",
                        hide_index=True,
                    )

            owed1, owed2 = st.columns(2)
            with owed1:
                st.metric("🏢 Company Owes Employee", f"{holiday_position['company_owes_employee']:.1f} days")
            with owed2:
                st.metric("👤 Employee Owes Company", f"{holiday_position['employee_owes_company']:.1f} days")

            if is_hr:
                st.divider()
                st.subheader("⚙️ HR Entitlement Adjustment")
                st.caption("Use this when the employee's contractual entitlement differs from the system calculation.")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("System Calculated", f"{calculated_base:.1f} days")
                with col2:
                    adjusted_value = st.number_input("Final Entitlement", min_value=0.0, max_value=50.0, step=0.5, value=float(entitlement), key=f"hrp_adj_ent_{emp['emp_id']}")
                with col3:
                    adjustment_reason = st.text_input("Reason", value=st.session_state.hrp_adjustment_notes.get(emp["emp_id"], ""), placeholder="Reason for adjustment", key=f"hrp_adj_reason_{emp['emp_id']}")
                if st.button("✅ Save Entitlement", type="primary", key=f"hrp_save_ent_{emp['emp_id']}"):
                    old_data = {"entitlement": entitlement}
                    st.session_state.hrp_entitlement_overrides[emp["emp_id"]] = adjusted_value
                    st.session_state.hrp_adjustment_notes[emp["emp_id"]] = adjustment_reason
                    emp["entitlement_override"] = adjusted_value
                    emp["adjustment_note"] = adjustment_reason
                    _hrp_save_employees()
                    log_action("HR_ENTITLEMENT_ADJUSTED", emp["emp_id"], old_data=old_data, new_data={"entitlement": adjusted_value, "reason": adjustment_reason})
                    st.success(f"Holiday entitlement for {emp['emp_id']} is now {adjusted_value:.1f} days.")
                    st.rerun()
                if emp["emp_id"] in st.session_state.hrp_adjustment_notes:
                    note = st.session_state.hrp_adjustment_notes[emp["emp_id"]]
                    if note:
                        st.caption(f"HR Adjustment Note: {note}")
        elif not is_hr:
            st.info("Your employee record could not be found. Please contact HR.")

    with tab_leave:
        st.subheader("Leave Management")
        if not emp:
            st.info("Select an employee in **Employee Details** first." if is_hr else "No employee record is available.")
        elif is_hr:
            st.info("HR records leave on behalf of the employee. All leave submitted by HR is automatically marked Approved.")

            leave_employee_ids = [
                e["emp_id"] for e in st.session_state.hrp_employees
                if e.get("status", "Active") == "Active"
            ]
            if not leave_employee_ids:
                st.warning("There are no active employees available to book leave for.")
                st.stop()

            leave_current_id = (
                st.session_state.hrp_leave_employee_id
                if st.session_state.get("hrp_leave_employee_id") in leave_employee_ids
                else (
                    st.session_state.hrp_current_emp_id
                    if st.session_state.get("hrp_current_emp_id") in leave_employee_ids
                    else leave_employee_ids[0]
                )
            )
            leave_current_index = leave_employee_ids.index(leave_current_id)

            selected_leave_employee_id = st.selectbox(
                "👤 Employee — Book Leave For",
                options=leave_employee_ids,
                index=leave_current_index,
                format_func=lambda eid: _hrp_employee_label(eid),
                key="hrp_leave_employee_selector",
            )
            st.session_state.hrp_leave_employee_id = selected_leave_employee_id
            leave_employee = _hrp_get_employee(selected_leave_employee_id)

            if leave_employee:
                st.caption(
                    f"Booking leave for **{leave_employee['name']}** "
                    f"(Employee ID: **{leave_employee['emp_id']}**) · "
                    f"{leave_employee.get('department', '')} · "
                    f"{leave_employee.get('job_title', '')}"
                )

            form_version = st.session_state.get("hrp_leave_form_version", 0)
            col1, col2 = st.columns(2)
            with col1:
                leave_type = st.selectbox("Leave Type", options=HR_PORTAL_LEAVE_TYPES, key=f"hrp_leave_type_{form_version}")
                leave_start = st.date_input("Start Date", value=date.today(), key=f"hrp_leave_start_{form_version}")
            with col2:
                leave_end = st.date_input("End Date", value=date.today(), key=f"hrp_leave_end_{form_version}")
                half_day = st.checkbox("Half Day", disabled=leave_type != "Full Day Holiday", key=f"hrp_half_day_{form_version}")
            request_source = st.selectbox("Request / Notification Source", ["Email", "Phone Call", "In Person", "Other"], key=f"hrp_request_source_{form_version}")
            request_reference = st.text_input("Email / Call Reference", placeholder="Optional email subject, date, or reference", key=f"hrp_request_reference_{form_version}")
            notes = st.text_area("Notes / Reason", placeholder="e.g. holiday requested by email; sickness reported by phone; family emergency", key=f"hrp_leave_notes_{form_version}")
            if leave_end < leave_start:
                st.error("End date cannot be before the start date.")
            else:
                calculated_days = _hrp_calculate_leave_days(leave_type, leave_start, leave_end, half_day)
                st.metric("Calculated Leave Days", f"{calculated_days:.1f}")
                st.caption("Weekends are excluded from the calculation.")
                if st.button("📤 Record Leave — Automatically Approved", type="primary", key=f"hrp_record_leave_{form_version}"):
                    blocked_dates = _hrp_blocked_leave_dates(leave_start, leave_end)
                    if blocked_dates:
                        details = ", ".join(f"{d.strftime('%d/%m/%Y')} ({code} — {reason})" for d, code, reason in blocked_dates)
                        st.error(f"Leave cannot be booked on company non-working dates: {details}")
                    elif calculated_days <= 0:
                        st.error("The selected dates do not contain any working days.")
                    else:
                        persistent_leave_records = _hrp_load_leave_records()
                        st.session_state.hrp_leave_records = persistent_leave_records
                        employee_key = str(leave_employee["emp_id"]).strip().casefold()
                        duplicate = next((r for r in persistent_leave_records
                                          if str(r.get("employee_id", "")).strip().casefold() == employee_key
                                          and r.get("date_from") == leave_start
                                          and r.get("date_to") == leave_end
                                          and r.get("type") == leave_type
                                          and float(r.get("days", 0) or 0) == float(calculated_days)
                                          and str(r.get("request_reference", "")).strip() == request_reference.strip()
                                          and str(r.get("notes", "")).strip() == notes.strip()), None)
                        if duplicate:
                            st.warning(f"This leave has already been recorded as {duplicate['leave_id']}. The duplicate was not added.")
                        else:
                            new_record = {
                                "leave_id": _hrp_create_leave_id(), "employee_id": leave_employee["emp_id"],
                                "date_from": leave_start, "date_to": leave_end, "type": leave_type,
                                "days": calculated_days, "status": "Approved", "request_source": request_source,
                                "request_reference": request_reference.strip(), "notes": notes.strip(),
                                "recorded_by": (current_user_info or {}).get("full_name", "HR") if isinstance(current_user_info, dict) else "HR",
                                "recorded_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "entry_source": "HR Direct",
                                "requested_by": "", "requested_at": "",
                                "approved_by": (current_user_info or {}).get("full_name", "HR") if isinstance(current_user_info, dict) else "HR",
                                "approved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "rejection_reason": "",
                            }
                            with _HR_PORTAL_LEAVE_LOCK:
                                latest = _hrp_load_leave_records()
                                latest_ids = {str(r.get("leave_id", "")).strip() for r in latest}
                                if new_record["leave_id"] in latest_ids:
                                    st.warning(f"This leave has already been recorded as {new_record['leave_id']}.")
                                else:
                                    latest.append(new_record)
                                    _hrp_save_leave_records(latest, sync_drive=True)
                            st.session_state.hrp_leave_records = _hrp_load_leave_records()
                            if not any(str(r.get("leave_id", "")) == new_record["leave_id"] for r in st.session_state.hrp_leave_records):
                                raise RuntimeError(f"Leave {new_record['leave_id']} was not present after save verification.")
                            log_action("HR_LEAVE_RECORDED", new_record["leave_id"], new_data=new_record)
                            st.session_state.hrp_leave_form_version = form_version + 1
                            st.success(f"Leave recorded successfully for {leave_employee['name']}. {calculated_days:.1f} day(s) — Approved. The form has been cleared for the next entry.")
                            st.rerun()
        else:
            st.info("You are viewing your leave information. Leave is managed by HR.")
            employee_leave = _hrp_get_employee_leave(emp["emp_id"])
            if employee_leave:
                display_records = [{"Leave ID": r["leave_id"], "Date From": r["date_from"], "Date To": r["date_to"], "Leave Type": r["type"], "Days": r["days"], "Status": r["status"], "Notes": r["notes"]} for r in employee_leave]
                st.dataframe(pd.DataFrame(display_records), width="stretch", hide_index=True)
            else:
                st.write("No leave records found.")

    with tab_history:
        st.subheader("Leave History")

        if is_hr:
            st.session_state.hrp_leave_records = _hrp_load_leave_records()
            history_employee_ids = [
                e["emp_id"] for e in st.session_state.hrp_employees
                if e.get("status", "Active") == "Active"
            ]
            for e in st.session_state.hrp_employees:
                if e.get("emp_id") not in history_employee_ids:
                    if _hrp_get_employee_leave(e.get("emp_id")):
                        history_employee_ids.append(e.get("emp_id"))

            if not history_employee_ids:
                st.info("No employee records are available.")
            else:
                saved_history_id = st.session_state.get("hrp_history_employee_id")
                if saved_history_id not in history_employee_ids:
                    saved_history_id = history_employee_ids[0]

                history_index = history_employee_ids.index(saved_history_id)
                selected_history_employee_id = st.selectbox(
                    "👤 Employee — View Leave History For",
                    options=history_employee_ids,
                    index=history_index,
                    format_func=lambda eid: (
                        f"{_hrp_get_employee(eid)['name']} — {eid}"
                        if _hrp_get_employee(eid) else eid
                    ),
                    key="hrp_history_employee_selector",
                )
                st.session_state.hrp_history_employee_id = selected_history_employee_id
                history_employee = _hrp_get_employee(selected_history_employee_id)

                if history_employee:
                    st.caption(
                        f"Viewing leave history for **{history_employee['name']}** "
                        f"(Employee ID: **{history_employee['emp_id']}**) · "
                        f"{history_employee.get('department', '')} · "
                        f"{history_employee.get('job_title', '')}"
                    )

                employee_leave = [r for r in _hrp_get_employee_leave(selected_history_employee_id) if r.get("status") == "Approved"]
                if employee_leave:
                    history_data = [{
                        "Leave ID": r["leave_id"], "Date From": r["date_from"],
                        "Date To": r["date_to"], "Type": r["type"], "Days": r["days"],
                        "Status": r["status"], "Notes": r["notes"]
                    } for r in employee_leave]
                    df_history = pd.DataFrame(history_data)
                    st.dataframe(df_history, width="stretch", hide_index=True)

                    st.divider()
                    st.subheader("✏️ Edit / 🗑️ Delete Leave Record")
                    st.caption("HR can correct a wrongly entered leave record or remove an accidental duplicate. Deleted records are removed from the HR leave ledger.")
                    leave_choices = [r["leave_id"] for r in employee_leave]
                    selected_leave_id = st.selectbox(
                        "Select Leave Record",
                        leave_choices,
                        key=f"hrp_leave_action_selector_{selected_history_employee_id}"
                    )
                    selected_leave = next((r for r in employee_leave if r["leave_id"] == selected_leave_id), None)
                    if selected_leave:
                        with st.form(f"hrp_edit_leave_form_{selected_history_employee_id}_{selected_leave_id}"):
                            ec1, ec2 = st.columns(2)
                            with ec1:
                                edit_leave_type = st.selectbox(
                                    "Leave Type", HR_PORTAL_LEAVE_TYPES,
                                    index=HR_PORTAL_LEAVE_TYPES.index(selected_leave["type"]) if selected_leave["type"] in HR_PORTAL_LEAVE_TYPES else 0
                                )
                                edit_leave_start = st.date_input("Start Date", value=selected_leave["date_from"])
                                edit_leave_source = st.selectbox(
                                    "Request / Notification Source",
                                    ["Email", "Phone Call", "In Person", "Other"],
                                    index=["Email", "Phone Call", "In Person", "Other"].index(selected_leave.get("request_source", "Email")) if selected_leave.get("request_source", "Email") in ["Email", "Phone Call", "In Person", "Other"] else 0
                                )
                            with ec2:
                                edit_leave_end = st.date_input("End Date", value=selected_leave["date_to"])
                                edit_leave_reference = st.text_input("Email / Call Reference", value=selected_leave.get("request_reference", ""))
                                edit_leave_notes = st.text_area("Notes / Reason", value=selected_leave.get("notes", ""))
                            edit_half_day = edit_leave_type == "Half Day Holiday"
                            edit_days = _hrp_calculate_leave_days(edit_leave_type, edit_leave_start, edit_leave_end, edit_half_day)
                            st.caption(f"Calculated days: {edit_days:.1f}. HR edits remain Approved automatically.")
                            if st.form_submit_button("💾 Save Leave Changes", type="primary"):
                                if edit_leave_end < edit_leave_start or edit_days <= 0:
                                    st.error("Please select a valid working-day date range.")
                                else:
                                    old_leave = dict(selected_leave)
                                    selected_leave.update({
                                        "date_from": edit_leave_start, "date_to": edit_leave_end, "type": edit_leave_type,
                                        "days": edit_days, "status": "Approved", "request_source": edit_leave_source,
                                        "request_reference": edit_leave_reference.strip(), "notes": edit_leave_notes.strip(),
                                    })
                                    _hrp_save_leave_records()
                                    log_action("HR_LEAVE_EDITED", selected_leave_id, old_data=old_leave, new_data=dict(selected_leave))
                                    st.success(f"Leave record {selected_leave_id} updated.")
                                    st.rerun()

                        delete_confirm_key = f"hrp_confirm_leave_delete_{selected_history_employee_id}_{selected_leave_id}"
                        if st.button("🗑️ Delete This Leave Record", key=f"hrp_delete_leave_{selected_history_employee_id}_{selected_leave_id}"):
                            st.session_state[delete_confirm_key] = True
                        if st.session_state.get(delete_confirm_key):
                            st.warning(f"This will permanently delete leave record {selected_leave_id}. This is appropriate for an accidental duplicate or incorrect entry.")
                            cdel1, cdel2 = st.columns(2)
                            with cdel1:
                                if st.button("⚠️ Confirm Delete", type="primary", key=f"hrp_confirm_leave_delete_yes_{selected_history_employee_id}_{selected_leave_id}"):
                                    old_leave = dict(selected_leave)
                                    st.session_state.hrp_leave_records = [r for r in st.session_state.hrp_leave_records if r.get("leave_id") != selected_leave_id]
                                    _hrp_save_leave_records()
                                    log_action("HR_LEAVE_DELETED", selected_leave_id, old_data=old_leave)
                                    st.session_state.pop(delete_confirm_key, None)
                                    st.success(f"Leave record {selected_leave_id} deleted.")
                                    st.rerun()
                            with cdel2:
                                if st.button("Cancel", key=f"hrp_confirm_leave_delete_no_{selected_history_employee_id}_{selected_leave_id}"):
                                    st.session_state.pop(delete_confirm_key, None)
                                    st.rerun()

                    st.divider()
                    summary = _hrp_get_leave_summary(selected_history_employee_id)
                    st.subheader("Leave Summary")
                    col1, col2, col3, col4, col5 = st.columns(5)
                    with col1: st.metric("Holiday", f"{summary['holiday']:.1f}")
                    with col2: st.metric("Sick", f"{summary['sick']:.1f}")
                    with col3: st.metric("Unpaid", f"{summary['unpaid']:.1f}")
                    with col4: st.metric("Maternity", f"{summary['maternity']:.1f}")
                    with col5: st.metric("Paternity", f"{summary['paternity']:.1f}")
                    csv = df_history.to_csv(index=False)
                    st.download_button(
                        "📥 Export Leave History CSV", csv,
                        file_name=f"{selected_history_employee_id}_leave_history.csv",
                        mime="text/csv", key=f"hrp_export_csv_{selected_history_employee_id}"
                    )
                else:
                    st.info(f"No leave history is currently recorded for {history_employee['name'] if history_employee else selected_history_employee_id}.")
        elif emp:
            employee_leave = [r for r in _hrp_get_employee_leave(emp["emp_id"]) if r.get("status") == "Approved"]
            if employee_leave:
                history_data = [{"Leave ID": r["leave_id"], "Date From": r["date_from"], "Date To": r["date_to"], "Type": r["type"], "Days": r["days"], "Status": r["status"], "Notes": r["notes"]} for r in employee_leave]
                st.dataframe(pd.DataFrame(history_data), width="stretch", hide_index=True)
            else:
                st.info("No leave history is currently recorded.")
        else:
            st.info("No employee record is available.")

    if is_hr and hr_reports_tab is not None:
        with hr_reports_tab:
            render_hr_download_reports()

    if not is_hr:
        pass


def render_department_manager_leave_request(current_user_info=None):
    """Department-manager leave request workflow."""
    user = current_user_info or {}
    manager_name = str(user.get("full_name", "Department Manager")).strip()
    manager_department = str(user.get("dept", "")).strip()

    st.subheader("📝 Leave Request")
    st.caption("Submit leave on behalf of an employee in your department. HR must approve the request before it becomes official leave.")

    if not manager_department:
        st.error("Your account is not assigned to a department. Please contact the Super Admin.")
        return

    try:
        _hrp_init_storage()
        refreshed_employees = _hrp_load_employees()
        st.session_state.hrp_employees = refreshed_employees
        st.session_state.hrp_leave_records = _hrp_load_leave_records()
    except Exception as e:
        st.error(f"Unable to load HR employee or leave records: {e}")
        return

    manager_department_key = re.sub(r"\s+", " ", manager_department).strip().casefold()
    employees = [
        e for e in st.session_state.get("hrp_employees", [])
        if e.get("status", "Active").strip().casefold() == "active"
        and re.sub(r"\s+", " ", str(e.get("department", "")).strip()).casefold() == manager_department_key
    ]

    if not employees:
        st.info(f"No active employees have been added to the **{manager_department}** department by HR yet.")
        return

    employee_ids = [e["emp_id"] for e in employees]
    key_suffix = re.sub(r"[^A-Za-z0-9_]+", "_", manager_department) or "dept"
    selector_key = f"dept_leave_employee_{key_suffix}"
    previous_id = st.session_state.get(selector_key, employee_ids[0])
    if previous_id not in employee_ids:
        st.session_state[selector_key] = employee_ids[0]

    selected_id = st.selectbox(
        "👤 Employee — Request Leave For",
        employee_ids,
        format_func=lambda eid: _hrp_employee_label(eid),
        key=selector_key,
    )
    selected_employee = next((e for e in employees if e.get("emp_id") == selected_id), None)
    if not selected_employee:
        st.error("Selected employee could not be found.")
        return

    st.markdown("### 👤 Employee Details")
    d1, d2, d3, d4 = st.columns(4)
    with d1:
        st.write(f"**Name**\n\n{selected_employee.get('name', '-')}")
    with d2:
        st.write(f"**Employee ID**\n\n{selected_employee.get('emp_id', '-')}")
    with d3:
        st.write(f"**Position**\n\n{selected_employee.get('job_title', '-')}")
    with d4:
        st.write(f"**Department**\n\n{selected_employee.get('department', '-')}")
    st.info(
        f"Requesting leave for **{selected_employee['name']}** ({selected_employee['emp_id']}) "
        f"· Start date: {selected_employee.get('start_date', '-')} · Status: {selected_employee.get('status', '-') }"
    )

    form_version = st.session_state.get(f"dept_leave_form_version_{key_suffix}", 0)
    with st.form(f"dept_leave_request_form_{key_suffix}_{form_version}", clear_on_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            leave_type = st.selectbox("Leave Type", HR_PORTAL_LEAVE_TYPES, key=f"dept_leave_type_{key_suffix}_{form_version}")
            leave_start = st.date_input("Start Date", value=date.today(), key=f"dept_leave_start_{key_suffix}_{form_version}")
        with c2:
            leave_end = st.date_input("End Date", value=date.today(), key=f"dept_leave_end_{key_suffix}_{form_version}")
            half_day = st.checkbox("Half Day", disabled=leave_type != "Full Day Holiday", key=f"dept_leave_half_{key_suffix}_{form_version}")
        request_reference = st.text_input("Request Reference", placeholder="Optional email, call, or internal reference", key=f"dept_leave_ref_{key_suffix}_{form_version}")
        notes = st.text_area("Notes / Reason", placeholder="Reason or details supplied by the employee", key=f"dept_leave_notes_{key_suffix}_{form_version}")
        submitted = st.form_submit_button("📤 Submit Leave Request to HR", type="primary", width="stretch")

    if submitted:
        if leave_end < leave_start:
            st.error("End date cannot be before the start date.")
            return
        blocked_dates = _hrp_blocked_leave_dates(leave_start, leave_end)
        if blocked_dates:
            details = ", ".join(f"{d.strftime('%d/%m/%Y')} ({code} — {reason})" for d, code, reason in blocked_dates)
            st.error(f"Leave request cannot include company non-working dates: {details}")
            return
        days = _hrp_calculate_leave_days(leave_type, leave_start, leave_end, half_day)
        if days <= 0:
            st.error("The selected dates do not contain any working days.")
            return
        duplicate = next((
            r for r in st.session_state.hrp_leave_records
            if r.get("employee_id") == selected_employee["emp_id"]
            and r.get("date_from") == leave_start
            and r.get("date_to") == leave_end
            and r.get("type") == leave_type
            and r.get("status") == "Pending HR Approval"
        ), None)
        if duplicate:
            st.warning(f"A leave request for these dates is already waiting for HR approval ({duplicate['leave_id']}).")
            return
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        record = {
            "leave_id": _hrp_create_leave_id(),
            "employee_id": selected_employee["emp_id"],
            "date_from": leave_start,
            "date_to": leave_end,
            "type": leave_type,
            "days": days,
            "status": "Pending HR Approval",
            "request_source": "Department Manager",
            "request_reference": request_reference.strip(),
            "notes": notes.strip(),
            "recorded_by": manager_name,
            "recorded_at": now,
            "entry_source": "Department Manager Request",
            "requested_by": manager_name,
            "requested_at": now,
            "approved_by": "",
            "approved_at": "",
            "rejection_reason": "",
        }
        st.session_state.hrp_leave_records.append(record)
        _hrp_save_leave_records()
        log_action("HR_PORTAL_LEAVE_REQUESTED", record["leave_id"], new_data=record)
        st.session_state[f"dept_leave_form_version_{key_suffix}"] = form_version + 1
        st.success(f"Leave request {record['leave_id']} submitted to HR for approval.")
        st.rerun()

    pending = [
        r for r in st.session_state.hrp_leave_records
        if r.get("employee_id") in employee_ids
        and r.get("requested_by") == manager_name
        and r.get("status") in ("Pending HR Approval", "Rejected")
    ]
    if pending:
        st.divider()
        st.subheader("📋 My Leave Requests")
        rows = []
        for r in reversed(pending):
            employee = _hrp_get_employee(r["employee_id"])
            rows.append({
                "Leave ID": r["leave_id"],
                "Employee": employee["name"] if employee else r["employee_id"],
                "From": r["date_from"],
                "To": r["date_to"],
                "Type": r["type"],
                "Days": r["days"],
                "Status": r["status"],
                "HR Rejection Reason": r.get("rejection_reason", ""),
            })
        st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)


def render_hr_leave_approvals():
    """HR approval inbox for department-manager leave requests."""
    try:
        _hrp_init_storage()
        st.session_state.hrp_employees = _hrp_load_employees()
        st.session_state.hrp_leave_records = _hrp_load_leave_records()
    except Exception as e:
        st.error(f"Unable to refresh leave approval requests: {e}")
        return

    st.subheader("✅ Leave Approvals")
    st.caption("Department-manager leave requests are shown here. HR direct-entry leave is already Approved and does not require approval.")

    pending = [r for r in st.session_state.hrp_leave_records if r.get("status") == "Pending HR Approval"]
    rejected = [r for r in st.session_state.hrp_leave_records if r.get("status") == "Rejected"]

    pending_tab, rejected_tab = st.tabs([
        f"⏳ Pending Approval ({len(pending)})",
        f"❌ Rejected ({len(rejected)})",
    ])

    with pending_tab:
        if not pending:
            st.success("There are no leave requests waiting for HR approval.")
        else:
            for record in reversed(pending):
                employee = _hrp_get_employee(record.get("employee_id"))
                employee_name = employee.get("name", record.get("employee_id")) if employee else record.get("employee_id")
                department = employee.get("department", "") if employee else ""
                with st.expander(f"🟡 {record['leave_id']} — {employee_name} — {record['type']} — {record['days']:.1f} day(s)", expanded=True):
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.write(f"**Employee:** {employee_name}")
                        st.write(f"**Employee ID:** {record.get('employee_id')}")
                        st.write(f"**Department:** {department}")
                    with c2:
                        st.write(f"**Dates:** {record['date_from']} → {record['date_to']}")
                        st.write(f"**Leave Type:** {record['type']}")
                        st.write(f"**Days:** {record['days']:.1f}")
                    with c3:
                        st.write(f"**Requested By:** {record.get('requested_by', '')}")
                        st.write(f"**Requested At:** {record.get('requested_at', '')}")
                        st.write(f"**Reference:** {record.get('request_reference', '') or '-'}")
                    if record.get("notes"):
                        st.info(f"**Notes:** {record['notes']}")
                    rejection_reason = st.text_area("Rejection reason (required only if rejecting)", key=f"hr_reject_reason_{record['leave_id']}")
                    a1, a2 = st.columns(2)
                    with a1:
                        if st.button("✅ Approve Leave", type="primary", key=f"approve_leave_{record['leave_id']}"):
                            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            old = dict(record)
                            approver = str((st.session_state.get("user_info") or {}).get("full_name", "HR Manager"))
                            record.update({
                                "status": "Approved",
                                "approved_by": approver,
                                "approved_at": now,
                                "recorded_by": approver,
                                "recorded_at": now,
                                "rejection_reason": "",
                            })
                            _hrp_save_leave_records()
                            log_action("HR_PORTAL_LEAVE_APPROVED", record["leave_id"], old_data=old, new_data=dict(record), decision_by=approver, decision_date=now)
                            st.success(f"{record['leave_id']} approved. It is now part of the employee's official leave history.")
                            st.rerun()
                    with a2:
                        if st.button("❌ Reject Leave", key=f"reject_leave_{record['leave_id']}"):
                            if not rejection_reason.strip():
                                st.error("Please enter a rejection reason before rejecting the request.")
                            else:
                                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                old = dict(record)
                                approver = str((st.session_state.get("user_info") or {}).get("full_name", "HR Manager"))
                                record.update({
                                    "status": "Rejected",
                                    "rejection_reason": rejection_reason.strip(),
                                    "approved_by": approver,
                                    "approved_at": now,
                                })
                                _hrp_save_leave_records()
                                log_action("HR_PORTAL_LEAVE_REJECTED", record["leave_id"], old_data=old, new_data=dict(record), decision_by=approver, decision_date=now)
                                st.success(f"{record['leave_id']} rejected and moved to Rejected.")
                                st.rerun()

    with rejected_tab:
        if not rejected:
            st.info("There are no rejected department-manager leave requests.")
        else:
            for record in reversed(rejected):
                employee = _hrp_get_employee(record.get("employee_id"))
                employee_name = employee.get("name", record.get("employee_id")) if employee else record.get("employee_id")
                department = employee.get("department", "") if employee else ""
                with st.expander(f"🔴 {record['leave_id']} — {employee_name} — {record['type']} — {record['days']:.1f} day(s)"):
                    c1, c2 = st.columns(2)
                    with c1:
                        st.write(f"**Employee:** {employee_name}")
                        st.write(f"**Employee ID:** {record.get('employee_id')}")
                        st.write(f"**Department:** {department}")
                        st.write(f"**Requested By:** {record.get('requested_by', '')}")
                    with c2:
                        st.write(f"**Dates:** {record['date_from']} → {record['date_to']}")
                        st.write(f"**Leave Type:** {record['type']}")
                        st.write(f"**Days:** {record['days']:.1f}")
                        st.write(f"**Rejected By:** {record.get('approved_by', '')}")
                        st.write(f"**Rejected At:** {record.get('approved_at', '')}")
                    st.error(f"**Rejection reason:** {record.get('rejection_reason', '') or 'No reason recorded.'}")


def render_hr_department(current_user_info=None, is_super_admin=False, is_director=False, director_name="", has_hr_access=False):
    """HR department area. Department managers get requests only; HR Manager keeps the full direct-entry portal."""
    sub_portal_tab, sub_settlement_tab = st.tabs([
        "🧑‍💼 HR Portal (Employee / Holiday / Leave)",
        "💷 HR Leave Settlement"
    ])

    with sub_portal_tab:
        render_hr_portal(current_user_info)

    with sub_settlement_tab:
        if is_super_admin:
            render_hr_leave_super_admin()
        elif is_director:
            render_hr_leave_director_portal(director_name)
        elif has_hr_access:
            settlement_submit_tab, settlement_history_tab = st.tabs([
                "➕ New Settlement",
                "📋 My Submitted Settlements"
            ])
            with settlement_submit_tab:
                render_hr_leave_form(director_name)
            with settlement_history_tab:
                render_hr_leave_my_submissions(director_name)
        else:
            st.error("You do not have access to the HR Department module.")


def initialise_work_orders():
    safe_init_excel(WORK_ORDERS_PATH, WORK_ORDER_COLUMNS)

def load_work_orders(force=False):
    if not force and "_work_orders_cache" in st.session_state:
        return list(st.session_state["_work_orders_cache"])
    initialise_work_orders()
    try:
        df = _read_excel_records(WORK_ORDERS_PATH)
        records = []
        for r in df.to_dict(orient="records"):
            try: amount = float(r.get("Amount (£)", 0) or 0)
            except: amount = 0.0
            try: hours = float(r.get("Hours", 0) or 0)
            except: hours = 0.0
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

def clear_all_work_orders():
    try:
        save_all_work_orders([])
        return True
    except Exception: return False

def save_all_work_orders(records, sync=True):
    rows = []
    for r in records:
        rows.append({
            "Work Order ID": str(r.get("id", "")), "Manual Work Order No.": str(r.get("manual_work_order_no", "")).strip() or str(r.get("id", "")),
            "Employee Name": str(r.get("emp_name", "")), "Department": str(r.get("dept", "")), "Work Date": str(r.get("work_date", "")),
            "Hours": float(r.get("hours", 0)), "Amount (£)": float(r.get("amount", 0)), "Manager": str(r.get("manager", "")),
            "Description": str(r.get("desc", "")), "Attachment Name": str(r.get("attachment_name", "None")), "Status": str(r.get("status", "pending_manager")),
            "Site Address": str(r.get("site_address", "")), "Customer Job No.": str(r.get("customer_job_no", "")),
            "Manager Comments": str(r.get("manager_comments", "")), "Manager Decision Date": str(r.get("manager_decision_date", "")), "Manager Decision By": str(r.get("manager_decision_by", "")),
            "Director Comments": str(r.get("director_comments", "")), "Director Decision Date": str(r.get("director_decision_date", "")), "Director Decision By": str(r.get("director_decision_by", "")),
            "Submitted By": str(r.get("submitted_by", "")), "Submitted Date": str(r.get("submitted_date", "")),
            "Payroll Status": str(r.get("payroll_status", "Pending")), "Payroll Date": str(r.get("payroll_date", "")),
            "Payroll By": str(r.get("payroll_by", "")), "PDF File Path": str(r.get("pdf_path", "")),
        })
    pd.DataFrame(rows, columns=WORK_ORDER_COLUMNS).to_excel(WORK_ORDERS_PATH, index=False, engine="openpyxl")
    _set_data_cache("_work_orders_cache", list(records))
    if sync: sync_saved_file_to_drive(WORK_ORDERS_PATH)

def get_next_work_order_id(records):
    nums = []
    for r in records:
        raw = str(r.get("id", ""))
        try: nums.append(int(raw.replace("WO-", "")))
        except: pass
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
    candidates = [("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"), ("/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf", "/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf")]
    for regular, bold in candidates:
        if os.path.exists(regular) and os.path.exists(bold): return regular, bold
    return None, None

def _pdf_text(value):
    if value is None: return ""
    return str(value).replace("\x00", "")

def work_order_pdf(req, upload_to_drive=True):
    if not PDF_AVAILABLE: return None
    try:
        pdf = FPDF()
        pdf.add_page()
        regular_font, bold_font = _pdf_font_paths()
        if regular_font and bold_font:
            pdf.add_font("DejaVu", "", regular_font)
            pdf.add_font("DejaVu", "B", bold_font)
            font_family = "DejaVu"
        else: font_family = "Helvetica"
        def safe(value):
            text = _pdf_text(value)
            if font_family == "Helvetica": return text.encode("latin-1", "replace").decode("latin-1")
            return text
        if os.path.exists(LOGO_PATH):
            try:
                pdf.image(LOGO_PATH, x=75, y=10, w=60); pdf.ln(25)
            except Exception: pdf.ln(3)
        else: pdf.ln(3)
        label_width = 45
        content_width = pdf.w - pdf.l_margin - pdf.r_margin
        value_width = content_width - label_width
        pdf.set_font(font_family, "B", 16)
        pdf.cell(0, 10, safe("WORK ORDER - PAYMENT AUTHORISATION"), ln=True, align="C")
        pdf.ln(5)
        pdf.set_font(font_family, "", 10)
        rows = [("Work Order No.", get_work_order_number(req)), ("Employee", req.get("emp_name", "")), ("Department", req.get("dept", "")), ("Site Address", req.get("site_address", "")), ("Customer Job No.", req.get("customer_job_no", "")), ("Work Date", req.get("work_date", "")), ("Amount", f"GBP {float(req.get('amount', 0)):.2f}"), ("Manager", req.get("manager", "")), ("Submitted By", req.get("submitted_by", "")), ("Submitted Date", req.get("submitted_date", ""))]
        for label, value in rows:
            pdf.set_font(font_family, "B", 10)
            pdf.cell(label_width, 6, safe(label + ":"), 0, 0)
            pdf.set_font(font_family, "", 10)
            pdf.cell(value_width, 6, safe(value), ln=True)
        pdf.ln(2)
        pdf.set_font(font_family, "B", 10)
        pdf.cell(0, 6, safe("Work Performed / Description:"), ln=True)
        pdf.set_font(font_family, "", 10)
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(0, 6, safe(req.get("desc", "")))
        pdf.ln(3)
        pdf.set_font(font_family, "B", 10)
        pdf.cell(0, 6, safe("Director Final Approval"), ln=True)
        pdf.set_font(font_family, "", 10)
        pdf.set_x(pdf.l_margin)
        approved_by = req.get('director_decision_by', '')
        pdf.cell(35, 8, safe("Approved By:"), 0, 0)
        pdf.cell(80, 8, safe(approved_by), 0, 0)
        status = str(req.get('status', '')).strip().lower()
        if status in ["approved_payment", "approved"] and os.path.exists(APPROVED_STAMP_PATH):
            try: pdf.image(APPROVED_STAMP_PATH, x=pdf.get_x() + 5, y=pdf.get_y(), w=35)
            except Exception: pass
        pdf.ln(8)
        pdf.cell(35, 6, safe("Approval Date:"), 0, 0)
        pdf.cell(0, 6, safe(req.get('director_decision_date', '')), ln=True)
        pdf.cell(35, 6, safe("Comments:"), 0, 0)
        pdf.multi_cell(0, 6, safe(req.get('director_comments', '')))
        os.makedirs(WORK_ORDER_PDF_DIR, exist_ok=True)
        safe_emp = "_".join(str(req.get("emp_name", "Employee")).split()) or "Employee"
        safe_wo = "_".join(str(get_work_order_number(req)).split()) or "WO"
        safe_date = str(req.get("work_date", "")).replace("-", "")
        safe_amount = f"GBP{float(req.get('amount', 0)):.2f}"
        filename = f"Work_Order_{safe_wo}_{safe_emp}_{safe_amount}_{safe_date}.pdf"
        path = os.path.join(WORK_ORDER_PDF_DIR, filename)
        pdf.output(path)
        if upload_to_drive: _upload_to_drive_bg(path, os.path.basename(path))
        return path
    except Exception as e:
        st.error(f"Work Order PDF Error: {e}")
        return None

def display_work_order_pdf(req):
    path = req.get("pdf_path", "")
    if not path or not os.path.exists(path):
        path = work_order_pdf(req, upload_to_drive=False)
        if path:
            records = load_work_orders()
            for r in records:
                if str(r.get("id")) == str(req.get("id")):
                    r["pdf_path"] = path
            save_all_work_orders(records, sync=False)
    if path and os.path.exists(path):
        with open(path, "rb") as f:
            st.download_button("📄 Download Work Order PDF", f.read(), file_name=os.path.basename(path), key=f"wo_pdf_{req.get('id')}")

def work_order_total_pdf(records, employee_filter, from_date, to_date, prepared_by=""):
    if not PDF_AVAILABLE: return None
    try:
        selected = list(records)
        pdf = FPDF()
        pdf.add_page()
        regular_font, bold_font = _pdf_font_paths()
        if regular_font and bold_font:
            pdf.add_font("DejaVu", "", regular_font)
            pdf.add_font("DejaVu", "B", bold_font)
            family = "DejaVu"
        else: family = "Helvetica"
        def safe(v):
            text = _pdf_text(v)
            if family == "Helvetica": return text.encode("latin-1", "replace").decode("latin-1")
            return text
        if os.path.exists(LOGO_PATH):
            try:
                pdf.image(LOGO_PATH, x=75, y=10, w=60); pdf.ln(25)
            except Exception: pdf.ln(3)
        else: pdf.ln(3)
        pdf.set_font(family, "B", 15)
        pdf.cell(0, 9, safe("APPROVED WORK ORDER TOTAL"), ln=True, align="C")
        pdf.ln(3)
        pdf.set_font(family, "", 10)
        scope = employee_filter if employee_filter and employee_filter != "All Employees" else "All Employees"
        pdf.cell(0, 6, safe(f"Employee: {scope}"), ln=True)
        pdf.cell(0, 6, safe(f"Work date range: {from_date} to {to_date}"), ln=True)
        if prepared_by: pdf.cell(0, 6, safe(f"Prepared by: {prepared_by}"), ln=True)
        pdf.ln(3)
        pdf.set_font(family, "B", 9)
        pdf.cell(45, 7, safe("Work Order No."), border=1)
        pdf.cell(55, 7, safe("Employee"), border=1)
        pdf.cell(30, 7, safe("Work Date"), border=1)
        pdf.cell(28, 7, safe("Amount"), border=1, align="R")
        pdf.cell(0, 7, safe("Approved By"), border=1, align="L", ln=True)
        pdf.set_font(family, "", 9)
        total = 0.0
        for r in selected:
            amount = float(r.get("amount", 0) or 0)
            total += amount
            pdf.cell(45, 7, safe(get_work_order_number(r)), border=1)
            pdf.cell(55, 7, safe(r.get("emp_name", "")), border=1)
            pdf.cell(30, 7, safe(r.get("work_date", "")), border=1)
            pdf.cell(28, 7, safe(f"{amount:.2f}"), border=1, align="R")
            pdf.cell(0, 7, safe(r.get("director_decision_by", "")), border=1, ln=True)
        pdf.set_font(family, "B", 10)
        pdf.cell(130, 8, safe("TOTAL"), border=1, align="R")
        pdf.cell(28, 8, safe(f"{total:.2f}"), border=1, align="R")
        pdf.cell(0, 8, safe(""), border=1, ln=True)
        os.makedirs(WORK_ORDER_PDF_DIR, exist_ok=True)
        safe_emp = "all_employees" if scope == "All Employees" else "_".join(scope.split())
        filename = (f"Approved_Work_Order_Total_{safe_emp}_" f"GBP{total:.2f}_{from_date}_to_{to_date}.pdf")
        path = os.path.join(WORK_ORDER_PDF_DIR, filename)
        pdf.output(path)
        _upload_to_drive_bg(path, os.path.basename(path))
        return path
    except Exception as e:
        st.error(f"Work Order Total PDF Error: {e}")
        return None

def render_work_order_total(records, scope_department=None, key_prefix="wo_total", prepared_by=""):
    st.markdown("### 💷 Approved Work Order Total")
    st.caption("Select an employee and work-date range. Only Director-approved work orders are included.")
    scoped = [r for r in records if r.get("status") in ("approved_payment", "approved")]
    if scope_department: scoped = [r for r in scoped if str(r.get("dept", "")) == str(scope_department)]
    employees = sorted({str(r.get("emp_name", "")).strip() for r in scoped if str(r.get("emp_name", "")).strip()})
    c1, c2, c3 = st.columns(3)
    with c1: employee = st.selectbox("👤 Employee", ["All Employees"] + employees, key=f"{key_prefix}_employee")
    with c2: from_date = st.date_input("📅 From Date", value=date.today().replace(day=1), key=f"{key_prefix}_from")
    with c3: to_date = st.date_input("📅 To Date", value=date.today(), key=f"{key_prefix}_to")
    if from_date > to_date:
        st.error("From Date cannot be after To Date.")
        return
    selected = []
    for r in scoped:
        try: d = pd.to_datetime(str(r.get("work_date", "")), errors="coerce").date()
        except Exception: d = None
        if d is None: continue
        if not (from_date <= d <= to_date): continue
        if employee != "All Employees" and str(r.get("emp_name", "")).strip() != employee: continue
        selected.append(r)
    total = sum(float(r.get("amount", 0) or 0) for r in selected)
    st.metric("💷 Total Approved", f"£{total:,.2f}")
    st.metric("📋 Work Orders Included", len(selected))
    if selected:
        rows = [{"Work Order No.": get_work_order_number(r), "Employee": r.get("emp_name", ""), "Work Date": r.get("work_date", ""), "Amount (£)": float(r.get("amount", 0) or 0), "Approved By": r.get("director_decision_by", "")} for r in selected]
        st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
        if st.button("📄 Generate & Download Total PDF", key=f"{key_prefix}_pdf", type="primary"):
            with st.spinner("Generating PDF..."):
                path = work_order_total_pdf(selected, employee, from_date, to_date, prepared_by=prepared_by)
            if path and os.path.exists(path):
                with open(path, "rb") as f:
                    st.download_button("⬇️ Download Total PDF", f.read(), file_name=os.path.basename(path), key=f"{key_prefix}_download")
    else:
        st.info("No approved work orders match the selected employee and date range.")

def render_work_order_bulk_download(records, key_prefix="wo_bulk"):
    st.markdown("### 📦 Download Multiple Work Order PDFs (ZIP)")
    scoped = [r for r in records if r.get("status") in ("approved_payment", "approved")]
    if not scoped:
        st.info("No approved work orders available.")
        return
    employees = sorted({str(r.get("emp_name", "")).strip() for r in scoped if str(r.get("emp_name", "")).strip()})
    c1, c2, c3 = st.columns(3)
    with c1: employee = st.selectbox("👤 Employee", ["All Employees"] + employees, key=f"{key_prefix}_employee")
    with c2: from_date = st.date_input("📅 From Date", value=date.today().replace(day=1), key=f"{key_prefix}_from")
    with c3: to_date = st.date_input("📅 To Date", value=date.today(), key=f"{key_prefix}_to")
    if from_date > to_date:
        st.error("From Date cannot be after To Date.")
        return
    selected = []
    for r in scoped:
        try: d = pd.to_datetime(str(r.get("work_date", "")), errors="coerce").date()
        except Exception: d = None
        if d is None: continue
        if not (from_date <= d <= to_date): continue
        if employee != "All Employees" and str(r.get("emp_name", "")).strip() != employee: continue
        selected.append(r)
    st.metric("📋 Work Orders Matching", len(selected))
    if not selected:
        st.info("No approved work orders match the selected employee and date range.")
        return
    if st.button("📦 Generate ZIP of Work Order PDFs", type="primary", key=f"{key_prefix}_gen"):
        import zipfile
        zip_buffer = io.BytesIO()
        pdf_count = 0
        failed = []
        with st.spinner(f"Generating {len(selected)} PDFs..."):
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for r in selected:
                    path = r.get("pdf_path", "")
                    if not path or not os.path.exists(path): path = work_order_pdf(r, upload_to_drive=False)
                    if path and os.path.exists(path):
                        with open(path, "rb") as f: zipf.writestr(os.path.basename(path), f.read())
                        pdf_count += 1
                    else: failed.append(get_work_order_number(r))
        zip_buffer.seek(0)
        if pdf_count > 0:
            safe_emp = "all_employees" if employee == "All Employees" else "_".join(employee.split())
            filename = f"Work_Order_PDFs_{safe_emp}_{from_date}_to_{to_date}.zip"
            st.session_state[f"{key_prefix}_data"] = zip_buffer.getvalue()
            st.session_state[f"{key_prefix}_name"] = filename
            st.success(f"✅ Generated {pdf_count} PDF(s). Click below to download.")
            if failed: st.warning(f"⚠️ Could not include: {', '.join(failed)}")
        else: st.error("❌ No PDFs could be generated.")
    if st.session_state.get(f"{key_prefix}_data"):
        st.download_button("⬇️ Download ZIP", data=st.session_state[f"{key_prefix}_data"], file_name=st.session_state.get(f"{key_prefix}_name", "Work_Order_PDFs.zip"), mime="application/zip", type="primary", key=f"{key_prefix}_dl")

def render_work_orders_super_admin():
    st.subheader("🛠️ Work Orders — Super Admin (View Only)")
    st.info("✅ View all work orders across ALL departments. **Approval → Director only.**")
    orders = load_work_orders()
    search = st.text_input("🔎 Search work orders", placeholder="Search by Work Order No., employee, manager, department, status, amount or date...", key="wo_super_admin_search")
    if search.strip():
        q = search.lower().strip()
        orders = [r for r in orders if q in " ".join(str(v) for v in r.values()).lower()]
    pending_mgr = [r for r in orders if r.get("status") == "pending_manager"]
    pending_dir = [r for r in orders if r.get("status") == "pending_director"]
    approved = [r for r in orders if r.get("status") in ("approved_payment", "approved")]
    rejected = [r for r in orders if r.get("status") in ("rejected_director", "rejected", "returned_to_employee")]
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("🟡 Pending Manager", len(pending_mgr))
    with c2: st.metric("🟡 Pending Director", len(pending_dir))
    with c3: st.metric("🟢 Approved", len(approved))
    with c4: st.metric("🔴 Rejected / Returned", len(rejected))
    st.divider()
    t1, t2, t3, t4 = st.tabs([
        f"⏳ Pending ({len(pending_mgr) + len(pending_dir)})",
        f"✅ Approved ({len(approved)})",
        f"❌ Rejected / Returned ({len(rejected)})",
        "💷 Approved Work Order Total"
    ])
    def show_wo_details(r):
        st.write(f"🧾 **Work Order No.:** {get_work_order_number(r)}")
        st.write(f"👤 **Contractor / Employee Labour:** {r.get('emp_name','-')}")
        st.write(f"🏢 **Department:** {r.get('dept','-')}")
        st.write(f"📍 **Site Address:** {r.get('site_address','-')}")
        st.write(f"📘 **Customer Job No.:** {r.get('customer_job_no','-')}")
        st.write(f"📅 **Work Date:** {r.get('work_date','-')}")
        st.write(f"💷 **Amount:** £{float(r.get('amount',0) or 0):.2f}")
        st.write(f"👔 **Manager:** {r.get('manager','-')}")
        st.write(f"📝 **Submitted by:** {r.get('submitted_by','-')} on {r.get('submitted_date','-')}")
        if r.get('manager_decision_by'): st.write(f"👔 **Manager review:** {r.get('manager_decision_by')} on {r.get('manager_decision_date','')}")
        if r.get('manager_comments'): st.info(f"💬 **Manager Comments:** {r.get('manager_comments')}")
        if r.get('director_decision_by'): st.write(f"🎯 **Director:** {r.get('director_decision_by')} on {r.get('director_decision_date','')}")
        if r.get('director_comments'): st.warning(f"💬 **Director Comments:** {r.get('director_comments')}")
        st.info(f"📝 **Description:**\n{r.get('desc','')}")
        st.divider(); st.markdown("#### 📎 Attachments"); display_attachments(r)
    with t1:
        pending_all = pending_mgr + pending_dir
        if not pending_all: st.success("✅ No pending work orders.")
        for r in reversed(pending_all):
            st_raw = str(r.get("status", "")).strip().lower()
            label = "PENDING MANAGER" if st_raw == "pending_manager" else "PENDING DIRECTOR"
            submitter = str(r.get("submitted_by", "") or "").strip() or "Unknown"
            with st.expander(f"🟡 {get_work_order_number(r)} | {r.get('emp_name')} | £{float(r.get('amount',0) or 0):.2f} | {label} | Submitted by: {submitter}"):
                show_wo_details(r)
    with t2:
        if not approved: st.info("✅ No approved work orders.")
        for r in reversed(approved):
            submitter = str(r.get("submitted_by", "") or "").strip() or "Unknown"
            with st.expander(f"🟢 {get_work_order_number(r)} | {r.get('emp_name')} | £{float(r.get('amount',0) or 0):.2f} | Approved by {r.get('director_decision_by','')} | Submitted by: {submitter}"):
                show_wo_details(r); st.divider(); st.markdown("#### 📄 Work Order PDF"); display_work_order_pdf(r)
    with t3:
        if not rejected: st.info("❌ No rejected or returned work orders.")
        for r in reversed(rejected):
            st_raw = str(r.get("status", "")).strip().lower()
            label = "RETURNED TO EMPLOYEE" if st_raw == "returned_to_employee" else "REJECTED"
            submitter = str(r.get("submitted_by", "") or "").strip() or "Unknown"
            with st.expander(f"🔴 {get_work_order_number(r)} | {r.get('emp_name')} | £{float(r.get('amount',0) or 0):.2f} | {label} | Submitted by: {submitter}"):
                show_wo_details(r)
    with t4:
        render_work_order_total(orders, key_prefix="wo_super_admin_total", prepared_by="Super Admin")

def render_work_order_employee_portal(current_user, current_dept):
    st.subheader("🛠️ Work Orders")
    orders = load_work_orders()
    managers = _manager_options_for_department(current_dept)
    _success_msg = st.session_state.pop("emp_new_wo_success", None)
    if _success_msg: st.success(_success_msg)
    editing_id = st.session_state.get("editing_work_order_id_emp")
    user_info = st.session_state.get("user_info", {})
    can_access_wo_total = user_info.get("can_access_wo_total", False)
    if can_access_wo_total:
        main_tab, total_tab = st.tabs(["🛠️ Work Orders", "💷 Approved Work Order Total"])
    else:
        main_tab = st.tabs(["🛠️ Work Orders"])[0]
    with main_tab:
        if editing_id:
            rec = next((r for r in orders if str(r.get("id")) == str(editing_id)), None)
            editable_statuses = ("pending_manager", "returned_to_employee", "rejected_director", "pending")
            if rec and str(rec.get("status", "")).strip().lower() in editable_statuses:
                st.warning(f"✏️ **Editing Work Order {get_work_order_number(rec)}** — save changes to resubmit.")
                with st.form("work_order_emp_edit_form", clear_on_submit=False):
                    ec1, ec2 = st.columns(2)
                    with ec1:
                        e_wo_no = st.text_input("🧾 Work Order No.", value=str(rec.get("manual_work_order_no", "") or rec.get("id", "")))
                        e_emp = st.text_input("👤 Contractor / Employee Labour", value=str(rec.get("emp_name", "")))
                        e_site = st.text_area("📍 Site Address", value=str(rec.get("site_address", "")), height=90)
                        e_cjno = st.text_input("📘 Customer Job No.", value=str(rec.get("customer_job_no", "")))
                    with ec2:
                        try: _d = datetime.strptime(str(rec.get("work_date", ""))[:10], "%Y-%m-%d").date()
                        except Exception: _d = date.today()
                        e_date = st.date_input("📅 Date", value=_d)
                        e_amount = st.number_input("💷 Amount (£)", min_value=0.01, step=1.0, format="%.2f", value=max(float(rec.get("amount", 0.01) or 0.01), 0.01))
                        e_desc = st.text_area("📝 Description", value=str(rec.get("desc", "")), height=150)
                    if managers:
                        current_mgr = str(rec.get("manager", "") or "")
                        mgr_idx = managers.index(current_mgr) if current_mgr in managers else 0
                        e_manager = st.selectbox("👔 Send to Manager", managers, index=mgr_idx)
                    else:
                        e_manager = st.text_input("👔 Manager Name", value=str(rec.get("manager", "")))
                    e_files = st.file_uploader("📎 Add New Attachments (optional)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key=f"emp_edit_files_{editing_id}")
                    existing_att = str(rec.get("attachment_name", "None") or "None")
                    if existing_att and existing_att.lower() not in ("none", "nan", ""):
                        st.caption(f"📎 Existing attachments: {existing_att}")
                    csave, ccancel = st.columns(2)
                    with csave: save_edit = st.form_submit_button("💾 Save Changes & Resubmit", type="primary", width="stretch")
                    with ccancel: cancel_edit = st.form_submit_button("❌ Cancel", width="stretch")
                    if save_edit:
                        if not e_wo_no.strip() or not e_emp.strip() or not e_manager.strip() or not e_desc.strip():
                            st.error("❌ Work Order No., Employee, Manager and Description are required.")
                        else:
                            existing_list = [n.strip() for n in existing_att.split(",") if n.strip() and n.strip().lower() not in ("none", "nan")]
                            new_attachments = list(existing_list)
                            if e_files:
                                for i, f in enumerate(e_files, start=len(new_attachments) + 1):
                                    safe_name = os.path.basename(f.name).replace("/", "_").replace("\\", "_")
                                    fn = f"{editing_id}_EDIT_F{i}_{safe_name}"
                                    fp = os.path.join(UPLOAD_DIR, fn)
                                    with open(fp, "wb") as out_file: out_file.write(f.getbuffer())
                                    _upload_to_drive_bg(fp, fn)
                                    new_attachments.append(fn)
                            old_data = {"manual_work_order_no": rec.get("manual_work_order_no"), "emp_name": rec.get("emp_name"), "site_address": rec.get("site_address"), "customer_job_no": rec.get("customer_job_no"), "work_date": rec.get("work_date"), "amount": rec.get("amount"), "desc": rec.get("desc"), "manager": rec.get("manager"), "status": rec.get("status")}
                            for x in orders:
                                if str(x.get("id")) == str(editing_id):
                                    x["manual_work_order_no"] = e_wo_no.strip(); x["emp_name"] = e_emp.strip(); x["site_address"] = e_site.strip(); x["customer_job_no"] = e_cjno.strip(); x["work_date"] = str(e_date); x["amount"] = float(e_amount); x["desc"] = e_desc.strip(); x["manager"] = e_manager.strip(); x["attachment_name"] = ", ".join(new_attachments) or "None"; x["status"] = "pending_manager"; x["manager_comments"] = ""; x["manager_decision_date"] = ""; x["manager_decision_by"] = ""; x["director_comments"] = ""; x["director_decision_date"] = ""; x["director_decision_by"] = ""; x["pdf_path"] = ""; x["submitted_by"] = current_user; x["submitted_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    break
                            save_all_work_orders(orders)
                            new_data = {"manual_work_order_no": e_wo_no.strip(), "emp_name": e_emp.strip(), "site_address": e_site.strip(), "customer_job_no": e_cjno.strip(), "work_date": str(e_date), "amount": float(e_amount), "desc": e_desc.strip(), "manager": e_manager.strip(), "status": "pending_manager"}
                            log_action("WORK_ORDER_EDITED", editing_id, old_data=old_data, new_data=new_data)
                            st.session_state.editing_work_order_id_emp = None
                            st.success("✅ Work Order updated and re-sent to Manager for review.")
                            st.rerun()
                    if cancel_edit:
                        st.session_state.editing_work_order_id_emp = None; st.rerun()
                st.divider()
            else: st.session_state.editing_work_order_id_emp = None
        st.markdown("### 📤 Submit New Work Order")
        st.caption("Complete the work-order details below. No hours/time entry is required.")
        wid = get_next_work_order_id(orders)
        form_version = st.session_state.get("emp_new_wo_form_version", 0)
        K_WO = f"emp_new_wo_no_v{form_version}"
        K_EMP = f"emp_new_contractor_v{form_version}"
        K_SITE = f"emp_new_site_v{form_version}"
        K_CJNO = f"emp_new_cjno_v{form_version}"
        K_DATE = f"emp_new_date_v{form_version}"
        K_AMT = f"emp_new_amount_v{form_version}"
        K_DESC = f"emp_new_desc_v{form_version}"
        K_MGR = f"emp_new_manager_v{form_version}"
        K_FILES = f"emp_new_files_v{form_version}"
        with st.form(f"employee_new_work_order_form_v{form_version}", clear_on_submit=False):
            c1, c2 = st.columns(2)
            with c1:
                work_order_no = st.text_input("🧾 Work Order No.", key=K_WO, placeholder="Enter the Work Order No. from the work-order sheet")
                contractor_employee = st.text_input("👤 Contractor / Employee Labour", key=K_EMP, value=current_user, placeholder="Enter contractor or employee name")
                site_address = st.text_area("📍 Site Address", key=K_SITE, placeholder="Enter the full site address", height=90)
                customer_job_no = st.text_input("📘 Customer Job No.", key=K_CJNO, placeholder="Enter Customer Job No.")
            with c2:
                work_date = st.date_input("📅 Date", key=K_DATE, value=date.today())
                amount = st.number_input("💷 Amount (£)", min_value=0.01, step=1.0, format="%.2f", key=K_AMT, value=0.01)
                description = st.text_area("📝 Description", key=K_DESC, placeholder="Describe the work completed...", height=150)
            if managers: manager = st.selectbox("👔 Send to Manager", managers, key=K_MGR)
            else: manager = st.text_input("👔 Manager Name", key=K_MGR, placeholder="Enter manager name")
            files = st.file_uploader("📎 Supporting Work Order Document (optional)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key=K_FILES)
            submitted = st.form_submit_button("📤 Submit Work Order to Manager", type="primary", width="stretch")
            if submitted:
                errors = []
                if not work_order_no.strip(): errors.append("Work Order No.")
                if not contractor_employee.strip(): errors.append("Contractor / Employee Labour")
                if not site_address.strip(): errors.append("Site Address")
                if not customer_job_no.strip(): errors.append("Customer Job No.")
                if not description.strip(): errors.append("Description")
                if not manager.strip(): errors.append("Manager")
                duplicate = any(str(r.get("manual_work_order_no", "")).strip().lower() == work_order_no.strip().lower() for r in orders if str(r.get("manual_work_order_no", "")).strip())
                if duplicate:
                    st.error("❌ **Work Order No. already Exist.** Please change the Work Order No. and submit again — your form data has been kept.")
                elif errors:
                    st.error("Please correct: " + ", ".join(errors) + ".")
                else:
                    attachments = []
                    for i, f in enumerate(files or [], 1):
                        safe_name = os.path.basename(f.name).replace("/", "_").replace("\\", "_")
                        fn = f"{wid}_F{i}_{safe_name}"
                        fp = os.path.join(UPLOAD_DIR, fn)
                        with open(fp, "wb") as out_file: out_file.write(f.getbuffer())
                        _upload_to_drive_bg(fp, fn)
                        attachments.append(fn)
                    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    rec = {"id": wid, "manual_work_order_no": work_order_no.strip(), "emp_name": contractor_employee.strip(), "dept": current_dept, "work_date": str(work_date), "hours": 0.0, "customer_job_no": customer_job_no.strip(), "site_address": site_address.strip(), "amount": float(amount), "manager": manager.strip(), "desc": description.strip(), "attachment_name": ", ".join(attachments) or "None", "status": "pending_manager", "manager_comments": "", "manager_decision_date": "", "manager_decision_by": "", "director_comments": "", "director_decision_date": "", "director_decision_by": "", "submitted_by": current_user, "submitted_date": now, "payroll_status": "Pending", "payroll_date": "", "payroll_by": "", "pdf_path": ""}
                    orders.append(rec)
                    save_all_work_orders(orders)
                    log_action("WORK_ORDER_CREATED", wid, decision_by=current_user)
                    for k in (K_WO, K_EMP, K_SITE, K_CJNO, K_DATE, K_AMT, K_DESC, K_MGR, K_FILES):
                        st.session_state.pop(k, None)
                    st.session_state["emp_new_wo_form_version"] = form_version + 1
                    st.session_state["emp_new_wo_success"] = (f"✅ Work Order No. {work_order_no.strip()} submitted to {manager} for review.")
                    st.rerun()
        st.divider()
        tab_mine, tab_dept = st.tabs([
            "📋 My Work Orders",
            f"🏢 Department Work Orders ({current_dept})"
        ])
        with tab_mine:
            q = st.text_input("🔎 Search my work orders", placeholder="Search by ID, employee, manager, status, amount, date or description...", key="wo_employee_search")
            mine = [r for r in orders if r.get("submitted_by") == current_user]
            if q.strip():
                ql = q.lower().strip()
                mine = [r for r in mine if ql in " ".join(str(v) for v in r.values()).lower()]
            if not mine: st.info("📋 You have not submitted any work orders yet.")
            else:
                for r in reversed(mine):
                    status_raw = str(r.get("status", "")).strip().lower()
                    display_status = {"pending_manager": "PENDING MANAGER", "pending_director": "PENDING DIRECTOR", "approved_payment": "APPROVED", "approved": "APPROVED", "rejected_director": "REJECTED BY DIRECTOR", "rejected": "REJECTED", "returned_to_employee": "RETURNED — NEEDS EDITS"}.get(status_raw, status_raw.replace("_", " ").upper())
                    icon = {"pending_manager": "🟡", "pending_director": "🟡", "approved_payment": "🟢", "approved": "🟢", "rejected_director": "🔴", "rejected": "🔴", "returned_to_employee": "🟠"}.get(status_raw, "⚪")
                    with st.expander(f"{icon} {get_work_order_number(r)} | {r.get('emp_name')} | £{r.get('amount', 0):.2f} | {display_status}"):
                        st.write(f"🧾 Work Order No.: **{get_work_order_number(r)}** | 👔 Manager: {r.get('manager')} | 🏢 {r.get('dept')} | 📅 {r.get('work_date')}")
                        st.write(f"💷 £{r.get('amount', 0):.2f}")
                        if r.get("site_address"): st.write(f"📍 Site Address: {r.get('site_address')}")
                        if r.get("customer_job_no"): st.write(f"📘 Customer Job No.: {r.get('customer_job_no')}")
                        st.info(f"📝 {r.get('desc')}")
                        display_attachments(r)
                        if r.get("manager_decision_by"): st.write(f"👔 Manager reviewed by: {r.get('manager_decision_by')} on {r.get('manager_decision_date')}")
                        if r.get("manager_comments"):
                            if status_raw == "returned_to_employee": st.error(f"❌ Manager Comments: {r.get('manager_comments')}")
                            else: st.info(f"💬 Manager Comments: {r.get('manager_comments')}")
                        if r.get("director_decision_by"): st.write(f"🎯 Approved By: {r.get('director_decision_by')} on {r.get('director_decision_date')}")
                        if r.get("director_comments"): st.warning(f"💬 Director Comments: {r.get('director_comments')}")
                        if status_raw in ("pending_manager", "returned_to_employee", "rejected_director", "rejected", "pending"):
                            st.divider()
                            if st.button(f"✏️ Edit Work Order {get_work_order_number(r)}", key=f"wo_emp_edit_{r.get('id')}", type="secondary"):
                                st.session_state.editing_work_order_id_emp = r.get("id"); st.rerun()
        with tab_dept:
            st.caption("Work orders already recorded for your department (including those submitted by managers). 👉 Check the Work Order No. here **before** submitting a new one to avoid duplicates.")
            dept_orders = [r for r in orders if str(r.get("dept", "")).strip() == str(current_dept).strip()]
            dq = st.text_input("🔎 Search department work orders", placeholder="Search by Work Order No., employee, manager, status, amount or date...", key="wo_emp_dept_search")
            if dq.strip():
                dql = dq.lower().strip()
                dept_orders = [r for r in dept_orders if dql in " ".join(str(v) for v in r.values()).lower()]
            if not dept_orders: st.info("📋 No work orders recorded for your department yet.")
            else:
                rows = []
                for r in sorted(dept_orders, key=lambda x: str(x.get("work_date", "")), reverse=True):
                    rows.append({"Work Order No.": get_work_order_number(r), "Employee": r.get("emp_name", ""), "Manager": r.get("manager", ""), "Work Date": r.get("work_date", ""), "Amount (£)": float(r.get("amount", 0) or 0), "Status": str(r.get("status", "")).replace("_", " ").upper(), "Submitted By": r.get("submitted_by", "")})
                df_view = pd.DataFrame(rows)
                st.dataframe(df_view, width="stretch", hide_index=True)
                st.caption(f"📊 {len(rows)} work order(s) in **{current_dept}**")
                st.divider(); st.markdown("#### 📄 Full Details")
                for r in reversed(dept_orders):
                    status_raw = str(r.get("status", "")).strip().lower()
                    display_status = {"pending_manager": "PENDING MANAGER", "pending_director": "PENDING DIRECTOR", "approved_payment": "APPROVED", "approved": "APPROVED", "rejected_director": "REJECTED", "rejected": "REJECTED", "returned_to_employee": "RETURNED"}.get(status_raw, status_raw.replace("_", " ").upper())
                    with st.expander(f"🧾 {get_work_order_number(r)} | {r.get('emp_name')} | £{float(r.get('amount', 0) or 0):.2f} | {display_status}"):
                        st.write(f"👔 **Manager:** {r.get('manager', '-')} | 📅 **Date:** {r.get('work_date', '-')}")
                        st.write(f"📝 **Submitted by:** {r.get('submitted_by', '-')}")
                        if r.get("site_address"): st.write(f"📍 **Site Address:** {r.get('site_address')}")
                        if r.get("customer_job_no"): st.write(f"📘 **Customer Job No.:** {r.get('customer_job_no')}")
                        st.info(f"📝 {r.get('desc', '')}")
    if can_access_wo_total:
        with total_tab:
            render_work_order_total(orders, scope_department=current_dept, key_prefix="wo_emp_total", prepared_by=current_user)

def render_work_order_manager_portal(manager_name, manager_dept, show_total=True):
    st.subheader("🛠️ Work Orders")
    orders = load_work_orders()
    editing_id = st.session_state.get("editing_work_order_id")
    if editing_id:
        rec = next((r for r in orders if str(r.get("id")) == str(editing_id)), None)
        editable_statuses = ("pending_manager", "pending_director", "rejected_director", "returned_to_employee", "pending")
        if rec and str(rec.get("status", "")).strip().lower() in editable_statuses:
            st.warning(f"✏️ **Editing Work Order {get_work_order_number(rec)}** — save changes to update.")
            with st.form("work_order_edit_form", clear_on_submit=False):
                ec1, ec2 = st.columns(2)
                with ec1:
                    e_wo_no = st.text_input("🧾 Work Order No.", value=str(rec.get("manual_work_order_no", "") or rec.get("id", "")))
                    e_emp = st.text_input("👤 Contractor / Employee Labour", value=str(rec.get("emp_name", "")))
                    e_site = st.text_area("📍 Site Address", value=str(rec.get("site_address", "")), height=90)
                    e_cjno = st.text_input("📘 Customer Job No.", value=str(rec.get("customer_job_no", "")))
                with ec2:
                    try: _d = datetime.strptime(str(rec.get("work_date", ""))[:10], "%Y-%m-%d").date()
                    except Exception: _d = date.today()
                    e_date = st.date_input("📅 Date", value=_d)
                    e_amount = st.number_input("💷 Amount (£)", min_value=0.01, step=1.0, format="%.2f", value=max(float(rec.get("amount", 0.01) or 0.01), 0.01))
                    e_desc = st.text_area("📝 Description", value=str(rec.get("desc", "")), height=150)
                csave, ccancel = st.columns(2)
                with csave: save_edit = st.form_submit_button("💾 Save Changes", type="primary", width="stretch")
                with ccancel: cancel_edit = st.form_submit_button("❌ Cancel", width="stretch")
                if save_edit:
                    old_data = {"manual_work_order_no": rec.get("manual_work_order_no"), "emp_name": rec.get("emp_name"), "site_address": rec.get("site_address"), "customer_job_no": rec.get("customer_job_no"), "work_date": rec.get("work_date"), "amount": rec.get("amount"), "desc": rec.get("desc")}
                    for x in orders:
                        if str(x.get("id")) == str(editing_id):
                            x["manual_work_order_no"] = e_wo_no.strip() or str(x.get("id")); x["emp_name"] = e_emp.strip(); x["site_address"] = e_site.strip(); x["customer_job_no"] = e_cjno.strip(); x["work_date"] = str(e_date); x["amount"] = float(e_amount); x["desc"] = e_desc.strip(); x["pdf_path"] = ""
                            break
                    save_all_work_orders(orders)
                    new_data = {"manual_work_order_no": e_wo_no.strip(), "emp_name": e_emp.strip(), "site_address": e_site.strip(), "customer_job_no": e_cjno.strip(), "work_date": str(e_date), "amount": float(e_amount), "desc": e_desc.strip()}
                    log_action("WORK_ORDER_EDITED", editing_id, old_data=old_data, new_data=new_data)
                    st.session_state.editing_work_order_id = None
                    st.success("✅ Work Order updated."); st.rerun()
                if cancel_edit:
                    st.session_state.editing_work_order_id = None; st.rerun()
            st.divider()
        else: st.session_state.editing_work_order_id = None
    manager_orders = [r for r in orders if str(r.get("manager", "")).strip() == str(manager_name).strip() or str(r.get("submitted_by", "")).strip() == str(manager_name).strip()]
    pending = [r for r in manager_orders if r.get("status") in ("pending_director", "pending_manager", "pending")]
    approved = [r for r in manager_orders if r.get("status") in ("approved_payment", "approved")]
    rejected = [r for r in manager_orders if r.get("status") in ("rejected_director", "rejected", "returned_to_employee")]
    user_info = st.session_state.get("user_info", {})
    can_access_wo_total = user_info.get("can_access_wo_total", False)
    if can_access_wo_total:
        main_tab, total_tab = st.tabs(["🛠️ Work Orders", "💷 Approved Work Order Total"])
    else:
        main_tab = st.tabs(["🛠️ Work Orders"])[0]
    with main_tab:
        st.markdown("### 📤 Submit New Work Order")
        st.caption("Complete the work-order details below. No hours/time entry is required.")
        form_version = st.session_state.get("wo_mgr_new_wo_form_version", 0)
        with st.form(f"wo_mgr_new_work_order_form_v{form_version}", clear_on_submit=False):
            c1, c2 = st.columns(2)
            with c1:
                work_order_no = st.text_input("🧾 Work Order No.", key=f"wo_mgr_wo_v{form_version}", placeholder="Enter the Work Order No. from the work-order sheet")
                contractor_employee = st.text_input("👤 Contractor / Employee Labour", key=f"wo_mgr_emp_v{form_version}", placeholder="Enter contractor or employee name")
                site_address = st.text_area("📍 Site Address", key=f"wo_mgr_site_v{form_version}", placeholder="Enter the full site address", height=90)
                customer_job_no = st.text_input("📘 Customer Job No.", key=f"wo_mgr_cj_v{form_version}", placeholder="Enter Customer Job No.")
            with c2:
                work_date = st.date_input("📅 Date", key=f"wo_mgr_date_v{form_version}", value=date.today())
                amount = st.number_input("💷 Amount (£)", min_value=0.01, step=1.0, format="%.2f", key=f"wo_mgr_wo_amt_v{form_version}")
                description = st.text_area("📝 Description", key=f"wo_mgr_wo_desc_v{form_version}", placeholder="Describe the work completed...", height=150)
            _all_users = load_users()
            director_names = sorted({str(u.get("full_name", "")).strip() for u in _all_users.values() if str(u.get("role", "")).strip().lower() == "director" and str(u.get("full_name", "")).strip()})
            director = director_names[0] if director_names else "Director"
            files = st.file_uploader("📎 Supporting Work Order Document (optional)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key=f"wo_mgr_wo_files_v{form_version}")
            submitted = st.form_submit_button("📤 Submit Work Order", type="primary", width="stretch")
        if submitted:
            errors = []
            if not work_order_no.strip(): errors.append("Work Order No.")
            if not contractor_employee.strip(): errors.append("Contractor / Employee Labour")
            if not site_address.strip(): errors.append("Site Address")
            if not customer_job_no.strip(): errors.append("Customer Job No.")
            if not description.strip(): errors.append("Description")
            duplicate = any(str(r.get("manual_work_order_no", "")).strip().lower() == work_order_no.strip().lower() for r in orders if str(r.get("manual_work_order_no", "")).strip())
            if duplicate: errors.append("Work Order No. already exists")
            if errors: st.error("Please correct: " + ", ".join(errors) + ".")
            else:
                wid = get_next_work_order_id(orders)
                attachments = []
                for i, f in enumerate(files or [], 1):
                    safe_name = os.path.basename(f.name).replace("/", "_").replace("\\", "_")
                    fn = f"{wid}_F{i}_{safe_name}"
                    fp = os.path.join(UPLOAD_DIR, fn)
                    with open(fp, "wb") as out_file: out_file.write(f.getbuffer())
                    _upload_to_drive_bg(fp, fn)
                    attachments.append(fn)
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                rec = {"id": wid, "manual_work_order_no": work_order_no.strip(), "emp_name": contractor_employee.strip(), "dept": manager_dept, "work_date": str(work_date), "hours": 0.0, "customer_job_no": customer_job_no.strip(), "site_address": site_address.strip(), "amount": float(amount), "manager": manager_name, "desc": description.strip(), "attachment_name": ", ".join(attachments) or "None", "status": "pending_director", "manager_comments": "", "manager_decision_date": "", "manager_decision_by": "", "director_comments": "", "director_decision_date": "", "director_decision_by": "", "submitted_by": manager_name, "submitted_date": now, "payroll_status": "Pending", "payroll_date": "", "payroll_by": "", "pdf_path": ""}
                orders.append(rec)
                save_all_work_orders(orders)
                log_action("WORK_ORDER_MANAGER_CREATED", wid, decision_by=manager_name)
                st.session_state["wo_mgr_new_wo_form_version"] = form_version + 1
                st.success(f"✅ Work Order {work_order_no.strip()} submitted to {director} for final approval.")
                st.rerun()
        st.divider()
        wo_pending, wo_approved, wo_rejected = st.tabs([f"⏳ Pending ({len(pending)})", f"✅ Approved ({len(approved)})", f"❌ Returned / Rejected ({len(rejected)})"])
        def filter_orders(items, key):
            q = st.text_input("🔎 Search Work Orders", placeholder="Search by Work Order No., employee, date, amount or description...", key=key).strip().lower()
            if not q: return items
            return [r for r in items if q in " ".join(str(v) for v in r.values()).lower()]
        def show_details(r):
            wo = r.get("manual_work_order_no") or r.get("id") or "-"
            st.write(f"🧾 **Work Order No.: {wo}** | 👤 **Contractor / Employee Labour:** {r.get('emp_name','-')}")
            st.write(f"📍 **Site Address:** {r.get('site_address','-')}")
            st.write(f"📘 **Customer Job No.:** {r.get('customer_job_no','-')} | 📅 **Date:** {r.get('work_date','-')}")
            st.write(f"💷 **Amount:** £{float(r.get('amount',0) or 0):.2f}")
            submitter = str(r.get("submitted_by", "") or "").strip() or "Unknown"
            sub_date = str(r.get("submitted_date", "") or "").strip()
            st.write(f"📝 **Submitted by:** {submitter}" + (f" on {sub_date}" if sub_date else ""))
            st.info(f"📝 **Description:**\n{r.get('desc','')}")
            st.write(f"👔 **Manager:** {r.get('manager') or manager_name}")
            if r.get("manager_decision_by"): st.write(f"👔 **Manager Review:** {r.get('manager_decision_by')} on {r.get('manager_decision_date','')}")
            if r.get("manager_comments"): st.info(f"💬 **Manager Comments:** {r.get('manager_comments')}")
            if r.get("director_decision_by"): st.write(f"🎯 **Approved By:** {r.get('director_decision_by')} on {r.get('director_decision_date','')}")
            if r.get("director_comments"): st.warning(f"💬 Director Comments: {r.get('director_comments')}")
            display_attachments(r)
        with wo_pending:
            items = filter_orders(pending, "wo_mgr_pending_final_search")
            if not items: st.info("⏳ No pending work orders.")
            for r in reversed(items):
                wo = r.get("manual_work_order_no") or r.get("id")
                st_raw = str(r.get("status", "")).strip().lower()
                submitter = str(r.get("submitted_by", "") or "").strip() or "Unknown"
                label = "AWAITING MANAGER" if st_raw == "pending_manager" else ("AWAITING DIRECTOR" if st_raw == "pending_director" else st_raw.upper())
                with st.expander(f"🟡 {wo} | {r.get('emp_name')} | £{float(r.get('amount',0) or 0):.2f} | {label} | Submitted by: {submitter}"):
                    show_details(r)
                    st.divider()
                    if st_raw in ("pending_manager", "pending"):
                        st.info(f"👤 This work order was submitted by **{submitter}** and is awaiting **your review**.")
                        mgr_comment = st.text_area("Manager Comments (optional)", key=f"wo_mgr_comment_{r.get('id')}", placeholder="Any notes for the Director or employee...")
                        c1, c2 = st.columns(2)
                        with c1:
                            if st.button("✅ Approve & Send to Director", key=f"wo_mgr_approve_{r.get('id')}", type="primary"):
                                for x in orders:
                                    if str(x.get("id")) == str(r.get("id")):
                                        x["status"] = "pending_director"; x["manager_comments"] = mgr_comment.strip(); x["manager_decision_by"] = manager_name; x["manager_decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S"); x["pdf_path"] = ""
                                        break
                                save_all_work_orders(orders)
                                log_action("WORK_ORDER_MANAGER_APPROVED", r.get("id"), decision_by=manager_name)
                                st.success(f"✅ Work Order {wo} approved and sent to Director.")
                                st.rerun()
                        with c2:
                            if st.button("❌ Reject & Return to Employee", key=f"wo_mgr_reject_{r.get('id')}"):
                                for x in orders:
                                    if str(x.get("id")) == str(r.get("id")):
                                        x["status"] = "returned_to_employee"; x["manager_comments"] = mgr_comment.strip(); x["manager_decision_by"] = manager_name; x["manager_decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S"); x["pdf_path"] = ""
                                        break
                                save_all_work_orders(orders)
                                log_action("WORK_ORDER_MANAGER_REJECTED", r.get("id"), decision_by=manager_name)
                                st.warning(f"❌ Work Order {wo} returned to employee for corrections.")
                                st.rerun()
                    else:
                        st.caption("Awaiting Director approval. You may still edit this record.")
                        if st.button(f"✏️ Edit Work Order {wo}", key=f"wo_mgr_pending_edit_{r.get('id')}"):
                            st.session_state.editing_work_order_id = r.get("id"); st.rerun()
        with wo_approved:
            items = filter_orders(approved, "wo_mgr_approved_final_search")
            if not items: st.info("✅ No approved work orders.")
            else:
                for r in reversed(items):
                    wo = r.get("manual_work_order_no") or r.get("id")
                    submitter = str(r.get("submitted_by", "") or "").strip() or "Unknown"
                    with st.expander(f"✅ {wo} | {r.get('emp_name')} | £{float(r.get('amount',0) or 0):.2f} | Submitted by: {submitter}"):
                        show_details(r)
                        st.success("Approved — read only.")
                        st.divider(); st.markdown("#### 📄 Work Order PDF"); display_work_order_pdf(r)
                st.divider()
                render_work_order_bulk_download(load_work_orders(), key_prefix="wo_mgr_bulk")
        with wo_rejected:
            items = filter_orders(rejected, "wo_mgr_rejected_final_search")
            if not items: st.info("❌ No rejected or returned work orders.")
            for r in reversed(items):
                wo = r.get("manual_work_order_no") or r.get("id")
                st_raw = str(r.get("status", "")).strip().lower()
                submitter = str(r.get("submitted_by", "") or "").strip() or "Unknown"
                lbl = "RETURNED TO EMPLOYEE" if st_raw == "returned_to_employee" else "REJECTED BY DIRECTOR"
                with st.expander(f"❌ {wo} | {r.get('emp_name')} | £{float(r.get('amount',0) or 0):.2f} | {lbl} | Submitted by: {submitter}"):
                    show_details(r)
                    st.warning("Editable — you can update and resubmit.")
                    if st.button(f"✏️ Edit & Resubmit {wo}", key=f"wo_mgr_rejected_edit_{r.get('id')}"):
                        st.session_state.editing_work_order_id = r.get("id"); st.rerun()
    if can_access_wo_total:
        with total_tab:
            if show_total:
                renderer = globals().get("render_work_order_total")
                if renderer: renderer(load_work_orders(), scope_department=manager_dept, key_prefix="wo_mgr_total_tab", prepared_by=manager_name)
                else: st.info("Approved Work Order Total is available in this tab.")

def render_work_order_director_portal(director_name):
    st.subheader("🛠️ Work Orders — Director Final Approval")
    st.info("✅ Review all work orders, Approve, Reject, OR Change Status at any time. All changes are logged.")
    orders = load_work_orders()
    pending = [r for r in orders if r.get("status") == "pending_director"]
    approved = [r for r in orders if r.get("status") == "approved_payment"]
    rejected = [r for r in orders if r.get("status") == "rejected_director"]
    user_info = st.session_state.get("user_info", {})
    can_access_wo_total = user_info.get("can_access_wo_total", True)
    if can_access_wo_total:
        t1, t2, t3, t4 = st.tabs([
            f"⏳ Manager Approved / Awaiting Director ({len(pending)})",
            f"✅ Approved for Payment ({len(approved)})",
            f"❌ Rejected ({len(rejected)})",
            "💷 Approved Work Order Total"
        ])
    else:
        t1, t2, t3 = st.tabs([
            f"⏳ Manager Approved / Awaiting Director ({len(pending)})",
            f"✅ Approved for Payment ({len(approved)})",
            f"❌ Rejected ({len(rejected)})"
        ])
    def search_list(items, key):
        q = st.text_input("🔎 Search Work Orders", placeholder="Search by Work Order No., employee, manager, submitter, amount, date, customer job no. or description...", key=key)
        if q.strip():
            q = q.lower().strip()
            items = [r for r in items if q in " ".join(str(v) for v in r.values()).lower()]
        return items
    def show_full_details(r):
        st.write(f"🧾 **Work Order No.:** {get_work_order_number(r)}")
        st.write(f"👤 **Contractor / Employee Labour:** {r.get('emp_name','-')}")
        st.write(f"🏢 **Department:** {r.get('dept','-')}")
        st.write(f"📍 **Site Address:** {r.get('site_address','-')}")
        st.write(f"📘 **Customer Job No.:** {r.get('customer_job_no','-')}")
        st.write(f"📅 **Work Date:** {r.get('work_date','-')}")
        st.write(f"💷 **Amount:** £{float(r.get('amount',0) or 0):.2f}")
        st.write(f"👔 **Manager:** {r.get('manager','-')}")
        st.write(f"📝 **Submitted by:** {r.get('submitted_by','-')} on {r.get('submitted_date','-')}")
        if r.get('manager_decision_by'): st.write(f"👔 **Manager review:** {r.get('manager_decision_by')} on {r.get('manager_decision_date','')}")
        st.info(f"📝 **Description:**\n{r.get('desc','')}")
        if r.get('manager_comments'): st.info(f"💬 **Manager Comments:** {r.get('manager_comments')}")
        if r.get('director_comments'): st.warning(f"💬 **Director Comments:** {r.get('director_comments')}")
        if r.get('director_decision_by'): st.write(f"🎯 **Approved By:** {r.get('director_decision_by')} on {r.get('director_decision_date','')}")
        st.divider()
        st.markdown("#### 📎 Attachments")
        display_attachments(r)
    def apply_status_change(req_id, new_status, comments, old_status):
        for x in orders:
            if str(x.get("id")) == str(req_id):
                x["status"] = new_status
                if new_status == "approved_payment":
                    x["director_comments"] = comments.strip() if comments.strip() else x.get("director_comments", "")
                    x["director_decision_by"] = director_name
                    x["director_decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    x["pdf_path"] = ""
                elif new_status == "rejected_director":
                    x["director_comments"] = comments.strip() if comments.strip() else x.get("director_comments", "")
                    x["director_decision_by"] = director_name
                    x["director_decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    x["pdf_path"] = ""
                elif new_status == "pending_director":
                    x["director_comments"] = (str(x.get("director_comments","")) + f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M')}] ⏳ Changed to Pending by {director_name}: {comments.strip()}").strip()
                    x["director_decision_by"] = ""
                    x["director_decision_date"] = ""
                    x["pdf_path"] = ""
                break
        save_all_work_orders(orders)
        log_action("WORK_ORDER_STATUS_CHANGED", req_id,
                   old_data={"status": old_status}, new_data={"status": new_status},
                   decision_by=director_name)
        st.success(f"✅ Work Order #{req_id} status changed to **{new_status.replace('_',' ').title()}**.")
    with t1:
        items = search_list(pending, "wo_dir_pending_search")
        if not items:
            st.info("⏳ No work orders awaiting Director approval.")
        for r in reversed(items):
            req_id = r.get("id")
            submitter = r.get('submitted_by') or r.get('manager') or "Unknown"
            with st.expander(
                f"🟡 {get_work_order_number(r)} | {r.get('emp_name')} | 🔄 Work Order | "
                f"£{r.get('amount',0):.2f} | Submitted by: {submitter}"
            ):
                show_full_details(r)
                st.markdown("### ✍️ Director Decision")
                comments = st.text_area("Director Comments", key=f"wo_dir_comm_{req_id}")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("✅ Approve for Payment", key=f"wo_dir_app_{req_id}", type="primary"):
                        apply_status_change(req_id, "approved_payment", comments, "pending_director")
                        st.rerun()
                with c2:
                    if st.button("❌ Reject", key=f"wo_dir_rej_{req_id}"):
                        if not comments.strip():
                            st.error("❌ A rejection reason/comment is required.")
                        else:
                            apply_status_change(req_id, "rejected_director", comments, "pending_director")
                            st.rerun()
    with t2:
        items = search_list(approved, "wo_dir_approved_search")
        if not items:
            st.info("✅ No Director-approved work orders.")
        st.info("🔄 **Change Status:** Move back to Pending ❘ Change to Rejected")
        for r in reversed(items):
            req_id = r.get("id")
            with st.expander(
                f"🟢 {get_work_order_number(r)} | {r.get('emp_name')} | 🔄 Work Order | "
                f"£{r.get('amount',0):.2f} | Approved by {r.get('director_decision_by','')}"
            ):
                show_full_details(r)
                if r.get("director_comments"):
                    st.error(f"❌ **Rejection reason:** {r.get('director_comments')}")
                st.markdown("### 🔄 Change Status")
                new_comments = st.text_area(
                    "Add comment (optional)", key=f"wo_dir_chg_comm_app_{req_id}",
                    placeholder="Reason for status change..."
                )
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("⏳ Move to Pending", key=f"wo_dir_app_to_pend_{req_id}"):
                        apply_status_change(req_id, "pending_director", new_comments, "approved_payment")
                        st.rerun()
                with c2:
                    if st.button("❌ Change to Rejected", key=f"wo_dir_app_to_rej_{req_id}"):
                        if not new_comments.strip():
                            st.error("❌ A rejection reason/comment is required.")
                        else:
                            apply_status_change(req_id, "rejected_director", new_comments, "approved_payment")
                            st.rerun()
                st.divider()
                st.markdown("#### 📄 Work Order PDF")
                display_work_order_pdf(r)
        if items:
            st.divider()
            render_work_order_bulk_download(load_work_orders(), key_prefix="wo_dir_bulk")
    with t3:
        items = search_list(rejected, "wo_dir_rejected_search")
        if not items:
            st.info("❌ No rejected work orders.")
        st.info("🔄 **Change Status:** Move back to Pending ❘ Change to Approved")
        for r in reversed(items):
            req_id = r.get("id")
            with st.expander(
                f"🔴 {get_work_order_number(r)} | {r.get('emp_name')} | 🔄 Work Order | "
                f"£{r.get('amount',0):.2f} | Rejected by {r.get('director_decision_by','')}"
            ):
                show_full_details(r)
                st.markdown("### 🔄 Change Status")
                new_comments = st.text_area(
                    "Add comment (optional)", key=f"wo_dir_chg_comm_rej_{req_id}",
                    placeholder="Reason for status change..."
                )
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("⏳ Move to Pending", key=f"wo_dir_rej_to_pend_{req_id}"):
                        apply_status_change(req_id, "pending_director", new_comments, "rejected_director")
                        st.rerun()
                with c2:
                    if st.button("✅ Change to Approved", key=f"wo_dir_rej_to_app_{req_id}", type="primary"):
                        apply_status_change(req_id, "approved_payment", new_comments, "rejected_director")
                        st.rerun()
    if can_access_wo_total:
        with t4:
            render_work_order_total(orders, key_prefix="wo_dir_total", prepared_by=director_name)

def render_work_order_payroll_portal(payroll_name):
    st.subheader("🛠️ Work Orders — Payroll")
    st.info("View Director-approved work orders and download the authorised PDF. Payroll does not change the work order status.")
    orders = load_work_orders()
    approved = [r for r in orders if r.get("status") == "approved_payment"]
    search = st.text_input("🔎 Search approved work orders", placeholder="Search by ID, employee, manager, director, department, amount or date...", key="wo_payroll_search")
    if search.strip():
        q = search.lower().strip()
        approved = [r for r in approved if q in " ".join(str(v) for v in r.values()).lower()]
    st.metric("✅ Approved for Payment", len(approved))
    st.divider()
    if not approved:
        st.success("✅ No Director-approved work orders found.")
        return
    for r in reversed(approved):
        with st.expander(f"🟢 {r.get('id')} | {r.get('emp_name')} | £{r.get('amount', 0):.2f} | Approved by {r.get('director_decision_by')}"):
            st.write(f"🧾 Work Order No.: **{get_work_order_number(r)}**")
            st.write(f"📝 Submitted by: {r.get('submitted_by')} | 👔 Manager: {r.get('manager')} | 🎯 Director: {r.get('director_decision_by')}")
            st.write(f"📅 Approval: {r.get('director_decision_date')} | 💷 £{r.get('amount', 0):.2f}")
            st.info(r.get('desc', ''))
            display_attachments(r)
            display_work_order_pdf(r)
    st.divider()
    render_work_order_bulk_download(load_work_orders(), key_prefix="wo_pay_bulk")
    st.divider()
    user_info = st.session_state.get("user_info", {})
    if user_info.get("can_access_wo_total", False):
        render_work_order_total(approved, key_prefix="wo_pay_total", prepared_by=payroll_name)

def initialise_inspector_bonus():
    safe_init_excel(INSPECTOR_BONUS_PATH, INSPECTOR_BONUS_COLUMNS)

def load_inspector_bonus(force=False):
    if not force and "_inspector_bonus_cache" in st.session_state:
        return list(st.session_state["_inspector_bonus_cache"])
    initialise_inspector_bonus()
    try:
        df = _read_excel_records(INSPECTOR_BONUS_PATH)
        records = []
        for r in df.to_dict(orient="records"):
            try: amount = float(r.get("Bonus Amount (£)", 0) or 0)
            except: amount = 0.0
            try: jobs = float(r.get("Total Jobs Completed", 0) or 0)
            except: jobs = 0.0
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
                "final_holiday_settlement": str(r.get("Final Holiday Settlement", "")).strip().casefold() in {"yes", "true", "1"},
            })
        _set_data_cache("_inspector_bonus_cache", records)
        return list(records)
    except Exception as e:
        st.error(f"Inspector Bonus Load Error: {e}")
        return []

def save_all_inspector_bonus(records, sync=True):
    rows = []
    for r in records:
        rows.append({"ID": str(r.get("id", "")), "Inspector Name": str(r.get("inspector_name", "")), "Month & Year": str(r.get("month_year", "")), "Days Absent": str(r.get("days_absent", "")), "Reasons for Absence": str(r.get("reasons", "")), "Total Jobs Completed": float(r.get("total_jobs", 0)), "Bonus Amount (£)": float(r.get("bonus_amount", 0)), "Status": str(r.get("status", "pending_director")), "Director Comments": str(r.get("director_comments", "")), "Director Decision Date": str(r.get("director_decision_date", "")), "Director Decision By": str(r.get("director_decision_by", "")), "Submitted By": str(r.get("submitted_by", "")), "Submitted Date": str(r.get("submitted_date", "")), "PDF File Path": str(r.get("pdf_path", ""))})
    pd.DataFrame(rows, columns=INSPECTOR_BONUS_COLUMNS).to_excel(INSPECTOR_BONUS_PATH, index=False, engine="openpyxl")
    _set_data_cache("_inspector_bonus_cache", list(records))
    if sync: sync_saved_file_to_drive(INSPECTOR_BONUS_PATH)

def get_next_inspector_bonus_id(records):
    nums = []
    for r in records:
        raw = str(r.get("id", ""))
        try: nums.append(int(raw.replace("IB-", "")))
        except: pass
    return f"IB-{max(nums) + 1 if nums else 1:04d}"

def inspector_bonus_pdf(req, force_regenerate=False, upload_to_drive=True):
    if not PDF_AVAILABLE: return None
    try:
        cached_path = req.get("pdf_path", "")
        if not force_regenerate and cached_path and os.path.exists(cached_path): return cached_path
        pdf = FPDF()
        pdf.add_page()
        regular_font, bold_font = _pdf_font_paths()
        if regular_font and bold_font:
            pdf.add_font("DejaVu", "", regular_font); pdf.add_font("DejaVu", "B", bold_font); family = "DejaVu"
        else: family = "Helvetica"
        def safe(v):
            text = _pdf_text(v)
            if family == "Helvetica": return text.encode("latin-1", "replace").decode("latin-1")
            return text
        if os.path.exists(LOGO_PATH):
            try:
                pdf.image(LOGO_PATH, x=75, y=10, w=60); pdf.ln(28)
            except Exception: pdf.ln(5)
        else: pdf.ln(5)
        pdf.set_font(family, "B", 16)
        pdf.cell(0, 10, safe("National Grid Inspector Bonus Approval Sheet"), ln=True, align="C")
        pdf.ln(4)
        pdf.set_font(family, "", 10)
        pdf.multi_cell(0, 6, safe("This sheet needs to be completed and passed to Andy to be signed off and given to Rachel by the 3rd of the month."), align="C")
        pdf.ln(8)
        def field(label, value, label_w=60, value_h=10):
            pdf.set_font(family, "B", 11); pdf.cell(label_w, value_h, safe(label), border=1)
            pdf.set_font(family, "", 11); pdf.cell(0, value_h, safe("   " + str(value)), border=1, ln=True); pdf.ln(2)
        field("Inspector Name:", req.get("inspector_name", ""))
        field("Month & Year:", req.get("month_year", ""))
        field("Days Absent:", req.get("days_absent", "") or "-")
        field("Reasons for Absence:", req.get("reasons", "") or "-")
        field("Total Jobs Completed", f"{float(req.get('total_jobs', 0)):.2f}")
        field("Bonus Amount:", f"£{float(req.get('bonus_amount', 0)):.2f}")
        status = str(req.get("status", "")).strip().lower()
        text_content = ""; stamp_path = None
        if status == "approved":
            approved_by = req.get("director_decision_by", "") or "Director"
            decision_date = req.get("director_decision_date", "")
            text_content = f"   {approved_by} on {decision_date}"
            stamp_path = APPROVED_STAMP_PATH
        elif status == "rejected":
            text_content = "   Rejected"; stamp_path = REJECTED_STAMP_PATH
        else: text_content = "   (Pending Director signature)"
        pdf.set_font(family, "B", 11); pdf.cell(60, 10, safe("Approved By:"), border=1)
        pdf.set_font(family, "", 11); pdf.cell(80, 10, safe(text_content), border=1)
        current_x = pdf.get_x(); current_y = pdf.get_y()
        pdf.cell(50, 10, "", border=1, ln=True)
        if stamp_path and os.path.exists(stamp_path):
            try: pdf.image(stamp_path, x=current_x + 5, y=current_y - 1, w=40)
            except Exception: pass
        pdf.ln(4)
        if req.get("director_comments"):
            pdf.set_font(family, "B", 10); pdf.cell(0, 6, safe("Director Comments:"), ln=True)
            pdf.set_font(family, "", 10); pdf.set_x(pdf.l_margin); pdf.multi_cell(0, 6, safe(req.get("director_comments", ""))); pdf.ln(4)
        pdf.set_font(family, "", 9)
        pdf.cell(0, 5, safe(f"Submitted By: {req.get('submitted_by', '')}"), ln=True)
        pdf.cell(0, 5, safe(f"Submitted Date: {req.get('submitted_date', '')}"), ln=True)
        pdf.cell(0, 5, safe(f"Record ID: {req.get('id', '')}"), ln=True)
        os.makedirs(INSPECTOR_BONUS_PDF_DIR, exist_ok=True)
        safe_id = "".join(str(req.get("id", "IB")).split()) or "IB"
        safe_name = "_".join(str(req.get("inspector_name", "Inspector")).split()) or "Inspector"
        safe_month = "_".join(str(req.get("month_year", "")).split()) or "Month"
        filename = f"Inspector_Bonus_{safe_id}_{safe_name}_{safe_month}.pdf"
        path = os.path.join(INSPECTOR_BONUS_PDF_DIR, filename)
        pdf.output(path)
        if upload_to_drive: _upload_to_drive_bg(path, os.path.basename(path))
        return path
    except Exception as e:
        st.error(f"Inspector Bonus PDF Error: {e}")
        return None

def display_inspector_bonus_pdf_button(req, key_prefix="ib"):
    if not PDF_AVAILABLE:
        st.warning("⚠️ PDF generation is unavailable. Please install fpdf2.")
        return
    status = str(req.get("status", "")).strip().lower()
    if status != "approved":
        st.info("📄 PDF download is available once the Director approves this bonus."); return
    req_id = str(req.get("id", "unknown"))
    dl_key = f"{key_prefix}_dl_{req_id}"
    gen_key = f"{key_prefix}_gen_{req_id}"
    cached_path = req.get("pdf_path", "")
    if cached_path and os.path.exists(cached_path):
        with open(cached_path, "rb") as f:
            st.download_button("⬇️ Download Bonus Approval PDF", data=f.read(), file_name=os.path.basename(cached_path), mime="application/pdf", type="primary", key=dl_key)
        return
    if st.button(f"📄 Generate PDF for {req_id}", key=gen_key, type="primary"):
        with st.spinner("Generating PDF..."):
            path = inspector_bonus_pdf(req, force_regenerate=True, upload_to_drive=True)
        if path and os.path.exists(path):
            records = load_inspector_bonus()
            for r in records:
                if str(r.get("id")) == req_id: r["pdf_path"] = path
            save_all_inspector_bonus(records, sync=False)
            st.success("✅ PDF generated. Click below to download."); st.rerun()
        else: st.error("❌ Could not generate PDF.")

def render_inspector_bonus_super_admin():
    st.subheader("💰 National Grid Inspector Bonus — Super Admin (View Only)")
    st.info("✅ View all Inspector Bonus records across ALL departments. **Approval → Director only.**")
    records = load_inspector_bonus()
    search = st.text_input("🔎 Search Inspector Bonuses", placeholder="Search by ID, inspector, month, status, amount...", key="ib_super_admin_search")
    if search.strip():
        q = search.lower().strip()
        records = [r for r in records if q in " ".join(str(v) for v in r.values()).lower()]
    pending = [r for r in records if str(r.get("status", "")).strip().lower() == "pending_director"]
    approved = [r for r in records if str(r.get("status", "")).strip().lower() == "approved"]
    rejected = [r for r in records if str(r.get("status", "")).strip().lower() == "rejected"]
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("🟡 Pending", len(pending))
    with c2: st.metric("🟢 Approved", len(approved))
    with c3: st.metric("🔴 Rejected", len(rejected))
    st.divider()
    t1, t2, t3 = st.tabs([f"⏳ Pending ({len(pending)})", f"✅ Approved ({len(approved)})", f"❌ Rejected ({len(rejected)})"])
    def show_ib_details(r):
        st.write(f"**Inspector:** {r.get('inspector_name')} | **Month:** {r.get('month_year')}")
        st.write(f"**Days Absent:** {r.get('days_absent') or '-'}")
        st.write(f"**Reasons:** {r.get('reasons') or '-'}")
        st.write(f"**Total Jobs Completed:** {r.get('total_jobs')}")
        st.write(f"**Bonus Amount:** £{r.get('bonus_amount',0):.2f}")
        st.caption(f"Submitted by {r.get('submitted_by')} on {r.get('submitted_date')}")
        if r.get("director_decision_by"): st.write(f"**Director:** {r.get('director_decision_by')} on {r.get('director_decision_date','')}")
        if r.get("director_comments"): st.info(f"💬 **Director Comments:** {r.get('director_comments')}")
    with t1:
        if not pending: st.success("✅ No pending Inspector Bonus records.")
        for r in reversed(pending):
            rec_id = r.get("id")
            with st.expander(f"🟡 {rec_id} | {r.get('inspector_name')} | {r.get('month_year')} | £{r.get('bonus_amount',0):.2f}"):
                show_ib_details(r)
    with t2:
        if not approved: st.info("✅ No approved Inspector Bonus records.")
        for r in reversed(approved):
            rec_id = r.get("id")
            with st.expander(f"🟢 {rec_id} | {r.get('inspector_name')} | {r.get('month_year')} | £{r.get('bonus_amount',0):.2f}"):
                show_ib_details(r); st.divider(); display_inspector_bonus_pdf_button(r, key_prefix="ib_super_admin")
    with t3:
        if not rejected: st.info("❌ No rejected Inspector Bonus records.")
        for r in reversed(rejected):
            rec_id = r.get("id")
            with st.expander(f"🔴 {rec_id} | {r.get('inspector_name')} | {r.get('month_year')} | £{r.get('bonus_amount',0):.2f}"):
                show_ib_details(r)

def render_inspector_bonus_portal(user_name, user_dept):
    st.subheader("💰 National Grid Inspector Bonus Approval")
    st.info("This sheet needs to be completed and passed to Andy to be signed off and given to Rachel by the 3rd of the month.")
    st.divider()
    bonus_records = load_inspector_bonus()
    editing_id = st.session_state.get("editing_inspector_bonus_id")
    if editing_id:
        rec = next((r for r in bonus_records if str(r.get("id")) == str(editing_id)), None)
        if rec and str(rec.get("status", "")).strip().lower() == "pending_director":
            st.warning(f"✏️ **Editing Bonus Record {editing_id}** — save changes to resubmit to the Director.")
            with st.form("inspector_bonus_edit_form", clear_on_submit=False):
                e1, e2 = st.columns(2)
                with e1:
                    e_inspector = st.text_input("Inspector Name:", value=rec.get("inspector_name", ""))
                    e_month = st.text_input("Month & Year:", value=rec.get("month_year", ""))
                    e_days = st.text_input("Days Absent:", value=rec.get("days_absent", ""))
                    e_reasons = st.text_area("Reasons for Absence:", value=rec.get("reasons", ""), height=100)
                with e2:
                    e_jobs = st.number_input("Total Jobs Completed:", min_value=0.0, step=1.0, format="%.2f", value=float(rec.get("total_jobs", 0) or 0))
                    e_amount = st.number_input("Bonus Amount (£):", min_value=0.01, step=1.0, format="%.2f", value=float(rec.get("bonus_amount", 0.01) or 0.01))
                c_save, c_cancel = st.columns(2)
                with c_save: save_edit = st.form_submit_button("💾 Save Changes & Resubmit", type="primary", width="stretch")
                with c_cancel: cancel_edit = st.form_submit_button("❌ Cancel", width="stretch")
                if save_edit:
                    if not e_inspector.strip() or not e_month.strip(): st.error("❌ Inspector Name and Month & Year are required.")
                    else:
                        old_data = {"inspector_name": rec.get("inspector_name"), "month_year": rec.get("month_year"), "days_absent": rec.get("days_absent"), "reasons": rec.get("reasons"), "total_jobs": rec.get("total_jobs"), "bonus_amount": rec.get("bonus_amount")}
                        for x in bonus_records:
                            if str(x.get("id")) == str(editing_id):
                                x["inspector_name"] = e_inspector.strip(); x["month_year"] = e_month.strip(); x["days_absent"] = e_days.strip(); x["reasons"] = e_reasons.strip(); x["total_jobs"] = float(e_jobs); x["bonus_amount"] = float(e_amount); x["status"] = "pending_director"; x["director_comments"] = ""; x["director_decision_date"] = ""; x["director_decision_by"] = ""; x["pdf_path"] = ""; x["submitted_by"] = user_name; x["submitted_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                break
                        save_all_inspector_bonus(bonus_records)
                        new_data = {"inspector_name": e_inspector.strip(), "month_year": e_month.strip(), "days_absent": e_days.strip(), "reasons": e_reasons.strip(), "total_jobs": float(e_jobs), "bonus_amount": float(e_amount)}
                        log_action("INSPECTOR_BONUS_EDITED", editing_id, old_data=old_data, new_data=new_data)
                        st.session_state.editing_inspector_bonus_id = None
                        st.success(f"✅ Bonus record {editing_id} updated and re-sent to Director."); st.rerun()
                if cancel_edit:
                    st.session_state.editing_inspector_bonus_id = None; st.rerun()
            st.divider()
        else: st.session_state.editing_inspector_bonus_id = None
    new_id = get_next_inspector_bonus_id(bonus_records)
    with st.form("inspector_bonus_form", clear_on_submit=True):
        st.markdown(f"**🔐 Internal Record ID:** `{new_id}`")
        c1, c2 = st.columns(2)
        with c1:
            inspector_name = st.text_input("Inspector Name:", placeholder="e.g. Abdul Ali")
            month_year = st.text_input("Month & Year:", placeholder="e.g. Jul-26")
            days_absent = st.text_input("Days Absent:", placeholder="e.g. 2 days, None")
            reasons = st.text_area("Reasons for Absence:", placeholder="Reason details...", height=100)
        with c2:
            total_jobs = st.number_input("Total Jobs Completed:", min_value=0.0, step=1.0, format="%.2f")
            bonus_amount = st.number_input("Bonus Amount (£):", min_value=0.01, step=1.0, format="%.2f")
            st.caption("ℹ️ After submission, this will be sent to the Director for approval.")
        submitted = st.form_submit_button("📤 Submit for Director Approval", type="primary", width="stretch")
        if submitted:
            if not inspector_name.strip(): st.error("❌ Inspector Name is required.")
            elif not month_year.strip(): st.error("❌ Month & Year is required.")
            else:
                rec = {"id": new_id, "inspector_name": inspector_name.strip(), "month_year": month_year.strip(), "days_absent": days_absent.strip(), "reasons": reasons.strip(), "total_jobs": float(total_jobs), "bonus_amount": float(bonus_amount), "status": "pending_director", "director_comments": "", "director_decision_date": "", "director_decision_by": "", "submitted_by": user_name, "submitted_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "pdf_path": ""}
                bonus_records.append(rec)
                save_all_inspector_bonus(bonus_records)
                log_action("INSPECTOR_BONUS_CREATED", new_id, new_data=rec)
                st.success(f"✅ Bonus Approval {new_id} for {inspector_name} submitted to Director for approval."); st.rerun()
    st.divider()
    st.subheader("📋 Submitted Bonus Approvals")
    search = st.text_input("🔎 Search records", placeholder="Search by ID, inspector, month, amount...", key="ib_search")
    filtered = list(bonus_records)
    if search.strip():
        q = search.lower().strip()
        filtered = [r for r in filtered if q in " ".join(str(v) for v in r.values()).lower()]
    if not filtered: st.info("📋 No bonus approval records found.")
    else:
        for r in reversed(filtered):
            status_raw = str(r.get("status", "")).strip().lower()
            if status_raw == "approved": icon = "🟢"; status_label = "APPROVED"
            elif status_raw == "rejected": icon = "🔴"; status_label = "REJECTED"
            else: icon = "🟡"; status_label = "PENDING DIRECTOR"
            with st.expander(f"{icon} {r.get('id')} | {r.get('inspector_name')} | {r.get('month_year')} | £{r.get('bonus_amount',0):.2f} | {status_label}"):
                st.write(f"**Inspector:** {r.get('inspector_name')} | **Month:** {r.get('month_year')}")
                st.write(f"**Days Absent:** {r.get('days_absent') or '-'}")
                st.write(f"**Reasons:** {r.get('reasons') or '-'}")
                st.write(f"**Total Jobs Completed:** {r.get('total_jobs')}")
                st.write(f"**Bonus Amount:** £{r.get('bonus_amount',0):.2f}")
                st.write(f"**Status:** {status_label}")
                if r.get("director_decision_by"): st.write(f"**Director:** {r.get('director_decision_by')} on {r.get('director_decision_date','')}")
                if r.get("director_comments"): st.info(f"💬 **Director Comments:** {r.get('director_comments')}")
                st.caption(f"Submitted by {r.get('submitted_by')} on {r.get('submitted_date')}")
                display_inspector_bonus_pdf_button(r, key_prefix="ib_mgr")
                if status_raw == "pending_director":
                    st.divider()
                    if st.button(f"✏️ Edit Bonus Record {r.get('id')}", key=f"ib_edit_{r.get('id')}", type="secondary"):
                        st.session_state.editing_inspector_bonus_id = r.get("id"); st.rerun()

def render_inspector_bonus_director_portal(director_name):
    st.subheader("💰 National Grid Inspector Bonus — Director Approval")
    st.info("Review Inspector Bonus submissions. Approve, Reject, OR Change Status at any time. All changes are logged.")
    st.divider()
    records = load_inspector_bonus()
    pending = [r for r in records if str(r.get("status", "")).strip().lower() == "pending_director"]
    approved = [r for r in records if str(r.get("status", "")).strip().lower() == "approved"]
    rejected = [r for r in records if str(r.get("status", "")).strip().lower() == "rejected"]
    t1, t2, t3 = st.tabs([f"⏳ Pending ({len(pending)})", f"✅ Approved ({len(approved)})", f"❌ Rejected ({len(rejected)})"])
    def apply_decision(rec_id, new_status, comments, old_status):
        for x in records:
            if str(x.get("id")) == str(rec_id):
                x["status"] = new_status
                if new_status == "pending_director":
                    x["director_comments"] = (str(x.get("director_comments","")) + f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M')}] ⏳ Changed to Pending by {director_name}: {comments.strip()}").strip()
                    x["director_decision_by"] = ""; x["director_decision_date"] = ""
                else:
                    x["director_comments"] = comments.strip() if comments.strip() else x.get("director_comments", "")
                    x["director_decision_by"] = director_name
                    x["director_decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                x["pdf_path"] = ""
                break
        save_all_inspector_bonus(records)
        if old_status == "pending_director":
            action = "INSPECTOR_BONUS_APPROVED" if new_status == "approved" else ("INSPECTOR_BONUS_REJECTED" if new_status == "rejected" else "INSPECTOR_BONUS_STATUS_CHANGED")
        else: action = "INSPECTOR_BONUS_STATUS_CHANGED"
        log_action(action, rec_id, old_data={"status": old_status}, new_data={"status": new_status, "comments": comments}, decision_by=director_name)
        st.success(f"✅ Record {rec_id} → **{new_status.replace('_',' ').title()}**.")
    def show_details(r):
        st.write(f"**Inspector:** {r.get('inspector_name')} | **Month:** {r.get('month_year')}")
        st.write(f"**Days Absent:** {r.get('days_absent') or '-'}")
        st.write(f"**Reasons:** {r.get('reasons') or '-'}")
        st.write(f"**Total Jobs Completed:** {r.get('total_jobs')}")
        st.write(f"**Bonus Amount:** £{r.get('bonus_amount',0):.2f}")
        st.caption(f"Submitted by {r.get('submitted_by')} on {r.get('submitted_date')}")
        if r.get("director_decision_by"): st.write(f"**Decision By:** {r.get('director_decision_by')} on {r.get('director_decision_date','')}")
        if r.get("director_comments"): st.info(f"💬 **Director Comments:** {r.get('director_comments')}")
    with t1:
        if not pending: st.success("✅ No pending Inspector Bonus submissions.")
        for r in reversed(pending):
            rec_id = r.get("id")
            with st.expander(f"🟡 {rec_id} | {r.get('inspector_name')} | 🔄 Inspector Bonus | {r.get('month_year')} | £{r.get('bonus_amount',0):.2f}"):
                show_details(r); st.divider()
                comments = st.text_area("Director Comments", key=f"ib_dir_comm_{rec_id}")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("✅ Approve", key=f"ib_dir_app_{rec_id}", type="primary"):
                        apply_decision(rec_id, "approved", comments, "pending_director"); st.rerun()
                with c2:
                    if st.button("❌ Reject", key=f"ib_dir_rej_{rec_id}"):
                        if not comments.strip():
                            st.error("❌ A rejection reason/comment is required.")
                        else:
                            apply_decision(rec_id, "rejected", comments, "pending_director"); st.rerun()
    with t2:
        if not approved: st.info("✅ No approved Inspector Bonus records yet.")
        if approved: st.info("🔄 **Change Status:** Move back to Pending ❘ Change to Rejected")
        for r in reversed(approved):
            rec_id = r.get("id")
            with st.expander(f"🟢 {rec_id} | {r.get('inspector_name')} | {r.get('month_year')} | £{r.get('bonus_amount',0):.2f}"):
                show_details(r); st.divider()
                display_inspector_bonus_pdf_button(r, key_prefix="ib_dir")
                st.divider(); st.markdown("### 🔄 Change Status")
                new_comments = st.text_area("Add comment (optional)", key=f"ib_dir_chg_app_{rec_id}", placeholder="Reason for status change...")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("⏳ Move to Pending", key=f"ib_dir_app_to_pend_{rec_id}"):
                        apply_decision(rec_id, "pending_director", new_comments, "approved"); st.rerun()
                with c2:
                    if st.button("❌ Change to Rejected", key=f"ib_dir_app_to_rej_{rec_id}"):
                        if not new_comments.strip():
                            st.error("❌ A rejection reason/comment is required.")
                        else:
                            apply_decision(rec_id, "rejected", new_comments, "approved"); st.rerun()
    with t3:
        if not rejected: st.info("❌ No rejected Inspector Bonus records.")
        if rejected: st.info("🔄 **Change Status:** Move back to Pending ❘ Change to Approved")
        for r in reversed(rejected):
            rec_id = r.get("id")
            with st.expander(f"🔴 {rec_id} | {r.get('inspector_name')} | {r.get('month_year')} | £{r.get('bonus_amount',0):.2f}"):
                show_details(r); st.divider()
                st.markdown("### 🔄 Change Status")
                new_comments = st.text_area("Add comment (optional)", key=f"ib_dir_chg_rej_{rec_id}", placeholder="Reason for status change...")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("⏳ Move to Pending", key=f"ib_dir_rej_to_pend_{rec_id}"):
                        apply_decision(rec_id, "pending_director", new_comments, "rejected"); st.rerun()
                with c2:
                    if st.button("✅ Change to Approved", key=f"ib_dir_rej_to_app_{rec_id}", type="primary"):
                        apply_decision(rec_id, "approved", new_comments, "rejected"); st.rerun()

def render_inspector_bonus_payroll_portal(payroll_name):
    st.subheader("💰 National Grid Inspector Bonus — Payroll")
    st.info("View Director-approved Inspector Bonuses and download the authorised PDF.")
    st.divider()
    records = load_inspector_bonus()
    approved = [r for r in records if str(r.get("status", "")).strip().lower() == "approved"]
    search = st.text_input("🔎 Search approved Inspector Bonuses", placeholder="Search by ID, inspector, month, amount...", key="ib_payroll_search")
    if search.strip():
        q = search.lower().strip()
        approved = [r for r in approved if q in " ".join(str(v) for v in r.values()).lower()]
    st.metric("✅ Approved Bonuses", len(approved))
    st.divider()
    if not approved:
        st.success("✅ No approved Inspector Bonuses found.")
        return
    for r in reversed(approved):
        rec_id = r.get("id")
        with st.expander(f"🟢 {rec_id} | {r.get('inspector_name')} | {r.get('month_year')} | £{r.get('bonus_amount',0):.2f}"):
            st.write(f"**Inspector:** {r.get('inspector_name')} | **Month:** {r.get('month_year')}")
            st.write(f"**Days Absent:** {r.get('days_absent') or '-'}")
            st.write(f"**Reasons:** {r.get('reasons') or '-'}")
            st.write(f"**Total Jobs Completed:** {r.get('total_jobs')}")
            st.write(f"**Bonus Amount:** £{r.get('bonus_amount',0):.2f}")
            st.write(f"**Approved By:** {r.get('director_decision_by')} on {r.get('director_decision_date','')}")
            if r.get("director_comments"): st.info(f"💬 **Director Comments:** {r.get('director_comments')}")
            st.caption(f"Submitted by {r.get('submitted_by')} on {r.get('submitted_date')}")
            st.divider()
            display_inspector_bonus_pdf_button(r, key_prefix="ib_payroll")


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
            try: rid = int(r.get("ID", 0))
            except Exception: rid = 0
            try: amount = float(r.get("Amount (£)", 0) or 0)
            except Exception: amount = 0.0
            try: days = float(r.get("Number of Days", 0) or 0)
            except Exception: days = 0.0
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
        backup_ok = sync_saved_file_to_drive(HR_LEAVE_PATH)
        if not backup_ok:
            print("WARNING: HR Leave Settlement local save succeeded, but Google Drive backup did not complete.")

def get_next_hr_leave_id(records):
    if not records: return 1
    return max(int(r.get("id", 0)) for r in records) + 1

def load_hr_daily_rates(force=False):
    if not force and "_hr_daily_rates_cache" in st.session_state:
        return list(st.session_state["_hr_daily_rates_cache"])
    initialise_hr_leave()
    try:
        df = _read_excel_records(HR_DAILY_RATES_PATH)
        records = []
        for r in df.to_dict(orient="records"):
            try: rate = float(r.get("Daily Rate (£)", 0) or 0)
            except Exception: rate = 0.0
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
    safe_init_excel(STORE_ITEMS_PATH, STORE_ITEMS_COLUMNS)
    try:
        df_items = _read_excel_records(STORE_ITEMS_PATH)
        if df_items.empty:
            pd.DataFrame(DEFAULT_STORE_ITEMS).to_excel(STORE_ITEMS_PATH, index=False, engine="openpyxl")
            sync_saved_file_to_drive(STORE_ITEMS_PATH)
    except Exception:
        pass

def load_store_deductions(force=False):
    if not force and "_store_deduction_cache" in st.session_state:
        return list(st.session_state["_store_deduction_cache"])
    initialise_store_deduction()
    try:
        df = _read_excel_records(STORE_DEDUCTION_PATH)
        records = []
        for r in df.to_dict(orient="records"):
            try: rid = int(r.get("ID", 0))
            except Exception: rid = 0
            try: total = float(r.get("Total Deduction (£)", 0) or 0)
            except Exception: total = 0.0
            try: items_json = r.get("Items Deducted JSON", "[]")
            except Exception: items_json = "[]"
            try: items = json.loads(items_json) if items_json else []
            except Exception: items = []
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
    if sync: sync_saved_file_to_drive(STORE_DEDUCTION_PATH)

def get_next_store_deduction_id(records):
    if not records: return 1
    return max(int(r.get("id", 0)) for r in records) + 1

def load_store_items(force=False):
    if not force and "_store_items_cache" in st.session_state:
        return list(st.session_state["_store_items_cache"])
    initialise_store_deduction()
    try:
        df = _read_excel_records(STORE_ITEMS_PATH)
        records = []
        for r in df.to_dict(orient="records"):
            try: price = float(r.get("Price (£)", 0) or 0)
            except Exception: price = 0.0
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
        try: return float(v or 0)
        except Exception: return 0.0
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
    if sync: sync_saved_file_to_drive(LEAVER_CLEARANCE_PATH)


def get_next_leaver_clearance_id(records):
    nums = []
    for r in records or []:
        m = re.fullmatch(r"LC-(\d+)", str(r.get("clearance_id", "")).strip().upper())
        if m: nums.append(int(m.group(1)))
    return f"LC-{(max(nums) + 1 if nums else 1):06d}"


def _get_latest_final_settlement(employee_id):
    records = _hrp_get_final_settlement_records(employee_id)
    return records[0] if records else None


def _get_leaver_holiday_balance(employee, leaving_date):
    try:
        if not isinstance(leaving_date, date): leaving_date = pd.to_datetime(leaving_date).date()
    except Exception: leaving_date = date.today()
    calc = _hrp_get_leaving_entitlement(employee, leaving_date)
    used = _hrp_get_approved_holiday_days_to_date(employee.get("emp_id", ""), leaving_date)
    try:
        start = employee.get("start_date")
        if not isinstance(start, date): start = pd.to_datetime(start).date()
    except Exception: start = leaving_date
    closure = _hrp_get_company_closure_holiday_days(employee, start, leaving_date)
    return round(round(calc["net"] - calc["bank_holidays"], 1) - round(used + closure, 1), 1)


def _get_payroll_requests_for_clearance(rec):
    ids = {str(x).strip() for x in str(rec.get("payroll_request_ids", "")).split(",") if str(x).strip()}
    if not ids: return []
    return [r for r in load_records_from_excel(force=True) if str(r.get("id", "")) in ids]


def _sync_clearance_payroll_status(rec, records=None):
    reqs = _get_payroll_requests_for_clearance(rec)
    if not reqs:
        # A Payroll officer may explicitly confirm that final pay is correct and
        # that no Addition/Deduction is required. Do not reset that completed
        # state to Pending on the next page refresh/session.
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

    # Once Payroll has explicitly completed the clearance, do not downgrade the
    # master clearance back to Approved/Rejected on the next page refresh. If a
    # brand-new pending adjustment is later raised, that pending request correctly
    # moves Payroll back to Pending Director Approval.
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
    if not employee or not employee.get("leaving_date"): return None
    emp_id = str(employee.get("emp_id", "")).strip()
    if not emp_id: return None
    records = load_leaver_clearances(force=True)
    rec = next((r for r in records if str(r.get("employee_id", "")).casefold() == emp_id.casefold()), None)
    try:
        leaving_date = employee.get("leaving_date") if isinstance(employee.get("leaving_date"), date) else pd.to_datetime(employee.get("leaving_date")).date()
    except Exception: leaving_date = employee.get("leaving_date")
    holiday_balance = _get_leaver_holiday_balance(employee, leaving_date)
    settlement = _get_latest_final_settlement(emp_id)
    if settlement:
        settlement_status = str(settlement.get("status", "pending")).strip().title()
        settlement_id = str(settlement.get("id", "")); settlement_amount = float(settlement.get("amount", 0) or 0); settlement_type = str(settlement.get("type", ""))
    elif abs(holiday_balance) < 0.05:
        settlement_status = "Not Required"; settlement_id = ""; settlement_amount = 0.0; settlement_type = ""
    else:
        settlement_status = "Required"; settlement_id = ""; settlement_amount = 0.0; settlement_type = ""
    all_items = load_employee_items()
    outstanding = get_employee_outstanding_items(emp_id, all_items)
    outstanding_value = sum(float(r.get("qty_outstanding", 0) or 0) * float(r.get("unit_price", 0) or 0) for r in outstanding)
    checkins = load_item_checkins(force=True)
    emp_checkins = [c for c in checkins if str(c.get("employee_id", "")).casefold() == emp_id.casefold()]
    latest = emp_checkins[-1] if emp_checkins else None
    created_new = rec is None
    if created_new:
        rec = {"clearance_id": get_next_leaver_clearance_id(records), "employee_id": emp_id, "emp_name": employee.get("name", ""), "emp_dept": employee.get("department", ""), "leaving_date": str(leaving_date), "leaving_reason": str(employee.get("leaving_reason", "")), "created_by": created_by or "HR", "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "hr_status": "Complete", "store_status": "Pending", "payroll_status": "Pending", "payroll_request_ids": "", "final_status": "Open", "final_cleared_by": "", "final_cleared_at": "", "notes": ""}
        records.append(rec)
    rec.update({"emp_name": employee.get("name", rec.get("emp_name", "")), "emp_dept": employee.get("department", rec.get("emp_dept", "")), "leaving_date": str(leaving_date), "leaving_reason": str(employee.get("leaving_reason", rec.get("leaving_reason", ""))), "hr_status": "Complete", "holiday_balance": holiday_balance, "holiday_settlement_id": settlement_id, "holiday_settlement_status": settlement_status, "holiday_settlement_amount": settlement_amount, "holiday_settlement_type": settlement_type, "outstanding_item_value": outstanding_value})
    if latest:
        rec["store_checkin_id"] = str(latest.get("id", "")); rec["store_deduction_id"] = str(latest.get("deduction_request_id", "")); rec["store_return_id"] = str(latest.get("return_request_id", ""))
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
    if str(rec.get("store_status", "")).casefold() not in {"not required", "check-in complete", "completed", "cleared"}: return False
    if float(rec.get("outstanding_item_value", 0) or 0) > 0: return False
    # Only Store Deductions require Director approval. Returned items are cleared
    # directly from employee holdings and never create an Addition.
    linked_ids = {x.strip() for x in str(rec.get("store_deduction_id", "")).split(",") if x.strip()}
    if not linked_ids: return True
    try:
        store_reqs = load_store_deductions(force=True)
        linked = [r for r in store_reqs if str(r.get("id", "")) in linked_ids]
        if not linked: return True
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
    """Create a normal generic Addition/Deduction request for Director approval."""
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
    # Do not accidentally create the same leaver adjustment twice from a rerun.
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
        "id": new_id,
        "emp_name": rec.get("emp_name", ""),
        "dept": "Payroll Department",
        "type": trans_type,
        "category": "Leaver Payroll Adjustment",
        "date": str(rec.get("leaving_date", date.today())),
        "amount": amount,
        "manager": "Payroll",
        "desc": f"Leaver Clearance {clearance_id}: {reason}",
        "attachment_name": "None",
        "status": "pending",
        "director_comments": "",
        "decision_date": "",
        "decision_by": "",
        "submitted_by": user_name,
        "pdf_path": "",
        "edited_from_id": "",
        "old_data": "",
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


def _build_final_leaver_clearance_pdf(rec):
    if not PDF_AVAILABLE: return b""
    try:
        pdf = FPDF()
        pdf.add_page(); pdf.set_font("Arial", "B", 16); pdf.cell(0, 10, "EMPLOYEE LEAVER CLEARANCE FORM", ln=1, align="C")
        pdf.set_font("Arial", "", 10); pdf.cell(0, 7, f"Clearance ID: {rec.get('clearance_id')}   Employee: {rec.get('emp_name')} ({rec.get('employee_id')})", ln=1)
        pdf.cell(0, 7, f"Department: {rec.get('emp_dept')}   Leaving Date: {rec.get('leaving_date')}", ln=1); pdf.ln(3)
        for title, lines in [("HR CLEARANCE", [f"Status: {rec.get('hr_status')}", f"Holiday balance: {float(rec.get('holiday_balance',0) or 0):.1f} days", f"Settlement: {rec.get('holiday_settlement_id') or 'Not required'} / {rec.get('holiday_settlement_status')}", f"Settlement amount: £{float(rec.get('holiday_settlement_amount',0) or 0):,.2f}"]), ("STORE CLEARANCE", [f"Status: {rec.get('store_status')}", f"Check-in: {rec.get('store_checkin_id') or 'None'}", f"Outstanding property value: £{float(rec.get('outstanding_item_value',0) or 0):,.2f}", f"Store deduction: {rec.get('store_deduction_id') or 'None'}"]), ("PAYROLL CLEARANCE", [f"Status: {rec.get('payroll_status')}", f"Payroll requests: {rec.get('payroll_request_ids') or 'None'}", f"Approved additions: £{float(rec.get('payroll_total_addition',0) or 0):,.2f}", f"Approved deductions: £{float(rec.get('payroll_total_deduction',0) or 0):,.2f}"]), ("FINAL CLEARANCE", [f"Status: {rec.get('final_status')}", f"Cleared by: {rec.get('final_cleared_by') or '—'}", f"Cleared at: {rec.get('final_cleared_at') or '—'}"])] :
            pdf.set_font("Arial", "B", 12); pdf.cell(0, 8, title, ln=1); pdf.set_font("Arial", "", 10)
            for line in lines: pdf.cell(0, 6, line, ln=1)
            pdf.ln(2)
        return bytes(pdf.output(dest="S"))
    except Exception: return b""


def _save_leaver_clearance_pdf(rec):
    """Create a persistent PDF copy so the completed clearance can be downloaded again later."""
    pdf = _build_final_leaver_clearance_pdf(rec)
    if not pdf:
        return ""
    try:
        os.makedirs(PDF_DIR, exist_ok=True)
        path = os.path.join(PDF_DIR, f"Leaver_Clearance_{rec.get('clearance_id')}.pdf")
        with open(path, "wb") as f:
            f.write(pdf)
        return path
    except Exception:
        return ""


def render_leaver_clearance_hr():
    st.subheader("🔗 Employee Leaver Clearance — HR")
    st.caption("HR starts the company-wide clearance. The clearance record is permanent and remains searchable after the employee has left and been fully cleared.")

    # IMPORTANT: use the persisted HR register, not session state. This ensures a
    # leaver already declared in HR is recreated/loaded after a fresh login or restart.
    employees = _hrp_load_employees()
    user_name = st.session_state.get("user_info", {}).get("full_name", "HR")
    for emp in [e for e in employees if str(e.get("status", "")).casefold() == "left" and e.get("leaving_date")]:
        _ensure_leaver_clearance(emp, user_name)

    records = load_leaver_clearances(force=True)
    if not records:
        st.info("No employee leaver clearances have been created yet. Record an employee as Left in Employee Leaving to start one.")
        return

    # Search the permanent register at any time, including fully-cleared leavers.
    q = st.text_input("🔎 Search clearance", placeholder="Name, Employee ID or Clearance ID (e.g. LC-000125)", key="leaver_clearance_search")
    status_filter = st.selectbox("Status", ["All", "Open", "Fully Cleared"], key="leaver_clearance_status_filter")
    qn = str(q or "").strip().casefold()
    filtered = []
    for r in records:
        hay = " ".join([str(r.get("clearance_id", "")), str(r.get("emp_name", "")), str(r.get("employee_id", "")), str(r.get("emp_dept", ""))]).casefold()
        is_cleared = str(r.get("final_status", "")).casefold() == "cleared"
        if qn and qn not in hay:
            continue
        if status_filter == "Open" and is_cleared:
            continue
        if status_filter == "Fully Cleared" and not is_cleared:
            continue
        filtered.append(r)

    m1,m2,m3,m4=st.columns(4)
    m1.metric("Open",sum(str(r.get("final_status","Open")).casefold()!="cleared" for r in records))
    m2.metric("HR",sum(_clearance_hr_ready(r) for r in records))
    m3.metric("Store",sum(_clearance_store_ready(r) for r in records))
    m4.metric("Fully Cleared",sum(str(r.get("final_status","")).casefold()=="cleared" for r in records))
    st.caption(f"Showing {len(filtered)} of {len(records)} permanent clearance record(s).")

    for rec in reversed(filtered):
        _sync_clearance_payroll_status(rec)
        status=str(rec.get("final_status","Open")).casefold()
        with st.expander(f"{'🟢' if status=='cleared' else '🟡'} {rec.get('clearance_id')} | {rec.get('emp_name')} | {rec.get('employee_id')}", expanded=status!="cleared"):
            a,b,c,d,e=st.columns(5)
            a.metric("HR",rec.get("hr_status")); b.metric("Director",rec.get("holiday_settlement_status")); c.metric("Store",rec.get("store_status")); d.metric("Payroll",rec.get("payroll_status")); e.metric("Final",rec.get("final_status"))
            st.write(f"**Leaving:** {rec.get('leaving_date')} · **Reason:** {rec.get('leaving_reason','')}")
            st.write(f"**Holiday:** {float(rec.get('holiday_balance',0) or 0):.1f} days · Settlement #{rec.get('holiday_settlement_id') or '—'} · £{float(rec.get('holiday_settlement_amount',0) or 0):,.2f}")
            st.write(f"**Store:** Outstanding £{float(rec.get('outstanding_item_value',0) or 0):,.2f} · Check-in #{rec.get('store_checkin_id') or '—'} · Deduction #{rec.get('store_deduction_id') or '—'}")
            st.write(f"**Payroll:** Requests {rec.get('payroll_request_ids') or 'None'} · Approved additions £{float(rec.get('payroll_total_addition',0) or 0):,.2f} · deductions £{float(rec.get('payroll_total_deduction',0) or 0):,.2f}")
            if status == "cleared":
                st.success(f"Fully cleared by {rec.get('final_cleared_by')} on {rec.get('final_cleared_at')}")
            else:
                if not _clearance_director_hr_ready(rec): st.warning("⏳ Director approval is still required for the HR holiday settlement.")
                if not _clearance_store_ready(rec): st.info("📦 Store clearance is still outstanding.")
                if not _clearance_payroll_ready(rec): st.info("🧾 Payroll clearance is still outstanding.")
            pdf = _build_final_leaver_clearance_pdf(rec)
            if status == "cleared":
                _save_leaver_clearance_pdf(rec)
            st.download_button("📄 Download Leaver Clearance Form", data=pdf, file_name=f"Leaver_Clearance_{rec.get('clearance_id')}.pdf", mime="application/pdf", disabled=not bool(pdf), key=f"lc_hr_pdf_{rec.get('clearance_id')}")


def render_leaver_clearance_store(user_name):
    st.subheader("🔗 Leaver Clearance — Store")
    st.caption("Store receives every HR-cleared leaver, including employees with no company property. Store must explicitly clear the employee.")
    employees=_hrp_load_employees(); leavers=[e for e in employees if str(e.get("status","")).casefold()=="left" and e.get("leaving_date")]
    for emp in leavers: _ensure_leaver_clearance(emp,user_name)
    records=load_leaver_clearances(force=True); all_items=load_employee_items(); checkins=load_item_checkins(force=True)
    visible=False
    for rec in reversed(records):
        if str(rec.get("final_status","")).casefold()=="cleared": continue
        # Store sees the leaver immediately after HR declares them. Director approval
        # only controls final Store sign-off; it does not hide the employee from Store.
        visible=True; emp_id=rec.get("employee_id",""); outstanding=get_employee_outstanding_items(emp_id,all_items); value=sum(float(r.get("qty_outstanding",0) or 0)*float(r.get("unit_price",0) or 0) for r in outstanding); emp_checkins=[c for c in checkins if str(c.get("employee_id","")).casefold()==str(emp_id).casefold()]; latest=emp_checkins[-1] if emp_checkins else None
        rec["outstanding_item_value"]=value
        if latest: rec["store_checkin_id"]=str(latest.get("id","")); rec["store_deduction_id"]=str(latest.get("deduction_request_id","")); rec["store_return_id"]=str(latest.get("return_request_id",""))
        if value>0: rec["store_status"]="Pending"
        elif str(rec.get("store_status","")).casefold() not in {"completed","cleared"}: rec["store_status"]="Not Required" if not latest else "Check-in Complete"
        with st.expander(f"🟡 {rec.get('clearance_id')} | {rec.get('emp_name')} | {rec.get('employee_id')}",expanded=True):
            a,b,c=st.columns(3); a.metric("Property", "Outstanding" if outstanding else "None"); b.metric("Value",f"£{value:,.2f}"); c.metric("Status",rec.get("store_status"))
            if outstanding: st.dataframe(pd.DataFrame([{"Item":r.get("item_name",""),"Qty Outstanding":float(r.get("qty_outstanding",0) or 0),"Unit Price":f"£{float(r.get('unit_price',0) or 0):,.2f}","Value":f"£{float(r.get('qty_outstanding',0) or 0)*float(r.get('unit_price',0) or 0):,.2f}"} for r in outstanding]),width="stretch",hide_index=True); st.info("Use Leaver Item Check-in to record returned quantities. Any unreturned items continue through Store's Director approval process.")
            else: st.success("✅ No company property remains outstanding.")
            if not _clearance_director_hr_ready(rec):
                st.info("ℹ️ HR/Director holiday approval is separate from Store. Store can clear its own section now; any Store Deduction still requires Director approval.")
            if not outstanding and str(rec.get("store_status","")).casefold() not in {"completed","cleared"}:
                if st.button("✅ Approve / Clear Store — Nothing Outstanding",key=f"store_clear_{rec.get('clearance_id')}",type="primary",width="stretch"):
                    now=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    rec["store_status"]="Cleared"
                    rec["store_completed_by"]=user_name
                    rec["store_completed_at"]=now
                    save_all_leaver_clearances(records)
                    log_action("LEAVER_STORE_CLEARANCE_COMPLETED",rec.get("clearance_id"),new_data=rec)
                    st.success("Store clearance completed and recorded as Cleared.")
                    st.rerun()
            elif str(rec.get("store_status","")).casefold() in {"completed","cleared"}:
                st.success(f"✅ Store Cleared by {rec.get('store_completed_by')} on {rec.get('store_completed_at')}")
    if not visible: st.info("No leavers are currently waiting for Store clearance.")


def render_leaver_clearance_payroll(user_name):
    st.subheader("🧾 Leaver Clearance — Payroll")
    st.caption(
        "Check the leaver's final pay. Choose Addition or Deduction if an adjustment is needed. "
        "If nothing is required, choose Everything Clear and complete Payroll clearance."
    )

    # Make sure Payroll can see the same master clearance records even if HR/Store
    # has not opened the clearance screen during this session.
    for emp in _hrp_load_employees():
        if str(emp.get("status", "")).casefold() == "left" and emp.get("leaving_date"):
            _ensure_leaver_clearance(emp, user_name)

    records = load_leaver_clearances(force=True)
    if not records:
        st.info("No leaver clearances exist yet.")
        return

    for rec in reversed(records):
        _sync_clearance_payroll_status(rec)

        if str(rec.get("final_status", "")).casefold() == "cleared":
            continue

        with st.expander(
            f"🧾 {rec.get('clearance_id')} | {rec.get('emp_name')} | {rec.get('employee_id')}",
            expanded=True
        ):
            a, b, c = st.columns(3)
            a.metric("Store", rec.get("store_status"))
            b.metric("Payroll", rec.get("payroll_status"))
            c.metric("Final", rec.get("final_status"))

            st.write(
                f"**HR holiday settlement:** {rec.get('holiday_settlement_status')} · "
                f"£{float(rec.get('holiday_settlement_amount', 0) or 0):,.2f}"
            )

            linked = _get_payroll_requests_for_clearance(rec)
            if linked:
                st.dataframe(
                    pd.DataFrame([
                        {
                            "Request": r.get("id"),
                            "Type": r.get("type"),
                            "Amount": f"£{float(r.get('amount', 0) or 0):,.2f}",
                            "Status": r.get("status"),
                            "Description": r.get("desc"),
                        }
                        for r in linked
                    ]),
                    width="stretch",
                    hide_index=True,
                )
            else:
                st.info(
                    "No Payroll adjustment has been raised. If the final pay is already correct, "
                    "select **Everything Clear — No Addition/Deduction** below."
                )

            linked_statuses = {
                str(r.get("status", "pending")).casefold()
                for r in linked
            }
            has_open_adjustment = bool(linked_statuses.intersection({"pending", "approved"}))

            with st.form(f"payroll_lc_form_{rec.get('clearance_id')}"):
                typ = st.selectbox(
                    "Payroll outcome",
                    [
                        "Everything Clear — No Addition/Deduction",
                        "Addition",
                        "Deduction",
                    ],
                    key=f"plc_type_{rec.get('clearance_id')}",
                )

                # Amount and reason are only relevant when an adjustment is being raised.
                if typ in {"Addition", "Deduction"}:
                    amount = st.number_input(
                        f"{typ} Amount (£)",
                        min_value=0.0,
                        step=1.0,
                        format="%.2f",
                        key=f"plc_amount_{rec.get('clearance_id')}",
                    )
                    reason = st.text_area(
                        f"Reason for {typ}",
                        key=f"plc_reason_{rec.get('clearance_id')}",
                        placeholder=(
                            f"Enter the reason for the Payroll {typ.lower()}, "
                            "for example overtime, missing pay, or an authorised deduction."
                        ),
                    )
                    submit_label = f"📤 Raise Payroll {typ} — Send to Director"
                else:
                    amount = 0.0
                    reason = "Final pay checked — no Payroll Addition or Deduction is required."
                    submit_label = "✅ Clear Payroll — No Addition/Deduction Required"

                    if has_open_adjustment:
                        st.warning(
                            "An existing Payroll Addition/Deduction is still Pending or Approved. "
                            "You must complete or resolve that adjustment before Payroll can be "
                            "marked as having no adjustment."
                        )

                submitted = st.form_submit_button(
                    submit_label,
                    type="primary",
                    disabled=(typ == "Everything Clear — No Addition/Deduction" and has_open_adjustment),
                )

            if submitted:
                if typ == "Everything Clear — No Addition/Deduction":
                    # Never delete an existing Payroll request just because the user
                    # selected the no-adjustment option. Rejected requests remain in
                    # the audit trail; pending/approved requests are blocked above.
                    rec["payroll_status"] = "Cleared"
                    rec["payroll_total_addition"] = 0.0
                    rec["payroll_total_deduction"] = 0.0
                    rec["payroll_completed_by"] = user_name
                    rec["payroll_completed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    save_all_leaver_clearances(records)
                    _maybe_finalize_clearance(rec, records)

                    if str(rec.get("final_status", "")).casefold() == "cleared":
                        _save_leaver_clearance_pdf(rec)

                    log_action(
                        "LEAVER_PAYROLL_CLEARANCE_COMPLETED",
                        rec.get("clearance_id"),
                        new_data=rec,
                    )
                    st.success(
                        "✅ Payroll cleared — no Addition or Deduction is required. "
                        "The leaver can now proceed to final clearance."
                    )
                    st.rerun()

                else:
                    # Addition/Deduction becomes a normal pending request for Director approval.
                    try:
                        created = _create_payroll_clearance_request(
                            rec,
                            user_name,
                            typ,
                            amount,
                            str(reason).strip(),
                        )
                        save_all_leaver_clearances(records)
                        st.success(
                            f"✅ Payroll {typ} Request #{created.get('id')} has been sent "
                            "to the Director for approval."
                        )
                        st.info(
                            "Payroll remains pending until the Director approves or rejects "
                            "this request."
                        )
                        st.rerun()
                    except ValueError as exc:
                        st.error(str(exc))
                    except Exception as exc:
                        st.error(f"Unable to create the Payroll {typ} request: {exc}")

            # Refresh the status after any Director-approved/rejected request.
            _sync_clearance_payroll_status(rec)

            if not _clearance_store_ready(rec):
                st.info(
                    "📦 Payroll can be cleared independently. Store completion is only "
                    "required before the overall Final Clearance can become Fully Cleared."
                )

            if linked and _clearance_payroll_ready(rec) and str(
                rec.get("payroll_status", "")
            ).casefold() not in {"cleared"}:
                if st.button(
                    "✅ Complete Payroll Clearance — Approved Adjustment",
                    key=f"plc_complete_{rec.get('clearance_id')}",
                    type="primary",
                ):
                    rec["payroll_completed_by"] = user_name
                    rec["payroll_completed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    rec["payroll_status"] = "Cleared"
                    save_all_leaver_clearances(records)
                    _maybe_finalize_clearance(rec, records)
                    log_action(
                        "LEAVER_PAYROLL_CLEARANCE_COMPLETED",
                        rec.get("clearance_id"),
                        new_data=rec,
                    )
                    st.success("✅ Payroll clearance completed and recorded as Cleared.")
                    st.rerun()
            elif str(rec.get("payroll_status", "")).casefold() == "cleared":
                st.success(
                    f"✅ Payroll Cleared by {rec.get('payroll_completed_by')} "
                    f"on {rec.get('payroll_completed_at')}"
                )

    save_all_leaver_clearances(records)

def render_leaver_clearance_final():
    st.subheader("📋 Final Leaver Clearance Register")
    st.caption("Permanent, searchable record of HR, Director, Store and Payroll clearance. Reports can be downloaded at any time.")
    records=load_leaver_clearances(force=True)
    if not records:
        st.info("No leaver clearance records exist yet.")
        return
    q=st.text_input("🔎 Search permanent clearance register", placeholder="Name, Employee ID or Clearance ID", key="final_leaver_clearance_search")
    status_filter=st.selectbox("Status", ["All", "Open", "Fully Cleared"], key="final_leaver_clearance_status")
    qn=str(q or "").strip().casefold()
    filtered=[]
    for rec in records:
        hay=" ".join([str(rec.get("clearance_id","")),str(rec.get("emp_name","")),str(rec.get("employee_id","")),str(rec.get("emp_dept",""))]).casefold()
        cleared=str(rec.get("final_status","")).casefold()=="cleared"
        if qn and qn not in hay: continue
        if status_filter=="Open" and cleared: continue
        if status_filter=="Fully Cleared" and not cleared: continue
        filtered.append(rec)
    st.caption(f"Showing {len(filtered)} of {len(records)} clearance record(s).")
    for rec in reversed(filtered):
        _sync_clearance_payroll_status(rec)
        if _clearance_all_ready(rec) and str(rec.get("final_status","")).casefold()!="cleared":
            rec["final_status"]="Cleared"; rec["final_cleared_by"]="System"; rec["final_cleared_at"]=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            save_all_leaver_clearances(records); log_action("LEAVER_FINAL_CLEARANCE_COMPLETED",rec.get("clearance_id"),new_data=rec)
        if str(rec.get("final_status","")).casefold()=="cleared":
            _save_leaver_clearance_pdf(rec)
        with st.expander(f"{'🟢' if str(rec.get('final_status')).casefold()=='cleared' else '🟡'} {rec.get('clearance_id')} | {rec.get('emp_name')} | {rec.get('employee_id')}",expanded=False):
            st.write(f"HR: **{rec.get('hr_status')}** · Director: **{rec.get('holiday_settlement_status')}** · Store: **{rec.get('store_status')}** · Payroll: **{rec.get('payroll_status')}** · Final: **{rec.get('final_status')}**")
            pdf=_build_final_leaver_clearance_pdf(rec)
            st.download_button("📄 Download Final Clearance Form",data=pdf,file_name=f"Leaver_Clearance_{rec.get('clearance_id')}.pdf",mime="application/pdf",disabled=not bool(pdf),key=f"final_lc_pdf_{rec.get('clearance_id')}")
    save_all_leaver_clearances(records)


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
                try: return float(r.get(k, 0) or 0)
                except Exception: return 0.0
            try: rid = int(r.get("ID", 0))
            except Exception: rid = 0
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
    if not records: return 1
    return max(int(r.get("id", 0)) for r in records) + 1


def get_employee_outstanding_items(employee_id, records=None):
    """Return every record for an employee that still has a positive outstanding quantity."""
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
            try: rid = int(r.get("ID", 0))
            except Exception: rid = 0
            try: total = float(r.get("Total Deduction (£)", 0) or 0)
            except Exception: total = 0.0
            try: returned = json.loads(r.get("Returned Items JSON", "[]") or "[]")
            except Exception: returned = []
            try: not_returned = json.loads(r.get("Not Returned Items JSON", "[]") or "[]")
            except Exception: not_returned = []
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
    if not records: return 1
    return max(int(r.get("id", 0)) for r in records) + 1

# ============================================================
# 👥 HR LEAVE SETTLEMENT — PDF EXPORT
# ============================================================
def hr_leave_pdf(req, force_regenerate=False, upload_to_drive=True):
    if not PDF_AVAILABLE: return None
    try:
        cached_path = req.get("pdf_path", "")
        if not force_regenerate and cached_path and os.path.exists(cached_path): return cached_path
        pdf = FPDF()
        pdf.add_page()
        regular_font, bold_font = _pdf_font_paths()
        if regular_font and bold_font:
            pdf.add_font("DejaVu", "", regular_font)
            pdf.add_font("DejaVu", "B", bold_font)
            family = "DejaVu"
        else:
            family = "Helvetica"
        def safe(v):
            text = _pdf_text(v)
            if family == "Helvetica": return text.encode("latin-1", "replace").decode("latin-1")
            return text

        if os.path.exists(LOGO_PATH):
            try:
                pdf.image(LOGO_PATH, x=75, y=10, w=60); pdf.ln(28)
            except Exception: pdf.ln(5)
        else: pdf.ln(5)

        pdf.set_font(family, "B", 16)
        pdf.cell(0, 10, safe("HR LEAVE SETTLEMENT - APPROVAL FORM"), ln=True, align="C")
        pdf.ln(2)
        line_y = pdf.get_y()
        pdf.line(10, line_y, 200, line_y); pdf.line(10, line_y + 1.5, 200, line_y + 1.5)
        pdf.ln(8)

        pdf.set_font(family, "B", 11)
        pdf.cell(0, 6, safe("REQUEST DETAILS"), ln=True); pdf.ln(2)
        def field(label, value, label_w=60, value_h=7):
            pdf.set_font(family, "B", 10)
            pdf.cell(label_w, value_h, safe(label), border=0)
            pdf.set_font(family, "", 10)
            pdf.cell(0, value_h, safe(str(value)), border=0, ln=True)
        field("Request ID:", f"HRL-{req.get('id', '')}")
        field("Employee Name:", req.get("emp_name", ""))
        field("Employee Department:", req.get("emp_dept", ""))
        field("Transaction Type:", req.get("type", ""))
        field("Category / Reason:", req.get("category", ""))
        field("Owe / Owed:", req.get("owe_owed", ""))
        field("Request Date:", req.get("date", ""))
        field("Number of Days:", req.get("days", ""))
        field("Amount:", f"£{float(req.get('amount', 0)):.2f}")
        field("Line Manager:", req.get("manager", ""))
        if req.get("submitted_by"): field("Submitted By:", req.get("submitted_by", ""))
        if req.get("submitted_date"): field("Submitted Date:", req.get("submitted_date", ""))
        pdf.ln(6)

        pdf.set_font(family, "B", 11)
        pdf.cell(0, 6, safe("DESCRIPTION / JUSTIFICATION"), ln=True); pdf.ln(2)
        pdf.set_font(family, "", 10)
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(0, 6, safe(req.get("desc", "")))
        pdf.ln(8)

        pdf.set_font(family, "B", 11)
        pdf.cell(0, 6, safe("DIRECTOR APPROVAL"), ln=True); pdf.ln(2)
        status = str(req.get("status", "")).strip().lower()
        pdf.set_font(family, "", 10)
        if status == "approved":
            pdf.cell(60, 7, safe("Decision:"), 0, 0)
            pdf.set_font(family, "B", 10); pdf.set_text_color(0, 128, 0)
            pdf.cell(0, 7, safe("APPROVED"), ln=True)
            pdf.set_text_color(0, 0, 0); pdf.set_font(family, "", 10)
            pdf.cell(60, 7, safe("Approved By:"), 0, 0); pdf.cell(0, 7, safe(req.get("decision_by", "")), ln=True)
            pdf.cell(60, 7, safe("Approval Date / Time:"), 0, 0); pdf.cell(0, 7, safe(req.get("decision_date", "")), ln=True)
            if req.get("director_comments"):
                pdf.cell(60, 7, safe("Director Comments:"), 0, 0)
                pdf.multi_cell(0, 7, safe(req.get("director_comments", "")))
        elif status == "rejected":
            pdf.cell(60, 7, safe("Decision:"), 0, 0)
            pdf.set_font(family, "B", 10); pdf.set_text_color(200, 0, 0)
            pdf.cell(0, 7, safe("REJECTED"), ln=True)
            pdf.set_text_color(0, 0, 0); pdf.set_font(family, "", 10)
            pdf.cell(60, 7, safe("Rejected By:"), 0, 0); pdf.cell(0, 7, safe(req.get("decision_by", "")), ln=True)
            pdf.cell(60, 7, safe("Rejection Date / Time:"), 0, 0); pdf.cell(0, 7, safe(req.get("decision_date", "")), ln=True)
            pdf.cell(60, 7, safe("Reason for Rejection:"), 0, 0)
            pdf.multi_cell(0, 7, safe(req.get("rejection_reason", "")))
        else:
            pdf.cell(60, 7, safe("Decision:"), 0, 0); pdf.cell(0, 7, safe("Pending"), ln=True)

        pdf.ln(14)
        dash_y = pdf.get_y()
        for x in range(10, 200, 4): pdf.line(x, dash_y, x + 2, dash_y)
        if status == "approved" and os.path.exists(APPROVED_STAMP_PATH):
            try: pdf.image(APPROVED_STAMP_PATH, x=70, y=dash_y - 8, w=65)
            except Exception: pass
        elif status == "rejected" and os.path.exists(REJECTED_STAMP_PATH):
            try: pdf.image(REJECTED_STAMP_PATH, x=70, y=dash_y - 8, w=65)
            except Exception: pass
        pdf.ln(18)
        pdf.set_font(family, "", 8)
        pdf.cell(0, 5, safe("Authorised Signature / Director"), ln=True)

        pdf.add_page()
        pdf.set_font(family, "B", 12)
        pdf.cell(0, 8, safe("ATTACHMENTS"), ln=True); pdf.ln(6)
        pdf.set_font(family, "", 9)
        att = str(req.get("attachment_name", "None")).strip()
        display_files = []
        if att and att.lower() not in ("none", "nan", ""):
            for name in att.split(","):
                n = name.strip()
                if n and n.lower() not in ("none", ""): display_files.append(n)
        if display_files:
            for fname in display_files:
                file_path = os.path.join(UPLOAD_DIR, fname)
                if os.path.exists(file_path):
                    if fname.lower().endswith((".png", ".jpg", ".jpeg")):
                        pdf.ln(2)
                        try: pdf.image(file_path, x=10, w=190); pdf.ln(70)
                        except Exception: pdf.cell(0, 5, safe(f"     {fname} (preview unavailable)"), ln=True); pdf.ln(3)
                    else: pdf.cell(0, 5, safe(f"     {fname} (non-image file)"), ln=True); pdf.ln(3)
                else: pdf.cell(0, 5, safe(f"     {fname} (file not found)"), ln=True); pdf.ln(3)
        else: pdf.cell(0, 6, safe("No attachments were included with this request."), ln=True)

        os.makedirs(HR_LEAVE_PDF_DIR, exist_ok=True)
        safe_id = "".join(str(req.get("id", "HRL")).split()) or "HRL"
        safe_name = "_".join(str(req.get("emp_name", "Employee")).split()) or "Employee"
        safe_cat = "_".join(str(req.get("category", "Leave")).split()) or "Leave"
        safe_date = datetime.now().strftime("%Y-%m-%d")
        filename = f"HR_Leave_Settlement_{safe_id}_{safe_name}_{safe_cat}_{safe_date}.pdf"
        path = os.path.join(HR_LEAVE_PDF_DIR, filename)
        pdf.output(path)
        if upload_to_drive: _upload_to_drive_bg(path, os.path.basename(path))
        return path
    except Exception as e:
        st.error(f"HR Leave PDF Error: {e}")
        return None

def display_hr_leave_pdf_button(req, key_prefix="hrl"):
    if not PDF_AVAILABLE:
        st.warning("⚠️ PDF generation is unavailable. Please install fpdf2.")
        return
    status = str(req.get("status", "")).strip().lower()
    if status not in ("approved", "rejected"):
        st.info("📄 PDF download is available once the Director has decided on this request.")
        return
    req_id = str(req.get("id", "unknown"))
    dl_key = f"{key_prefix}_hrl_dl_{req_id}"
    gen_key = f"{key_prefix}_hrl_gen_{req_id}"
    cached_path = req.get("pdf_path", "")
    if cached_path and os.path.exists(cached_path):
        with open(cached_path, "rb") as f:
            st.download_button("⬇️ Download HR Leave Settlement PDF", data=f.read(),
                              file_name=os.path.basename(cached_path), mime="application/pdf",
                              type="primary", key=dl_key)
        return
    if st.button(f"📄 Generate HR Leave PDF for #{req_id}", key=gen_key, type="primary"):
        with st.spinner("Generating PDF..."):
            path = hr_leave_pdf(req, force_regenerate=True, upload_to_drive=True)
        if path and os.path.exists(path):
            records = load_hr_leave()
            for r in records:
                if str(r.get("id")) == req_id: r["pdf_path"] = path
            save_all_hr_leave(records, sync=False)
            st.success("✅ PDF generated. Click below to download.")
            st.rerun()
        else:
            st.error("❌ Could not generate PDF.")

# ============================================================
# 📦 STORE DEPARTMENT DEDUCTION — PDF EXPORT
# ============================================================
def store_deduction_pdf(req, force_regenerate=False, upload_to_drive=True):
    if not PDF_AVAILABLE: return None
    try:
        cached_path = req.get("pdf_path", "")
        if not force_regenerate and cached_path and os.path.exists(cached_path): return cached_path
        pdf = FPDF()
        pdf.add_page()
        regular_font, bold_font = _pdf_font_paths()
        if regular_font and bold_font:
            pdf.add_font("DejaVu", "", regular_font)
            pdf.add_font("DejaVu", "B", bold_font)
            family = "DejaVu"
        else:
            family = "Helvetica"
        def safe(v):
            text = _pdf_text(v)
            if family == "Helvetica": return text.encode("latin-1", "replace").decode("latin-1")
            return text

        if os.path.exists(LOGO_PATH):
            try:
                pdf.image(LOGO_PATH, x=75, y=10, w=60); pdf.ln(28)
            except Exception: pdf.ln(5)
        else: pdf.ln(5)

        is_addition = str(req.get("type", "Deduction")).strip() == "Addition"
        pdf.set_font(family, "B", 16)
        title_text = "STORE DEPARTMENT RETURN (ADDITION) - APPROVAL FORM" if is_addition else "STORE DEPARTMENT DEDUCTION - APPROVAL FORM"
        pdf.cell(0, 10, safe(title_text), ln=True, align="C")
        pdf.ln(2)
        line_y = pdf.get_y()
        pdf.line(10, line_y, 200, line_y); pdf.line(10, line_y + 1.5, 200, line_y + 1.5)
        pdf.ln(8)

        pdf.set_font(family, "B", 11)
        pdf.cell(0, 6, safe("EMPLOYEE DETAILS"), ln=True); pdf.ln(2)
        def field(label, value, label_w=60, value_h=7):
            pdf.set_font(family, "B", 10)
            pdf.cell(label_w, value_h, safe(label), border=0)
            pdf.set_font(family, "", 10)
            pdf.cell(0, value_h, safe(str(value)), border=0, ln=True)
        field("Request ID:", f"STORE-{req.get('id', '')}")
        field("Employee Name:", req.get("emp_name", ""))
        field("Employee Department:", req.get("emp_dept", ""))
        field("Date of Leaving:", req.get("date_leaving", ""))
        field("Line Manager:", req.get("manager", ""))
        field("Date of Submit:", req.get("date_submit", ""))
        if req.get("submitted_by"): field("Submitted By:", req.get("submitted_by", ""))
        pdf.ln(6)

        pdf.set_font(family, "B", 11)
        item_section_title = "ITEMS TO RETURN" if is_addition else "ITEMS TO DEDUCT"
        pdf.cell(0, 6, safe(item_section_title), ln=True); pdf.ln(2)
        pdf.set_font(family, "B", 10)
        pdf.cell(90, 7, safe("Item Name"), border=1)
        pdf.cell(30, 7, safe("Qty"), border=1, align="C")
        pdf.cell(40, 7, safe("Price"), border=1, align="R")
        pdf.cell(0, 7, safe(""), border=1, ln=True)
        pdf.set_font(family, "", 10)
        for item in req.get("items", []):
            pdf.cell(90, 7, safe(item.get("item_name", "")), border=1)
            pdf.cell(30, 7, safe(str(item.get("quantity", 1))), border=1, align="C")
            pdf.cell(40, 7, safe(f"£{float(item.get('price', 0)):.2f}"), border=1, align="R")
            pdf.cell(0, 7, safe(""), border=1, ln=True)
        pdf.set_font(family, "B", 11)
        total_label = "TOTAL EMPLOYEE ADDITION" if is_addition else "TOTAL EMPLOYEE DEDUCTION"
        pdf.cell(120, 8, safe(total_label), border=1, align="R")
        pdf.cell(40, 8, safe(f"£{float(req.get('total_deduction', 0)):.2f}"), border=1, align="R")
        pdf.cell(0, 8, safe(""), border=1, ln=True)
        pdf.ln(6)

        pdf.set_font(family, "B", 11)
        pdf.cell(0, 6, safe("DESCRIPTION / JUSTIFICATION"), ln=True); pdf.ln(2)
        pdf.set_font(family, "", 10)
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(0, 6, safe(req.get("desc", "")))
        pdf.ln(8)

        pdf.set_font(family, "B", 11)
        pdf.cell(0, 6, safe("DIRECTOR APPROVAL"), ln=True); pdf.ln(2)
        status = str(req.get("status", "")).strip().lower()
        pdf.set_font(family, "", 10)
        if status == "approved":
            pdf.cell(60, 7, safe("Decision:"), 0, 0)
            pdf.set_font(family, "B", 10); pdf.set_text_color(0, 128, 0)
            pdf.cell(0, 7, safe("APPROVED"), ln=True)
            pdf.set_text_color(0, 0, 0); pdf.set_font(family, "", 10)
            pdf.cell(60, 7, safe("Approved By:"), 0, 0); pdf.cell(0, 7, safe(req.get("decision_by", "")), ln=True)
            pdf.cell(60, 7, safe("Approval Date / Time:"), 0, 0); pdf.cell(0, 7, safe(req.get("decision_date", "")), ln=True)
            if req.get("director_comments"):
                pdf.cell(60, 7, safe("Director Comments:"), 0, 0)
                pdf.multi_cell(0, 7, safe(req.get("director_comments", "")))
        elif status == "rejected":
            pdf.cell(60, 7, safe("Decision:"), 0, 0)
            pdf.set_font(family, "B", 10); pdf.set_text_color(200, 0, 0)
            pdf.cell(0, 7, safe("REJECTED"), ln=True)
            pdf.set_text_color(0, 0, 0); pdf.set_font(family, "", 10)
            pdf.cell(60, 7, safe("Rejected By:"), 0, 0); pdf.cell(0, 7, safe(req.get("decision_by", "")), ln=True)
            pdf.cell(60, 7, safe("Rejection Date / Time:"), 0, 0); pdf.cell(0, 7, safe(req.get("decision_date", "")), ln=True)
            pdf.cell(60, 7, safe("Reason for Rejection:"), 0, 0)
            pdf.multi_cell(0, 7, safe(req.get("rejection_reason", "")))
        else:
            pdf.cell(60, 7, safe("Decision:"), 0, 0); pdf.cell(0, 7, safe("Pending"), ln=True)

        pdf.ln(14)
        dash_y = pdf.get_y()
        for x in range(10, 200, 4): pdf.line(x, dash_y, x + 2, dash_y)
        if status == "approved" and os.path.exists(APPROVED_STAMP_PATH):
            try: pdf.image(APPROVED_STAMP_PATH, x=70, y=dash_y - 8, w=65)
            except Exception: pass
        elif status == "rejected" and os.path.exists(REJECTED_STAMP_PATH):
            try: pdf.image(REJECTED_STAMP_PATH, x=70, y=dash_y - 8, w=65)
            except Exception: pass
        pdf.ln(18)
        pdf.set_font(family, "", 8)
        pdf.cell(0, 5, safe("Authorised Signature / Director"), ln=True)

        pdf.add_page()
        pdf.set_font(family, "B", 12)
        pdf.cell(0, 8, safe("ATTACHMENTS"), ln=True); pdf.ln(6)
        pdf.set_font(family, "", 9)
        att = str(req.get("attachment_name", "None")).strip()
        display_files = []
        if att and att.lower() not in ("none", "nan", ""):
            for name in att.split(","):
                n = name.strip()
                if n and n.lower() not in ("none", ""): display_files.append(n)
        if display_files:
            for fname in display_files:
                file_path = os.path.join(UPLOAD_DIR, fname)
                if os.path.exists(file_path):
                    if fname.lower().endswith((".png", ".jpg", ".jpeg")):
                        pdf.ln(2)
                        try: pdf.image(file_path, x=10, w=190); pdf.ln(70)
                        except Exception: pdf.cell(0, 5, safe(f"     {fname} (preview unavailable)"), ln=True); pdf.ln(3)
                    else: pdf.cell(0, 5, safe(f"     {fname} (non-image file)"), ln=True); pdf.ln(3)
                else: pdf.cell(0, 5, safe(f"     {fname} (file not found)"), ln=True); pdf.ln(3)
        else: pdf.cell(0, 6, safe("No attachments were included with this request."), ln=True)

        os.makedirs(STORE_DEDUCTION_PDF_DIR, exist_ok=True)
        safe_id = "".join(str(req.get("id", "STORE")).split()) or "STORE"
        safe_name = "_".join(str(req.get("emp_name", "Employee")).split()) or "Employee"
        safe_date = datetime.now().strftime("%Y-%m-%d")
        prefix = "Store_Return" if is_addition else "Store_Deduction"
        filename = f"{prefix}_{safe_id}_{safe_name}_{safe_date}.pdf"
        path = os.path.join(STORE_DEDUCTION_PDF_DIR, filename)
        pdf.output(path)
        if upload_to_drive: _upload_to_drive_bg(path, os.path.basename(path))
        return path
    except Exception as e:
        st.error(f"Store Deduction PDF Error: {e}")
        return None

def display_store_deduction_pdf_button(req, key_prefix="store"):
    if not PDF_AVAILABLE:
        st.warning("⚠️ PDF generation is unavailable. Please install fpdf2.")
        return
    status = str(req.get("status", "")).strip().lower()
    if status not in ("approved", "rejected"):
        st.info("📄 PDF download is available once the Director has decided on this request.")
        return
    req_id = str(req.get("id", "unknown"))
    dl_key = f"{key_prefix}_store_dl_{req_id}"
    gen_key = f"{key_prefix}_store_gen_{req_id}"
    cached_path = req.get("pdf_path", "")
    if cached_path and os.path.exists(cached_path):
        with open(cached_path, "rb") as f:
            st.download_button("⬇️ Download Store PDF", data=f.read(),
                              file_name=os.path.basename(cached_path), mime="application/pdf",
                              type="primary", key=dl_key)
        return
    if st.button(f"📄 Generate Store PDF for #{req_id}", key=gen_key, type="primary"):
        with st.spinner("Generating PDF..."):
            path = store_deduction_pdf(req, force_regenerate=True, upload_to_drive=True)
        if path and os.path.exists(path):
            records = load_store_deductions()
            for r in records:
                if str(r.get("id")) == req_id: r["pdf_path"] = path
            save_all_store_deductions(records, sync=False)
            st.success("✅ PDF generated. Click below to download.")
            st.rerun()
        else:
            st.error("❌ Could not generate PDF.")

# ============================================================
# 👥 HR LEAVE SETTLEMENT — UI
# ============================================================
def render_hr_leave_form(user_name):
    st.subheader("👥 New Request — HR Leave Settlement")
    st.caption("Submit a holiday settlement for Director approval. Employees recorded as Left are available for final holiday settlement.")
    hr_records = load_hr_leave()
    hr_cats = load_hr_categories()
    departments = load_departments()
    employees = _hrp_load_employees()
    left_employees = [
        e for e in employees
        if (str(e.get("status", "")).strip().casefold() == "left" or e.get("leaving_date"))
        and not _hrp_has_active_final_settlement(e.get("emp_id", ""))
    ]

    form_version = st.session_state.get("hr_leave_form_version", 0)
    K_EMP   = f"hr_emp_v{form_version}"
    K_OWE   = f"hr_owe_v{form_version}"
    K_CAT   = f"hr_cat_v{form_version}"
    K_DAYS  = f"hr_days_v{form_version}"
    K_AMT   = f"hr_amt_v{form_version}"
    K_DATE  = f"hr_date_v{form_version}"
    K_DEPT  = f"hr_dept_v{form_version}"
    K_MGR   = f"hr_mgr_v{form_version}"
    K_FILES = f"hr_files_v{form_version}"
    K_DESC  = f"hr_desc_v{form_version}"

    left_options = ["-- Select employee who has left --"] + [
        f"{e['name']} — {e['emp_id']}" for e in left_employees
    ]
    employee_by_label = {
        f"{e['name']} — {e['emp_id']}": e for e in left_employees
    }

    col_left, col_right = st.columns(2)
    with col_left:
        selected_left_label = st.selectbox(
            "👤 Employee Name",
            left_options,
            key=f"hr_left_employee_v{form_version}",
        )
        selected_left = employee_by_label.get(selected_left_label)
        if not selected_left:
            emp_name = st.text_input("Manual Employee Name", key=K_EMP)
        else:
            emp_name = selected_left.get("name", "")
            st.caption(f"Employee ID: {selected_left.get('emp_id', '')} · Status: Left · Leaving date: {selected_left.get('leaving_date') or '-'}")

    with col_right:
        if selected_left:
            emp_dept = selected_left.get("department", "Other")
            st.text_input("🏢 Employee Department", value=emp_dept, disabled=True, key=f"hr_left_dept_v{form_version}")
        else:
            emp_dept = st.selectbox("🏢 Employee Department", departments, key=K_DEPT)
        dt_val = st.date_input("📅 Date", value=(selected_left.get("leaving_date") or date.today()) if selected_left else date.today(), key=K_DATE)

    final_settlement = bool(selected_left)
    if selected_left:
        leaving_date = selected_left.get("leaving_date") or dt_val
        if not isinstance(leaving_date, date):
            try:
                leaving_date = pd.to_datetime(leaving_date).date()
            except Exception:
                leaving_date = dt_val
        calc = _hrp_get_leaving_entitlement(selected_left, leaving_date)
        recorded_used = _hrp_get_approved_holiday_days_to_date(selected_left.get("emp_id"), leaving_date)
        start_date = selected_left.get("start_date")
        if not isinstance(start_date, date):
            try: start_date = pd.to_datetime(start_date).date()
            except Exception: start_date = leaving_date
        closure = _hrp_get_company_closure_holiday_days(selected_left, start_date, leaving_date)
        used_to_leave = round(recorded_used + closure, 1)
        final_balance = round(calc["net"] - used_to_leave - calc["bank_holidays"], 1)
        transaction_type = "Addition" if final_balance > 0 else "Deduction" if final_balance < 0 else "Addition"
        owe_owed = "Company Owes Employee" if final_balance >= 0 else "Employee Owes Company"
        default_category = "Unused Holiday Payout" if final_balance >= 0 else "Overused Holiday Deduction"
        default_days = abs(final_balance)
        st.info(f"Final holiday settlement: **{default_days:.1f} days**. After Director approval, this settlement will close the employee holiday balance at **0.0 days**.")
    else:
        owe_owed = st.selectbox("⚖️ Owe / Owed", DEFAULT_OWE_OWED, key=K_OWE)
        transaction_type = "Addition" if owe_owed == "Company Owes Employee" else "Deduction"
        default_category = None
        default_days = None

    if final_settlement:
        category = default_category if default_category in hr_cats else ("Unused Holiday Payout" if final_balance >= 0 else "Overused Holiday Deduction")
        st.text_input("🏷️ Category / Reason", value=category, disabled=True)
        num_days = st.number_input("🔢 Number of Days", min_value=0.0, value=float(default_days), step=0.5, format="%.2f", disabled=True)
    else:
        category = st.selectbox("🏷️ Category / Reason", hr_cats, key=K_CAT)
        num_days = st.number_input("🔢 Number of Days", min_value=0.0, step=0.5, format="%.2f", key=K_DAYS)

    rate = get_hr_daily_rate(emp_dept, transaction_type)
    col_left2, col_right2 = st.columns(2)
    with col_left2:
        manager = st.text_input("👔 Line Manager", key=K_MGR)
    with col_right2:
        recalc_signature = f"{emp_dept}|{transaction_type}|{float(num_days)}|{final_settlement}"
        if st.session_state.get("_hr_amt_sig") != recalc_signature:
            if rate:
                st.session_state[K_AMT] = round(rate * float(num_days), 2)
            elif K_AMT not in st.session_state:
                st.session_state[K_AMT] = 0.01
            st.session_state["_hr_amt_sig"] = recalc_signature
        amount = st.number_input("💷 Amount (£)", min_value=0.01, step=1.0, format="%.2f", key=K_AMT)
        if rate:
            st.caption(f"ℹ️ Auto-calc: {num_days} × £{rate:.2f} = £{round(rate * float(num_days), 2):.2f} (editable)")
        else:
            st.caption("⚠️ No daily rate configured for this department/type — enter manually.")

    with st.form(f"hr_leave_form_v{form_version}", clear_on_submit=False):
        files = st.file_uploader("📎 Attachments", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key=K_FILES)
        desc_default = ""
        if final_settlement:
            desc_default = f"Final holiday settlement for employee leaving {leaving_date:%d/%m/%Y}. Final balance to settle: {abs(final_balance):.1f} day(s)."
        desc = st.text_area("📝 Description / Justification", value=desc_default if not st.session_state.get(K_DESC) else st.session_state.get(K_DESC), key=K_DESC)
        submitted = st.form_submit_button("📤 Send to Director", type="primary", width="stretch")

    if submitted:
        if not emp_name.strip() or not manager.strip() or not desc.strip():
            st.error("⚠️ Employee Name, Line Manager and Description are required.")
        elif float(num_days) <= 0:
            st.error("⚠️ Number of Days must be greater than 0.")
        elif float(amount) <= 0:
            st.error("⚠️ Amount must be greater than 0.")
        elif final_settlement and abs(float(num_days) - abs(float(final_balance))) > 0.01:
            st.error("⚠️ The final settlement days must match the calculated leaving balance.")
        else:
            if final_settlement and _hrp_has_active_final_settlement(selected_left.get("emp_id", "")):
                st.error(f"⚠️ A pending or approved final settlement already exists for {selected_left.get('name', selected_left.get('emp_id', 'this employee'))}. A duplicate cannot be submitted.")
                st.stop()
            new_id = get_next_hr_leave_id(hr_records)
            attachments = []
            for i, f in enumerate(files or [], 1):
                safe_name = os.path.basename(f.name).replace("/", "_").replace("\\", "_")
                fn = f"HRL_{new_id}_F{i}_{safe_name}"
                fp = os.path.join(UPLOAD_DIR, fn)
                with open(fp, "wb") as out_file: out_file.write(f.getbuffer())
                _upload_to_drive_bg(fp, fn)
                attachments.append(fn)
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            rec = {
                "id": new_id,
                "employee_id": selected_left.get("emp_id", "") if selected_left else "",
                "emp_name": emp_name.strip(), "emp_dept": emp_dept,
                "type": transaction_type, "category": category, "owe_owed": owe_owed,
                "date": str(dt_val), "days": float(num_days), "amount": float(amount),
                "manager": manager.strip(), "desc": desc.strip(),
                "attachment_name": ", ".join(attachments) or "None",
                "status": "pending", "director_comments": "", "rejection_reason": "",
                "decision_date": "", "decision_by": "",
                "submitted_by": user_name, "submitted_date": now, "pdf_path": "",
                "final_holiday_settlement": final_settlement,
            }
            hr_records.append(rec)
            save_all_hr_leave(hr_records)
            log_action("HR_LEAVE_CREATED", new_id, new_data=rec)
            for k in [K_EMP, K_OWE, K_CAT, K_DAYS, K_AMT, K_DATE, K_DEPT, K_MGR, K_FILES, K_DESC, "_hr_amt_sig", f"hr_left_employee_v{form_version}"]:
                st.session_state.pop(k, None)
            st.session_state["hr_leave_form_version"] = form_version + 1
            st.session_state["hr_leave_submission_notice"] = f"HR Leave Settlement #{new_id} sent to Director for approval."
            st.rerun()


def render_hr_leave_my_submissions(user_name):
    logged_user = st.session_state.get("user_info", {}) or {}
    submitted_by_values = {
        str(user_name or "").strip().casefold(),
        str(logged_user.get("full_name", "")).strip().casefold(),
        str(logged_user.get("username", "")).strip().casefold(),
    }
    submitted_by_values.discard("")

    st.subheader("📋 My Submitted HR Leave Requests")
    st.caption("All HR Leave Settlement requests you have submitted — filter by name, department or date.")
    notice = st.session_state.pop("hr_leave_submission_notice", "")
    if notice:
        st.success(f"✅ {notice}")
    st.divider()

    hr_records = load_hr_leave()
    mine = [
        r for r in hr_records
        if str(r.get("submitted_by", "")).strip().casefold() in submitted_by_values
    ]

    c1, c2, c3 = st.columns(3)
    with c1:
        search_name = st.text_input(
            "🔎 Search by Employee Name",
            key="hr_sub_search_name",
            placeholder="Type employee name...",
        )
    with c2:
        dept_options = sorted({
            str(r.get("emp_dept", "")).strip()
            for r in mine if str(r.get("emp_dept", "")).strip()
        })
        search_dept = st.selectbox(
            "🏢 Filter by Department",
            ["All Departments"] + dept_options,
            key="hr_sub_search_dept",
        )
    with c3:
        use_date = st.checkbox("📅 Filter by specific date", key="hr_sub_use_date")
        search_date = None
        if use_date:
            search_date = st.date_input("Pick a date", value=date.today(),
                                        key="hr_sub_search_date")

    filtered = list(mine)
    if search_name.strip():
        q = search_name.lower().strip()
        filtered = [r for r in filtered if q in str(r.get("emp_name", "")).lower()]
    if search_dept and search_dept != "All Departments":
        filtered = [r for r in filtered if str(r.get("emp_dept", "")) == search_dept]
    if search_date is not None:
        filtered = [r for r in filtered if str(r.get("date", "")) == str(search_date)]

    m1, m2, m3, m4 = st.columns(4)
    with m1: st.metric("📋 Total Found", len(filtered))
    with m2: st.metric("🟡 Pending", len([r for r in filtered if r["status"] == "pending"]))
    with m3: st.metric("🟢 Approved", len([r for r in filtered if r["status"] == "approved"]))
    with m4: st.metric("🔴 Rejected", len([r for r in filtered if r["status"] == "rejected"]))
    st.divider()

    if not filtered:
        st.info("📋 No HR Leave requests match your search.")
        return

    for r in reversed(filtered):
        status = r["status"]
        icon = "🟡" if status == "pending" else ("🟢" if status == "approved" else "🔴")
        with st.expander(f"{icon} #{r['id']} | {r['emp_name']} | {r['emp_dept']} | £{r['amount']:.2f} | {status.upper()}"):
            st.write(f"🔄 **{r['type']}** | ⚖️ {r['owe_owed']} | 🏷️ {r['category']}")
            st.write(f"🔢 **{r['days']} day(s)** | 💷 **£{r['amount']:.2f}** | 📅 {r['date']}")
            st.write(f"👔 **Line Manager:** {r['manager']}")
            st.info(f"📝 {r['desc']}")
            display_attachments(r)
            if r.get("director_comments"): st.info(f"💬 Director: {r['director_comments']}")
            if r.get("rejection_reason"): st.error(f"❌ Rejection Reason: {r['rejection_reason']}")
            if status in ("approved", "rejected"):
                st.divider()
                display_hr_leave_pdf_button(r, key_prefix=f"hr_sub_{user_name.replace(' ','_')}")


def render_hr_leave_director_portal(director_name):
    st.subheader("👥 HR Leave Settlement — Director Approval")
    st.info("Review HR leave settlement requests. The Director can move any request between Pending, Approved and Rejected. Rejection requires a reason.")
    st.divider()
    records = [r for r in load_hr_leave(force=False) if r.get("final_holiday_settlement", False)]
    pending = [r for r in records if str(r.get("status", "")).strip().lower() == "pending"]
    approved = [r for r in records if str(r.get("status", "")).strip().lower() == "approved"]
    rejected = [r for r in records if str(r.get("status", "")).strip().lower() == "rejected"]
    t1, t2, t3 = st.tabs([f"⏳ Pending ({len(pending)})", f"✅ Approved ({len(approved)})", f"❌ Rejected ({len(rejected)})"])

    def show_details(r):
        st.write(f"👤 **Employee:** {r.get('emp_name', '')} | 🏢 **Department:** {r.get('emp_dept', '')}")
        st.write(f"🔄 **Transaction Type:** {r.get('type', '')} | ⚖️ **Owe / Owed:** {r.get('owe_owed', '')}")
        st.write(f"🏷️ **Category / Reason:** {r.get('category', '')} | 📅 **Date:** {r.get('date', '')}")
        st.write(f"🔢 **Days:** {r.get('days', 0)} | 💷 **Amount:** £{float(r.get('amount', 0) or 0):.2f}")
        st.write(f"👔 **Line Manager:** {r.get('manager', '')}")
        st.write(f"📝 **Submitted by:** {r.get('submitted_by', '')} on {r.get('submitted_date', '')}")
        if r.get("final_holiday_settlement"):
            st.success("🏁 Final Holiday Settlement")
        st.info(f"📝 **Description:**\n{r.get('desc', '')}")
        if r.get("director_comments"):
            st.warning(f"💬 **Director Comments:** {r.get('director_comments')}")
        if str(r.get("status", "")).strip().lower() == "rejected":
            reason = r.get("rejection_reason") or r.get("director_comments") or "No rejection reason recorded."
            st.error(f"❌ **Rejection reason:** {reason}")
        display_attachments(r)

    def change_status(rid, new_status, comments="", rejection_reason=""):
        old_status = ""
        for x in records:
            if str(x.get("id")) == str(rid):
                old_status = str(x.get("status", "pending")).strip().lower()
                x["status"] = new_status
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                if new_status == "approved":
                    x["director_comments"] = comments.strip() if comments.strip() else x.get("director_comments", "")
                    x["decision_by"] = director_name
                    x["decision_date"] = now
                    x["rejection_reason"] = ""
                elif new_status == "rejected":
                    x["director_comments"] = comments.strip() if comments.strip() else x.get("director_comments", "")
                    x["rejection_reason"] = rejection_reason.strip()
                    x["decision_by"] = director_name
                    x["decision_date"] = now
                elif new_status == "pending":
                    x["decision_by"] = ""
                    x["decision_date"] = ""
                    if comments.strip():
                        x["director_comments"] = (str(x.get("director_comments", "")) + f"\n[{now[:16]}] Changed to Pending by {director_name}: {comments.strip()}").strip()
                break
        save_all_hr_leave(records)
        log_action("HR_LEAVE_STATUS_CHANGED", rid,
                   old_data={"status": old_status},
                   new_data={"status": new_status, "comments": comments, "rejection_reason": rejection_reason},
                   decision_by=director_name,
                   decision_date=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        st.success(f"✅ HR Leave #{rid} changed to **{new_status.title()}**.")
        st.rerun()

    def render_record(r, current_status, key_prefix):
        rid = r.get("id")
        with st.expander(f"{'🟡' if current_status == 'pending' else '🟢' if current_status == 'approved' else '🔴'} #{rid} | {r.get('emp_name', '')} | 🔄 {r.get('type', 'Not specified')} | £{float(r.get('amount', 0) or 0):.2f} | {current_status.upper()}"):
            show_details(r)
            st.markdown("### 🔄 Director Status Control")
            comments = st.text_area("Director Comments (optional)", key=f"{key_prefix}_comm_{rid}")
            c1, c2, c3 = st.columns(3)
            with c1:
                if current_status != "pending" and st.button("⏳ Move to Pending", key=f"{key_prefix}_pending_{rid}", width="stretch"):
                    change_status(rid, "pending", comments)
                elif current_status == "pending":
                    st.button("⏳ Pending", key=f"{key_prefix}_pending_disabled_{rid}", disabled=True, width="stretch")
            with c2:
                if current_status != "approved" and st.button("✅ Set Approved", key=f"{key_prefix}_approved_{rid}", type="primary", width="stretch"):
                    change_status(rid, "approved", comments)
                elif current_status == "approved":
                    st.button("✅ Approved", key=f"{key_prefix}_approved_disabled_{rid}", disabled=True, width="stretch")
            with c3:
                if current_status != "rejected" and st.button("❌ Set Rejected", key=f"{key_prefix}_rejected_{rid}", width="stretch"):
                    st.session_state[f"hr_dir_reject_{rid}"] = True
                elif current_status == "rejected":
                    st.button("❌ Rejected", key=f"{key_prefix}_rejected_disabled_{rid}", disabled=True, width="stretch")
            if st.session_state.get(f"hr_dir_reject_{rid}"):
                reason = st.text_area("Rejection Reason (required)", key=f"{key_prefix}_reason_{rid}")
                rc1, rc2 = st.columns(2)
                with rc1:
                    if st.button("Confirm Rejection", key=f"{key_prefix}_confirm_rej_{rid}", type="primary", width="stretch"):
                        if not reason.strip():
                            st.error("❌ Rejection reason is required.")
                        else:
                            change_status(rid, "rejected", comments, reason)
                with rc2:
                    if st.button("Cancel", key=f"{key_prefix}_cancel_rej_{rid}", width="stretch"):
                        st.session_state[f"hr_dir_reject_{rid}"] = False
                        st.rerun()
            st.divider()
            display_hr_leave_pdf_button(r, key_prefix=f"{key_prefix}_pdf_{rid}")

    with t1:
        if not pending:
            st.success("✅ No pending HR Leave requests.")
        for r in reversed(pending):
            render_record(r, "pending", "hrdir_pending")
    with t2:
        if not approved:
            st.info("✅ No approved HR Leave requests.")
        for r in reversed(approved):
            render_record(r, "approved", "hrdir_approved")
    with t3:
        if not rejected:
            st.info("❌ No rejected HR Leave requests.")
        for r in reversed(rejected):
            render_record(r, "rejected", "hrdir_rejected")

def render_hr_leave_super_admin():
    st.subheader("🛡️ HR Leave Settlement — Super Admin (View Only)")
    st.info("✅ View all HR Leave requests. **Approval → Director only.**")
    st.divider()
    records = load_hr_leave()
    search = st.text_input("🔎 Search HR Leave requests", placeholder="Search by ID, employee, department, category, amount, status...", key="hr_sa_search")
    if search.strip():
        q = search.lower().strip()
        records = [r for r in records if q in " ".join(str(v) for v in r.values()).lower()]
    pending  = [r for r in records if r["status"] == "pending"]
    approved = [r for r in records if r["status"] == "approved"]
    rejected = [r for r in records if r["status"] == "rejected"]
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("🟡 Pending", len(pending))
    with c2: st.metric("🟢 Approved", len(approved))
    with c3: st.metric("🔴 Rejected", len(rejected))
    st.divider()
    t1, t2, t3 = st.tabs([f"⏳ Pending ({len(pending)})", f"✅ Approved ({len(approved)})", f"❌ Rejected ({len(rejected)})"])
    def show(r):
        st.write(f"👤 **{r['emp_name']}** | 🏢 {r['emp_dept']} | 📅 {r['date']}")
        st.write(f"🔄 {r['type']} | ⚖️ {r['owe_owed']} | 🏷️ {r['category']}")
        st.write(f"🔢 {r['days']} day(s) | 💷 £{r['amount']:.2f} | 👔 {r['manager']}")
        st.info(f"📝 {r['desc']}")
        st.caption(f"Submitted by {r['submitted_by']} on {r['submitted_date']}")
        display_attachments(r)
    with t1:
        if not pending: st.success("✅ No pending HR Leave requests.")
        for r in reversed(pending):
            with st.expander(f"🟡 #{r['id']} | {r['emp_name']} | {r['emp_dept']} | £{r['amount']:.2f}"):
                show(r)
    with t2:
        if not approved: st.info("✅ No approved HR Leave requests.")
        for r in reversed(approved):
            with st.expander(f"🟢 #{r['id']} | {r['emp_name']} | £{r['amount']:.2f} | ✅ {r['decision_by']}"):
                show(r)
                if r.get("director_comments"): st.info(f"💬 {r['director_comments']}")
                st.divider()
                display_hr_leave_pdf_button(r, key_prefix="hr_sa_app")
    with t3:
        if not rejected: st.info("❌ No rejected HR Leave requests.")
        for r in reversed(rejected):
            with st.expander(f"🔴 #{r['id']} | {r['emp_name']} | £{r['amount']:.2f}"):
                show(r); st.error(f"❌ Reason: {r['rejection_reason']}")
                st.divider()
                display_hr_leave_pdf_button(r, key_prefix="hr_sa_rej")


def _super_admin_transaction_control():
    """Super Admin master search/edit/delete for all transaction workbooks."""
    st.subheader("🛡️ Super Admin — Transaction & Employee Control")
    st.info("Search any transaction across the main request systems, edit or delete it, and manage the complete employee master record. Changes are saved to the live Excel files and Google Drive backup automatically.")
    st.divider()

    def _save_portal_leave(records):
        _hr_portal_init()
        st.session_state["hrp_leave_records"] = list(records)
        _hrp_save_leave_records()

    def source_records():
        return [
            ("Addition & Deduction", "AD", load_records_from_excel(force=True), save_all_records, "id"),
            ("HR Leave Settlement", "HR", load_hr_leave(force=True), save_all_hr_leave, "id"),
            ("Store Transactions", "ST", load_store_deductions(force=True), save_all_store_deductions, "id"),
            ("Work Orders", "WO", load_work_orders(force=True), save_all_work_orders, "id"),
            ("Inspector Bonus", "IB", load_inspector_bonus(force=True), save_all_inspector_bonus, "id"),
            ("HR Portal Leave Records", "LV", (_hr_portal_init() or st.session_state.get("hrp_leave_records", [])), _save_portal_leave, "leave_id"),
            ("Employee Items", "EI", load_employee_items(force=True), save_all_employee_items, "id"),
            ("Item Check-ins", "IC", load_item_checkins(force=True), save_all_item_checkins, "id"),
        ]

    search_tab, employee_tab = st.tabs(["🔎 Transaction Search / Edit / Delete", "👤 Employee Master Edit"])

    with search_tab:
        sources = source_records()
        all_rows=[]
        for label, code, records, saver, id_key in sources:
            for r in records:
                rid = str(r.get(id_key, ""))
                name = str(r.get("emp_name") or r.get("inspector_name") or r.get("item_name") or "")
                dept = str(r.get("emp_dept") or r.get("dept") or "")
                status = str(r.get("status", ""))
                amount = r.get("amount", r.get("total_deduction", r.get("bonus_amount", r.get("total_value", 0))))
                try: amount=float(amount or 0)
                except Exception: amount=0.0
                typ = str(r.get("type") or ("Addition" if label == "Store Transactions" and str(r.get("type","")) == "Addition" else "") or "")
                if label == "Inspector Bonus": typ = "Bonus"
                actor = str(r.get("submitted_by") or r.get("issued_by") or r.get("checked_in_by") or r.get("decision_by") or "").strip()
                all_rows.append({"source":label,"code":code,"id":rid,"name":name,"dept":dept,"status":status,"amount":amount,"type":typ,"actor":actor,"record":r,"saver":saver,"id_key":id_key})

        q=st.text_input("🔎 Search all transactions", placeholder="ID, employee, department, user, status, addition, deduction, category, amount, description...", key="sa_master_tx_search")
        source_filter=st.selectbox("📂 Transaction type", ["All"]+[x[0] for x in sources], key="sa_master_tx_source")
        status_filter=st.selectbox("📌 Status", ["All","pending","approved","rejected","pending_manager","pending_director"], key="sa_master_tx_status")
        user_options=sorted({x["actor"] for x in all_rows if x.get("actor")})
        user_filter=st.selectbox("👤 Done / Submitted By", ["All Users"] + user_options, key="sa_master_tx_user")
        filtered=all_rows
        if source_filter != "All": filtered=[x for x in filtered if x["source"]==source_filter]
        if status_filter != "All": filtered=[x for x in filtered if x["status"].lower()==status_filter]
        if user_filter != "All Users": filtered=[x for x in filtered if x.get("actor")==user_filter]
        if q.strip():
            qq=q.strip().casefold(); filtered=[x for x in filtered if qq in json.dumps(x["record"], default=str, ensure_ascii=False).casefold()]
        st.metric("Transactions found", len(filtered))
        st.divider()
        if not filtered:
            st.info("No transactions match the search.")
        else:
            for idx, item in enumerate(reversed(filtered)):
                r=item["record"]; status=item["status"].lower(); icon="🟡" if "pending" in status else ("🟢" if status=="approved" else "🔴" if status=="rejected" else "⚪")
                title=f"{icon} {item['source']} | ID #{item['id']} | {item['name'] or '—'} | {item['type'] or 'Transaction'} | £{item['amount']:.2f} | {status.upper()}"
                with st.expander(title):
                    st.caption(f"Source: {item['source']} • Department: {item['dept'] or '—'} • Done by: {item.get('actor') or '—'}")
                    st.write("**All editable fields** — edit values below, then save. The transaction ID is kept fixed to prevent broken references.")
                    editable={}
                    with st.form(f"sa_edit_tx_{item['code']}_{item['id']}_{idx}"):
                        cols=st.columns(2)
                        for j,(k,v) in enumerate(r.items()):
                            if k == item.get("id_key"): continue
                            label=str(k).replace("_"," ").title()
                            target=cols[j%2]
                            if isinstance(v,(dict,list)):
                                editable[k]=target.text_area(label, value=json.dumps(v, ensure_ascii=False, default=str, indent=2), height=90)
                            else:
                                txt="" if v is None else str(v)
                                editable[k]=target.text_area(label, value=txt, height=68) if len(txt)>100 else target.text_input(label, value=txt)
                        save_col, del_col=st.columns(2)
                        with save_col:
                            save_btn=st.form_submit_button("💾 Save Transaction Changes", type="primary", width="stretch")
                        with del_col:
                            delete_btn=st.form_submit_button("🗑️ Delete Transaction", width="stretch")
                    if save_btn or delete_btn:
                        fresh=next((x[2] for x in source_records() if x[0]==item["source"]), [])
                        target_rec=next((x for x in fresh if str(x.get(item["id_key"]))==str(item["id"])), None)
                        if target_rec is None:
                            st.error("Transaction no longer exists. Refresh the page.")
                        elif delete_btn:
                            fresh=[x for x in fresh if str(x.get(item["id_key"]))!=str(item["id"])]
                            item["saver"](fresh)
                            log_action("SUPER_ADMIN_TRANSACTION_DELETED", item["id"], old_data={"source":item["source"],"record":target_rec}, decision_by="Super Admin", decision_date=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                            st.success(f"✅ {item['source']} #{item['id']} deleted.")
                            st.rerun()
                        else:
                            for k, raw in editable.items():
                                old=target_rec.get(k)
                                if isinstance(old,(dict,list)):
                                    try: target_rec[k]=json.loads(raw) if str(raw).strip() else ([] if isinstance(old,list) else {})
                                    except Exception: st.error(f"Invalid JSON in {k}."); break
                                elif isinstance(old,bool): target_rec[k]=str(raw).strip().casefold() in {"true","1","yes","y"}
                                elif isinstance(old,int) and not isinstance(old,bool):
                                    try: target_rec[k]=int(float(str(raw).strip() or 0))
                                    except Exception: target_rec[k]=raw
                                elif isinstance(old,float):
                                    try: target_rec[k]=float(str(raw).strip() or 0)
                                    except Exception: target_rec[k]=raw
                                else: target_rec[k]=raw
                            item["saver"](fresh)
                            log_action("SUPER_ADMIN_TRANSACTION_EDITED", item["id"], old_data={"source":item["source"]}, new_data={"source":item["source"],"record":target_rec}, decision_by="Super Admin", decision_date=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                            st.success(f"✅ {item['source']} #{item['id']} updated.")
                            st.rerun()

    with employee_tab:
        st.subheader("👤 Employee Master — Full Edit")
        employees=_hrp_load_employees()
        if not employees:
            st.info("No employees found.")
        else:
            labels=[f"{e.get('emp_id')} — {e.get('name')}" for e in employees]
            selected=st.selectbox("Select employee", labels, key="sa_employee_master_select")
            emp=employees[labels.index(selected)]
            old_id=str(emp.get("emp_id",""))
            with st.form("sa_employee_master_form", clear_on_submit=False):
                c1,c2=st.columns(2)
                new_id=c1.text_input("ACE-ID / Employee ID", value=old_id)
                new_name=c2.text_input("Full Name", value=emp.get("name",""))
                new_start=c1.date_input("Start Date", value=emp.get("start_date") or date.today())
                new_position=c2.text_input("Position / Job Title", value=emp.get("job_title",""))
                depts=load_departments(); dept_idx=depts.index(emp.get("department")) if emp.get("department") in depts else 0
                new_dept=c1.selectbox("Department", depts, index=dept_idx)
                agreements=["Permanent","Temporary","Fixed Term","Apprenticeship","Other"]; ag_idx=agreements.index(emp.get("agreement_type")) if emp.get("agreement_type") in agreements else 0
                new_agreement=c2.selectbox("Agreement Type", agreements, index=ag_idx)
                statuses=["Active","Inactive","Left"]; st_idx=statuses.index(emp.get("status")) if emp.get("status") in statuses else 0
                new_status=c1.selectbox("Status", statuses, index=st_idx)
                if str(emp.get("status", "")).strip().casefold() == "left":
                    st.info("↩️ This employee is currently marked as Left. Select Active above and save the employee record to reactivate them. Historical leave and settlement records are retained.")
                patterns=["Regular hours","Irregular hours / Part-Year"]; wp_idx=patterns.index(emp.get("working_pattern")) if emp.get("working_pattern") in patterns else 0
                new_pattern=c2.selectbox("Working Pattern", patterns, index=wp_idx)
                new_days=c1.number_input("Days Worked Per Week", min_value=0.0, max_value=7.0, step=0.5, value=float(emp.get("days_per_week",5)))
                ov=emp.get("entitlement_override")
                new_override=c2.number_input("Holiday Entitlement Override (blank = automatic)", min_value=0.0, max_value=365.0, step=0.5, value=float(ov) if ov is not None else 0.0)
                new_note=st.text_area("Entitlement Adjustment Note", value=emp.get("adjustment_note",""))
                new_leave=st.date_input("Leaving Date (use 1970-01-01 for none)", value=emp.get("leaving_date") or date(1970,1,1))
                new_reason=st.text_input("Leaving Reason", value=emp.get("leaving_reason",""))
                save_emp=st.form_submit_button("💾 Save Complete Employee Record", type="primary", width="stretch")
            if str(emp.get("status", "")).strip().casefold() == "left":
                if st.button("↩️ Reactivate Left Employee", key=f"sa_reactivate_left_outside_form_{old_id}", type="secondary", width="stretch"):
                    old_data = dict(emp)
                    emp["status"] = "Active"
                    emp["leaving_date"] = None
                    emp["leaving_reason"] = ""
                    pd.DataFrame([{
                        "Employee ID":x.get("emp_id",""),"Full Name":x.get("name",""),"Start Date":x.get("start_date",""),"Position / Job Title":x.get("job_title",""),"Department":x.get("department",""),"Agreement Type":x.get("agreement_type",""),"Status":x.get("status","Active"),"Working Pattern":x.get("working_pattern","Regular hours"),"Days Worked Per Week":x.get("days_per_week",5),"Holiday Entitlement Override":x.get("entitlement_override","") if x.get("entitlement_override") is not None else "","Entitlement Adjustment Note":x.get("adjustment_note",""),"Leaving Date":x.get("leaving_date","") or "","Leaving Reason":x.get("leaving_reason","")
                    } for x in employees], columns=HR_EMPLOYEE_COLUMNS).to_excel(HR_EMPLOYEES_PATH,index=False,engine="openpyxl")
                    sync_saved_file_to_drive(HR_EMPLOYEES_PATH)
                    st.session_state.hrp_employees = employees
                    log_action("SUPER_ADMIN_EMPLOYEE_REACTIVATED", old_id, old_data=old_data, new_data={"employee_id":old_id,"status":"Active"}, decision_by="Super Admin", decision_date=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    st.success(f"✅ {emp.get('name', old_id)} has been reactivated. They will no longer appear in Leavers, and historical records have been retained.")
                    st.rerun()

            if save_emp:
                new_id=str(new_id).strip().upper()
                valid_id, validated_id = _hrp_validate_employee_id(new_id, exclude_id=old_id)
                if not valid_id:
                    st.error(validated_id)
                else:
                    new_id = validated_id
                if valid_id:
                    emp["emp_id"]=new_id; emp["name"]=new_name.strip(); emp["start_date"]=new_start; emp["job_title"]=new_position.strip(); emp["department"]=new_dept; emp["agreement_type"]=new_agreement; emp["status"]=new_status; emp["working_pattern"]=new_pattern; emp["days_per_week"]=float(new_days); emp["entitlement_override"]=None if float(new_override)==0 else float(new_override); emp["adjustment_note"]=new_note.strip(); emp["leaving_date"]=None if new_status != "Left" or new_leave==date(1970,1,1) else new_leave; emp["leaving_reason"]="" if new_status != "Left" else new_reason.strip()
                    if new_id != old_id:
                        users=load_users()
                        changed=False
                        for u in users.values():
                            if str(u.get("employee_id",""))==old_id: u["employee_id"]=new_id; changed=True
                        if changed: save_users(users)
                        hr_leave=load_hr_leave(force=True)
                        for r in hr_leave:
                            if str(r.get("employee_id",""))==old_id: r["employee_id"]=new_id
                        save_all_hr_leave(hr_leave)
                        _hr_portal_init()
                        portal=st.session_state.get("hrp_leave_records",[])
                        for r in portal:
                            if str(r.get("employee_id",""))==old_id: r["employee_id"]=new_id
                        st.session_state["hrp_leave_records"]=portal
                        _hrp_save_leave_records()
                    pd.DataFrame([{
                        "Employee ID":x.get("emp_id",""),"Full Name":x.get("name",""),"Start Date":x.get("start_date",""),"Position / Job Title":x.get("job_title",""),"Department":x.get("department",""),"Agreement Type":x.get("agreement_type",""),"Status":x.get("status","Active"),"Working Pattern":x.get("working_pattern","Regular hours"),"Days Worked Per Week":x.get("days_per_week",5),"Holiday Entitlement Override":x.get("entitlement_override","") if x.get("entitlement_override") is not None else "","Entitlement Adjustment Note":x.get("adjustment_note",""),"Leaving Date":x.get("leaving_date","") or "","Leaving Reason":x.get("leaving_reason","")
                    } for x in employees], columns=HR_EMPLOYEE_COLUMNS).to_excel(HR_EMPLOYEES_PATH,index=False,engine="openpyxl")
                    sync_saved_file_to_drive(HR_EMPLOYEES_PATH)
                    st.session_state.hrp_employees=employees
                    log_action("SUPER_ADMIN_EMPLOYEE_EDITED", new_id, old_data={"employee_id":old_id}, new_data=emp, decision_by="Super Admin", decision_date=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    st.success(f"✅ Employee {new_id} updated successfully. Linked records were updated where required.")
                    st.rerun()

def render_hr_leave_payroll_portal():
    st.subheader("👥 HR Leave Settlement — Payroll (View Only)")
    st.info("View all approved HR Leave settlements for payroll processing.")
    st.divider()
    approved = [r for r in load_hr_leave() if r["status"] == "approved"]
    st.metric("✅ Approved HR Leave Settlements", len(approved))
    st.divider()
    if not approved: st.info("No approved HR Leave settlements yet.")
    for r in reversed(approved):
        with st.expander(f"🟢 #{r['id']} | {r['emp_name']} | {r['emp_dept']} | £{r['amount']:.2f} | {r['type']}"):
            st.write(f"🏷️ {r['category']} | ⚖️ {r['owe_owed']} | 🔢 {r['days']} day(s)")
            st.write(f"💷 **£{r['amount']:.2f}** | 👔 {r['manager']} | 📅 {r['date']}")
            st.write(f"✅ Approved by {r['decision_by']} on {r['decision_date']}")
            if r.get("director_comments"): st.info(f"💬 {r['director_comments']}")
            display_attachments(r)
            st.divider()
            display_hr_leave_pdf_button(r, key_prefix="hr_payroll")

def render_hr_leave_settings():
    st.markdown("### 🏷️ HR Leave Categories")
    st.caption("These appear in the HR Leave Settlement form dropdown.")
    cats = load_hr_categories()
    with st.form("hr_cat_add_form", clear_on_submit=True):
        new_cat = st.text_input("➕ Add New Category", placeholder="e.g. Maternity Leave Settlement")
        if st.form_submit_button("✅ Add Category"):
            if new_cat.strip() and new_cat.strip() not in cats:
                cats.append(new_cat.strip()); save_hr_categories(cats)
                log_action("HR_CATEGORY_ADDED", new_data={"name": new_cat.strip()})
                st.success(f"✅ Added: {new_cat}"); st.rerun()
            elif new_cat.strip() in cats:
                st.warning("⚠️ Category already exists.")
    for i, cat in enumerate(cats):
        c1, c2 = st.columns([5, 1])
        c1.markdown(f"• **{cat}**")
        with c2:
            if len(cats) > 1 and st.button("🗑️", key=f"hr_cat_del_{i}"):
                cats.pop(i); save_hr_categories(cats)
                log_action("HR_CATEGORY_DELETED", old_data={"name": cat})
                st.success(f"Deleted: {cat}"); st.rerun()

    st.divider()
    st.markdown("### 💷 Holiday Daily Rates (per Department + Type)")
    st.caption("The HR form uses these rates to auto-calculate the amount when Number of Days is entered.")
    rates = load_hr_daily_rates()
    departments = load_departments()
    with st.form("hr_rate_add_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1: new_dept = st.selectbox("🏢 Department", departments, key="hr_rate_dept")
        with c2: new_type = st.selectbox("🔄 Transaction Type", ["Addition", "Deduction"], key="hr_rate_type")
        with c3: new_rate = st.number_input("💷 Daily Rate (£)", min_value=0.01, step=1.0, format="%.2f", value=100.00, key="hr_rate_amt")
        if st.form_submit_button("➕ Add / Update Rate", type="primary"):
            existing = next((r for r in rates if r["dept"] == new_dept and r["type"] == new_type), None)
            if existing:
                existing["rate"] = float(new_rate); existing["active"] = True
                log_action("HR_RATE_UPDATED", new_data={"dept": new_dept, "type": new_type, "rate": float(new_rate)})
            else:
                rates.append({"dept": new_dept, "type": new_type, "rate": float(new_rate), "active": True})
                log_action("HR_RATE_ADDED", new_data={"dept": new_dept, "type": new_type, "rate": float(new_rate)})
            save_hr_daily_rates(rates)
            st.success(f"✅ Saved rate for {new_dept} / {new_type}."); st.rerun()
    st.divider()
    if not rates:
        st.info("📋 No daily rates configured yet.")
    else:
        for i, r in enumerate(rates):
            c1, c2, c3, c4 = st.columns([3, 2, 2, 1])
            c1.markdown(f"🏢 **{r['dept']}**")
            c2.markdown(f"🔄 {r['type']}")
            c3.markdown(f"💷 **£{r['rate']:.2f}** {'✅' if r['active'] else '⛔'}")
            with c4:
                if st.button("🗑️", key=f"hr_rate_del_{i}"):
                    rates.pop(i); save_hr_daily_rates(rates)
                    log_action("HR_RATE_DELETED", old_data={"dept": r["dept"], "type": r["type"]})
                    st.success("Deleted."); st.rerun()

# ============================================================
# 📦 STORE DEPARTMENT DEDUCTION & RETURN — UI
# ============================================================
def render_store_deduction_form(user_name, user_dept):
    st.subheader("📦 New Request — Store Department Deduction")
    st.caption("Submit unreturned company items for Director approval and final settlement deduction.")
    all_deductions = load_store_deductions(force=True)
    employees = [e for e in _hrp_load_employees()
                 if str(e.get("status", "Active")).strip().casefold() == "active"]
    if not employees:
        st.warning("No Active employees are registered by HR. Add the employee in HR Employee Directory first.")
        return
    employees = sorted(employees, key=lambda x: (str(x.get("name", "")).casefold(), str(x.get("emp_id", "")).casefold()))
    employee_labels = [
        f"{e.get('name','')} — {e.get('emp_id','')} · {e.get('department','')}"
        for e in employees
    ]
    employee_map = {label: e for label, e in zip(employee_labels, employees)}

    selected_employee_label = st.selectbox(
        "👤 Employee (HR Registered Employees Only)",
        employee_labels,
        key="store_emp_selector",
    )
    selected_employee = employee_map[selected_employee_label]
    emp_name = selected_employee.get("name", "")
    emp_dept = str(selected_employee.get("department", "")).strip()

    # Use the employee ID in the widget key so Streamlit cannot retain the
    # previous employee's department/date when the employee selection changes.
    employee_key = re.sub(r"[^A-Za-z0-9_]+", "_", str(selected_employee.get("emp_id", emp_name)))
    leaving_default = selected_employee.get("leaving_date") or date.today()
    if not isinstance(leaving_default, date):
        try: leaving_default = pd.to_datetime(leaving_default).date()
        except Exception: leaving_default = date.today()
    col1, col2 = st.columns(2)
    with col1:
        st.text_input(
            "🏢 Employee Department",
            value=emp_dept,
            disabled=True,
            key=f"store_emp_dept_display_{employee_key}",
        )
        date_submit = st.date_input("📅 Date of Submit", value=date.today(), key="store_date_submit")
    with col2:
        date_leaving = st.date_input("📅 Date of Leaving", value=leaving_default, key="store_date_leaving")
        manager = st.text_input("👔 Line Manager", key="store_manager")
        st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 📦 Items to Deduct")
    if "store_form_items" not in st.session_state:
        st.session_state.store_form_items = [{"item_name": "", "quantity": 1, "price": 0.0}]
    store_items_master = load_store_items()
    item_options = [""] + [it["name"] for it in store_items_master]
    for i, item_row in enumerate(st.session_state.store_form_items):
        c1, c2, c3, c4 = st.columns([3, 1.5, 2, 0.5])
        with c1:
            selected_item = st.selectbox("Item Name", options=item_options, key=f"store_item_select_{i}", index=item_options.index(item_row.get("item_name", "")) if item_row.get("item_name") in item_options else 0)
            if selected_item != item_row.get("item_name", ""):
                item_row["item_name"] = selected_item
                match = next((it for it in store_items_master if it["name"] == selected_item), None)
                if match:
                    item_row["price"] = match["price"]
                    if f"store_item_price_{i}" in st.session_state:
                        st.session_state[f"store_item_price_{i}"] = match["price"]
                st.rerun()
        with c2:
            item_row["quantity"] = st.number_input("Quantity", min_value=1, step=1, value=int(item_row.get("quantity", 1)), key=f"store_item_qty_{i}")
        with c3:
            item_row["price"] = st.number_input("Price (£)", min_value=0.0, step=1.0, format="%.2f", value=float(item_row["price"]), key=f"store_item_price_{i}")
        with c4:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑️", key=f"store_item_remove_{i}"):
                st.session_state.store_form_items.pop(i)
                st.rerun()
    if st.button("➕ Add Another Item", key="store_add_item_btn"):
        st.session_state.store_form_items.append({"item_name": "", "quantity": 1, "price": 0.0})
        st.rerun()
    total_deduction = sum(item.get("quantity", 1) * item.get("price", 0.0) for item in st.session_state.store_form_items)
    st.markdown(f"<h4 style='text-align: right; color: #ef4444;'>Total Employee Deduction: £{total_deduction:.2f}</h4>", unsafe_allow_html=True)
    st.divider()
    with st.form("store_deduction_form", clear_on_submit=False):
        files = st.file_uploader("📎 Attachments", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="store_files")
        desc = st.text_area("📝 Description / Justification", key="store_desc")
        submitted = st.form_submit_button("📤 Send to Director", type="primary", width="stretch")
    if submitted:
        valid_items = [it for it in st.session_state.store_form_items if it.get("item_name") and it.get("quantity", 0) > 0 and it.get("price", 0) > 0]
        if not emp_name.strip() or not manager.strip() or not desc.strip():
            st.error("⚠️ Employee Name, Line Manager and Description are required.")
        elif not valid_items:
            st.error("⚠️ Please select at least one item to deduct.")
        else:
            new_id = get_next_store_deduction_id(all_deductions)
            attachments = []
            for i, f in enumerate(files or [], 1):
                safe_name = os.path.basename(f.name).replace("/", "_").replace("\\", "_")
                fn = f"STORE_{new_id}_F{i}_{safe_name}"
                fp = os.path.join(UPLOAD_DIR, fn)
                with open(fp, "wb") as out_file: out_file.write(f.getbuffer())
                _upload_to_drive_bg(fp, fn)
                attachments.append(fn)
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            rec = {"id": new_id, "emp_name": emp_name.strip(), "date_leaving": str(date_leaving), "emp_dept": emp_dept, "manager": manager.strip(), "date_submit": str(date_submit), "type": "Deduction", "items": valid_items, "total_deduction": total_deduction, "desc": desc.strip(), "attachment_name": ", ".join(attachments) or "None", "status": "pending", "director_comments": "", "rejection_reason": "", "decision_date": "", "decision_by": "", "submitted_by": user_name, "submitted_date": now, "pdf_path": ""}
            all_deductions.append(rec)
            save_all_store_deductions(all_deductions)
            log_action("STORE_DEDUCTION_CREATED", new_id, new_data=rec)
            st.session_state.store_form_items = [{"item_name": "", "quantity": 1, "price": 0.0}]
            st.success(f"✅ Store Department Deduction #{new_id} sent to Director for approval.")
            st.rerun()

def render_store_return_form(user_name, user_dept):
    st.subheader("📦 New Request — Store Department Return (Addition)")
    st.caption("Process returned items to reverse a previous deduction, or add items not previously deducted.")
    all_transactions = load_store_deductions()
    approved_deds = [
        d for d in all_transactions
        if d["status"] == "approved" and d.get("type", "Deduction") == "Deduction"
    ]

    # Store Returns must use the HR employee register as the source of truth.
    # Do not allow free-text/manual employee names.
    employees = [
        e for e in _hrp_load_employees()
        if str(e.get("name", "")).strip()
    ]
    employees = sorted(
        employees,
        key=lambda x: (str(x.get("name", "")).casefold(), str(x.get("emp_id", "")).casefold()),
    )
    if not employees:
        st.warning("No employees are registered by HR. Add the employee in HR Employee Directory first.")
        return

    employee_labels = [
        f"{e.get('name','')} — {e.get('emp_id','')} · {e.get('department','')}"
        for e in employees
    ]
    employee_map = {label: e for label, e in zip(employee_labels, employees)}
    selected_employee_label = st.selectbox(
        "👤 Employee (HR Registered Employees Only)",
        employee_labels,
        key="store_ret_emp",
    )
    selected_employee = employee_map[selected_employee_label]
    selected_emp = str(selected_employee.get("name", "")).strip()
    emp_dept = str(selected_employee.get("department", "")).strip()

    employee_key = re.sub(
        r"[^A-Za-z0-9_]+",
        "_",
        str(selected_employee.get("emp_id", selected_emp)),
    )
    st.text_input(
        "🏢 Employee Department",
        value=emp_dept,
        disabled=True,
        key=f"store_ret_emp_dept_display_{employee_key}",
    )
    st.markdown("### 📦 Items to Return")
    st.caption("If the item was deducted in this software, click the button below to load them. You can then adjust the quantities for partial returns or add new items manually.")
    if "store_return_items" not in st.session_state:
        st.session_state.store_return_items = [{"item_name": "", "quantity": 1, "price": 0.0}]
    store_items_master = load_store_items()
    item_options = [""] + [it["name"] for it in store_items_master]
    if selected_emp and selected_emp != "-- Manual Entry --":
        if st.button("🔄 Fetch Approved Deductions for this Employee", key="store_fetch_deds_btn"):
            emp_deds = [d for d in approved_deds if d["emp_name"] == selected_emp]
            agg_items = {}
            for d in emp_deds:
                for item in d.get("items", []):
                    name = item.get("item_name")
                    qty = int(item.get("quantity", 1))
                    price = float(item.get("price", 0))
                    if name in agg_items:
                        agg_items[name]["quantity"] += qty
                    else:
                        agg_items[name] = {"item_name": name, "quantity": qty, "price": price}
            if agg_items:
                st.session_state.store_return_items = list(agg_items.values())
                st.success(f"Loaded {len(agg_items)} item(s) from previous deductions.")
                st.rerun()
            else:
                st.warning("No items found in previous deductions for this employee.")
    for i, item_row in enumerate(st.session_state.store_return_items):
        c1, c2, c3, c4 = st.columns([3, 1.5, 2, 0.5])
        with c1:
            selected_item = st.selectbox("Item Name", options=item_options, key=f"store_ret_select_{i}", index=item_options.index(item_row.get("item_name", "")) if item_row.get("item_name") in item_options else 0)
            if selected_item != item_row.get("item_name", ""):
                item_row["item_name"] = selected_item
                match = next((it for it in store_items_master if it["name"] == selected_item), None)
                if match:
                    item_row["price"] = match["price"]
                    if f"store_ret_price_{i}" in st.session_state:
                        st.session_state[f"store_ret_price_{i}"] = match["price"]
                st.rerun()
        with c2:
            item_row["quantity"] = st.number_input("Quantity", min_value=1, step=1, value=int(item_row.get("quantity", 1)), key=f"store_ret_qty_{i}")
        with c3:
            item_row["price"] = st.number_input("Price (£)", min_value=0.0, step=1.0, format="%.2f", value=float(item_row["price"]), key=f"store_ret_price_{i}")
        with c4:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑️", key=f"store_ret_remove_{i}"):
                st.session_state.store_return_items.pop(i)
                st.rerun()
    if st.button("➕ Add Another Item", key="store_add_ret_item_btn"):
        st.session_state.store_return_items.append({"item_name": "", "quantity": 1, "price": 0.0})
        st.rerun()
    total_addition = sum(item.get("quantity", 1) * item.get("price", 0.0) for item in st.session_state.store_return_items)
    st.markdown(f"<h4 style='text-align: right; color: #10b981;'>Total Employee Addition: £{total_addition:.2f}</h4>", unsafe_allow_html=True)
    st.divider()
    with st.form("store_return_form", clear_on_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            date_return = st.date_input("📅 Date of Return", value=date.today(), key="store_ret_date")
            manager = st.text_input("👔 Line Manager", key="store_ret_manager")
        with c2:
            date_submit = st.date_input("📅 Date of Submit", value=date.today(), key="store_ret_submit")
            st.markdown("<br>", unsafe_allow_html=True)
        files = st.file_uploader("📎 Attachments", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="store_ret_files")
        desc = st.text_area("📝 Description / Justification", key="store_ret_desc")
        submitted = st.form_submit_button("📤 Send to Director for Approval", type="primary", width="stretch")
        if submitted:
            valid_items = [it for it in st.session_state.store_return_items if it.get("item_name") and it.get("quantity", 0) > 0 and it.get("price", 0) > 0]
            if not selected_emp or not manager.strip() or not desc.strip():
                st.error("⚠️ Employee Name, Line Manager and Description are required.")
            elif not valid_items:
                st.error("⚠️ Please select at least one item to return.")
            else:
                new_id = get_next_store_deduction_id(all_transactions)
                attachments = []
                for i, f in enumerate(files or [], 1):
                    safe_name = os.path.basename(f.name).replace("/", "_").replace("\\", "_")
                    fn = f"STORE_RET_{new_id}_F{i}_{safe_name}"
                    fp = os.path.join(UPLOAD_DIR, fn)
                    with open(fp, "wb") as out_file: out_file.write(f.getbuffer())
                    _upload_to_drive_bg(fp, fn)
                    attachments.append(fn)
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                rec = {"id": new_id, "emp_name": selected_emp, "date_leaving": str(date_return), "emp_dept": emp_dept, "manager": manager.strip(), "date_submit": str(date_submit), "type": "Addition", "items": valid_items, "total_deduction": total_addition, "desc": desc.strip(), "attachment_name": ", ".join(attachments) or "None", "status": "pending", "director_comments": "", "rejection_reason": "", "decision_date": "", "decision_by": "", "submitted_by": user_name, "submitted_date": now, "pdf_path": ""}
                all_transactions.append(rec)
                save_all_store_deductions(all_transactions)
                log_action("STORE_DEDUCTION_CREATED", new_id, new_data=rec)
                st.session_state.store_return_items = [{"item_name": "", "quantity": 1, "price": 0.0}]
                st.success(f"✅ Store Return #{new_id} sent to Director for approval.")
                st.rerun()

def render_store_my_submissions(user_name):
    st.subheader("📋 My Submitted Store Requests")
    st.caption("All Store Department transactions (Deductions and Returns) you have submitted.")
    st.divider()
    records = load_store_deductions(force=True)
    mine = [r for r in records if r.get("submitted_by") == user_name]
    c1, c2, c3 = st.columns(3)
    with c1:
        search_name = st.text_input("🔎 Search by Employee Name", key="store_sub_search_name", placeholder="Type employee name...")
    with c2:
        dept_options = sorted({str(r.get("emp_dept", "")).strip() for r in mine if str(r.get("emp_dept", "")).strip()})
        search_dept = st.selectbox("🏢 Filter by Department", ["All Departments"] + dept_options, key="store_sub_search_dept")
    with c3:
        use_date = st.checkbox("📅 Filter by specific date", key="store_sub_use_date")
        search_date = None
        if use_date:
            search_date = st.date_input("Pick a date", value=date.today(), key="store_sub_search_date")
    filtered = list(mine)
    if search_name.strip():
        q = search_name.lower().strip()
        filtered = [r for r in filtered if q in str(r.get("emp_name", "")).lower()]
    if search_dept and search_dept != "All Departments":
        filtered = [r for r in filtered if str(r.get("emp_dept", "")) == search_dept]
    if search_date is not None:
        filtered = [r for r in filtered if str(r.get("date_submit", "")) == str(search_date)]
    m1, m2, m3, m4 = st.columns(4)
    with m1: st.metric("📋 Total Found", len(filtered))
    with m2: st.metric("🟡 Pending", len([r for r in filtered if r["status"] == "pending"]))
    with m3: st.metric("🟢 Approved", len([r for r in filtered if r["status"] == "approved"]))
    with m4: st.metric("🔴 Rejected", len([r for r in filtered if r["status"] == "rejected"]))
    st.divider()
    if not filtered:
        st.info("📋 No Store requests match your search.")
        return
    for r in reversed(filtered):
        status = r["status"]
        icon = "🟡" if status == "pending" else ("🟢" if status == "approved" else "🔴")
        type_label = "Deduction" if r.get("type", "Deduction") == "Deduction" else "Return (Addition)"
        type_color = "#ef4444" if r.get("type", "Deduction") == "Deduction" else "#10b981"
        with st.expander(f"{icon} #{r['id']} | {r['emp_name']} | {type_label} | £{r['total_deduction']:.2f} | {status.upper()}"):
            st.markdown(f"**Type:** <span style='color:{type_color}; font-weight:bold;'>{type_label}</span>", unsafe_allow_html=True)
            st.write(f"📅 **Leaving:** {r['date_leaving']} | **Submit:** {r['date_submit']}")
            st.write(f"👔 **Line Manager:** {r['manager']}")
            st.markdown("**Items:**")
            for item in r.get("items", []):
                st.write(f"- {item.get('item_name')} (Qty: {item.get('quantity', 1)}) : £{float(item.get('price', 0)):.2f}")
            st.write(f"**Total: £{r['total_deduction']:.2f}**")
            st.info(f"📝 {r['desc']}")
            display_attachments(r)
            if r.get("director_comments"): st.info(f"💬 Director: {r['director_comments']}")
            if r.get("rejection_reason"): st.error(f"❌ Rejection Reason: {r['rejection_reason']}")
            if status in ("approved", "rejected"):
                st.divider()
                display_store_deduction_pdf_button(r, key_prefix=f"store_sub_{user_name.replace(' ','_')}")

def render_store_director_portal(director_name, type_filter="Deduction"):
    title = "📦 Store Department Deduction" if type_filter == "Deduction" else "📦 Store Department Return (Addition)"
    st.subheader(f"{title} — Director Approval")
    st.info(f"Review {type_filter.lower()} requests. The Director can move any request between Pending, Approved and Rejected. Rejection requires a reason.")
    st.divider()
    records = load_store_deductions()
    filtered_records = [r for r in records if r.get("type", "Deduction") == type_filter]
    pending = [r for r in filtered_records if str(r.get("status", "")).strip().lower() == "pending"]
    approved = [r for r in filtered_records if str(r.get("status", "")).strip().lower() == "approved"]
    rejected = [r for r in filtered_records if str(r.get("status", "")).strip().lower() == "rejected"]
    t1, t2, t3 = st.tabs([f"⏳ Pending ({len(pending)})", f"✅ Approved ({len(approved)})", f"❌ Rejected ({len(rejected)})"])

    def show_details(r):
        st.write(f"👤 **Employee:** {r.get('emp_name', '')} | 🏢 **Department:** {r.get('emp_dept', '')}")
        st.write(f"📅 **Date of Leaving:** {r.get('date_leaving', '')} | **Date of Submit:** {r.get('date_submit', '')}")
        st.write(f"👔 **Line Manager:** {r.get('manager', '')}")
        st.markdown("**Items:**")
        for item in r.get("items", []):
            st.write(f"- {item.get('item_name')} (Qty: {item.get('quantity', 1)}) : £{float(item.get('price', 0)):.2f}")
        st.markdown(f"**Total: £{float(r.get('total_deduction', 0) or 0):.2f}**")
        st.write(f"📝 **Submitted by:** {r.get('submitted_by', '')} on {r.get('submitted_date', '')}")
        st.info(f"📝 **Description:**\n{r.get('desc', '')}")
        display_attachments(r)
        if r.get("director_comments"):
            st.warning(f"💬 **Director Comments:** {r.get('director_comments')}")
        if str(r.get("status", "")).strip().lower() == "rejected":
            reason = r.get("rejection_reason") or r.get("director_comments") or "No rejection reason recorded."
            st.error(f"❌ **Rejection reason:** {reason}")

    def change_status(rid, new_status, comments="", rejection_reason=""):
        old_status = ""
        for x in records:
            if str(x.get("id")) == str(rid):
                old_status = str(x.get("status", "pending")).strip().lower()
                x["status"] = new_status
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                if new_status == "approved":
                    x["director_comments"] = comments.strip() if comments.strip() else x.get("director_comments", "")
                    x["decision_by"] = director_name
                    x["decision_date"] = now
                    x["rejection_reason"] = ""
                elif new_status == "rejected":
                    x["director_comments"] = comments.strip() if comments.strip() else x.get("director_comments", "")
                    x["rejection_reason"] = rejection_reason.strip()
                    x["decision_by"] = director_name
                    x["decision_date"] = now
                elif new_status == "pending":
                    x["decision_by"] = ""
                    x["decision_date"] = ""
                    if comments.strip():
                        x["director_comments"] = (str(x.get("director_comments", "")) + f"\n[{now[:16]}] Changed to Pending by {director_name}: {comments.strip()}").strip()
                break
        save_all_store_deductions(records)
        log_action("STORE_DEDUCTION_STATUS_CHANGED", rid, old_data={"status": old_status}, new_data={"status": new_status}, decision_by=director_name)
        st.success(f"✅ Request #{rid} changed to **{new_status.title()}**.")
        st.rerun()

    def render_record(r, current_status, prefix):
        rid = r.get("id")
        with st.expander(f"{'🟡' if current_status == 'pending' else '🟢' if current_status == 'approved' else '🔴'} #{rid} | {r.get('emp_name', '')} | 🔄 {r.get('type', type_filter)} | £{float(r.get('total_deduction', 0) or 0):.2f} | {current_status.upper()}"):
            show_details(r)
            st.markdown("### 🔄 Director Status Control")
            comments = st.text_area("Director Comments (optional)", key=f"{prefix}_comm_{rid}")
            c1, c2, c3 = st.columns(3)
            with c1:
                if current_status != "pending" and st.button("⏳ Move to Pending", key=f"{prefix}_pending_{rid}", width="stretch"):
                    change_status(rid, "pending", comments)
            with c2:
                if current_status != "approved" and st.button("✅ Set Approved", key=f"{prefix}_approved_{rid}", type="primary", width="stretch"):
                    change_status(rid, "approved", comments)
            with c3:
                if current_status != "rejected" and st.button("❌ Set Rejected", key=f"{prefix}_rejected_{rid}", width="stretch"):
                    st.session_state[f"store_dir_reject_{type_filter}_{rid}"] = True
            if st.session_state.get(f"store_dir_reject_{type_filter}_{rid}"):
                reason = st.text_area("Rejection Reason (required)", key=f"{prefix}_reason_{rid}")
                rc1, rc2 = st.columns(2)
                with rc1:
                    if st.button("Confirm Rejection", key=f"{prefix}_confirm_rej_{rid}", type="primary", width="stretch"):
                        if not reason.strip():
                            st.error("❌ Rejection reason is required.")
                        else:
                            change_status(rid, "rejected", comments, reason)
                with rc2:
                    if st.button("Cancel", key=f"{prefix}_cancel_rej_{rid}", width="stretch"):
                        st.session_state[f"store_dir_reject_{type_filter}_{rid}"] = False
                        st.rerun()
            st.divider()
            display_store_deduction_pdf_button(r, key_prefix=f"store_dir_{type_filter}_{rid}")

    with t1:
        if not pending: st.success(f"✅ No pending {type_filter.lower()} requests.")
        for r in reversed(pending): render_record(r, "pending", f"store_dir_p_{type_filter}")
    with t2:
        if not approved: st.info(f"✅ No approved {type_filter.lower()} requests.")
        for r in reversed(approved): render_record(r, "approved", f"store_dir_a_{type_filter}")
    with t3:
        if not rejected: st.info(f"❌ No rejected {type_filter.lower()} requests.")
        for r in reversed(rejected): render_record(r, "rejected", f"store_dir_r_{type_filter}")

def render_store_super_admin(type_filter="Deduction"):
    title = "Store Department Deduction" if type_filter == "Deduction" else "Store Department Return (Addition)"
    st.subheader(f"🛡️ {title} — Super Admin (View Only)")
    st.info(f"✅ View all {type_filter.lower()} requests. **Approval → Director only.**")
    st.divider()
    records = load_store_deductions()
    filtered_records = [r for r in records if r.get("type", "Deduction") == type_filter]
    search = st.text_input("🔎 Search requests", placeholder="Search by ID, employee, department, amount, status...", key=f"store_sa_search_{type_filter}")
    if search.strip():
        q = search.lower().strip()
        filtered_records = [r for r in filtered_records if q in " ".join(str(v) for v in r.values()).lower()]
    pending  = [r for r in filtered_records if r["status"] == "pending"]
    approved = [r for r in filtered_records if r["status"] == "approved"]
    rejected = [r for r in filtered_records if r["status"] == "rejected"]
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("🟡 Pending", len(pending))
    with c2: st.metric("🟢 Approved", len(approved))
    with c3: st.metric("🔴 Rejected", len(rejected))
    st.divider()
    t1, t2, t3 = st.tabs([f"⏳ Pending ({len(pending)})", f"✅ Approved ({len(approved)})", f"❌ Rejected ({len(rejected)})"])
    def show(r):
        st.write(f"👤 **{r['emp_name']}** | 🏢 {r['emp_dept']} | 📅 Leaving: {r['date_leaving']} | Submit: {r['date_submit']}")
        st.write(f"👔 **Manager:** {r['manager']}")
        st.markdown("**Items:**")
        for item in r.get("items", []):
            st.write(f"- {item.get('item_name')} (Qty: {item.get('quantity', 1)}) : £{float(item.get('price', 0)):.2f}")
        st.markdown(f"**Total: £{r['total_deduction']:.2f}**")
        st.info(f"📝 {r['desc']}")
        st.caption(f"Submitted by {r['submitted_by']} on {r['submitted_date']}")
        display_attachments(r)
    with t1:
        if not pending: st.success(f"✅ No pending {type_filter.lower()} requests.")
        for r in reversed(pending):
            with st.expander(f"🟡 #{r['id']} | {r['emp_name']} | {r['emp_dept']} | £{r['total_deduction']:.2f}"):
                show(r)
    with t2:
        if not approved: st.info(f"✅ No approved {type_filter.lower()} requests.")
        for r in reversed(approved):
            with st.expander(f"🟢 #{r['id']} | {r['emp_name']} | £{r['total_deduction']:.2f} | ✅ {r['decision_by']}"):
                show(r)
                if r.get("director_comments"): st.info(f"💬 {r['director_comments']}")
                st.divider()
                display_store_deduction_pdf_button(r, key_prefix=f"store_sa_app_{type_filter}")
    with t3:
        if not rejected: st.info(f"❌ No rejected {type_filter.lower()} requests.")
        for r in reversed(rejected):
            with st.expander(f"🔴 #{r['id']} | {r['emp_name']} | £{r['total_deduction']:.2f}"):
                show(r); st.error(f"❌ Reason: {r['rejection_reason']}")
                st.divider()
                display_store_deduction_pdf_button(r, key_prefix=f"store_sa_rej_{type_filter}")

def render_store_payroll_portal(type_filter="Deduction"):
    title = "Store Department Deduction" if type_filter == "Deduction" else "Store Department Return (Addition)"
    st.subheader(f"📦 {title} — Payroll (View Only)")
    st.info(f"View all approved {type_filter.lower()} requests for payroll processing.")
    st.divider()
    records = load_store_deductions()
    filtered_records = [r for r in records if r.get("type", "Deduction") == type_filter]
    approved = [r for r in filtered_records if r["status"] == "approved"]
    st.metric(f"✅ Approved {type_filter}s", len(approved))
    st.divider()
    if not approved: st.info(f"No approved {type_filter.lower()} requests yet.")
    for r in reversed(approved):
        with st.expander(f"🟢 #{r['id']} | {r['emp_name']} | {r['emp_dept']} | £{r['total_deduction']:.2f}"):
            st.write(f"📅 Leaving: {r['date_leaving']} | Submit: {r['date_submit']} | 👔 {r['manager']}")
            st.markdown("**Items:**")
            for item in r.get("items", []):
                st.write(f"- {item.get('item_name')} (Qty: {item.get('quantity', 1)}) : £{float(item.get('price', 0)):.2f}")
            st.markdown(f"**Total: £{r['total_deduction']:.2f}**")
            st.write(f"✅ Approved by {r['decision_by']} on {r['decision_date']}")
            if r.get("director_comments"): st.info(f"💬 {r['director_comments']}")
            display_attachments(r)
            st.divider()
            display_store_deduction_pdf_button(r, key_prefix=f"store_payroll_{type_filter}")

def render_store_items_settings():
    st.markdown("### 📦 Store Items & Prices")
    st.caption("These items and prices will be available in the Store Department Deduction form dropdown.")
    items = load_store_items()
    if "editing_store_item_idx" not in st.session_state:
        st.session_state.editing_store_item_idx = None
    with st.form("add_store_item_form", clear_on_submit=True):
        new_item = st.text_input("➕ Add New Item", placeholder="e.g. Safety Helmet")
        new_price = st.number_input("💷 Price (£)", min_value=0.01, step=1.0, format="%.2f", value=50.00)
        if st.form_submit_button("✅ Add Item"):
            if new_item.strip() and not any(it["name"].lower() == new_item.strip().lower() for it in items):
                items.append({"name": new_item.strip(), "price": float(new_price), "active": True})
                save_store_items(items)
                log_action("STORE_ITEM_ADDED", new_data={"name": new_item.strip(), "price": float(new_price)})
                st.success(f"✅ Added: {new_item}")
                st.rerun()
            elif any(it["name"].lower() == new_item.strip().lower() for it in items):
                st.warning("⚠️ Item already exists.")
    st.divider()
    if not items:
        st.info("📋 No store items configured yet.")
    else:
        for i, item in enumerate(items):
            if st.session_state.editing_store_item_idx == i:
                st.markdown(f"#### ✏️ Editing Item")
                with st.form(f"edit_store_item_form_{i}", clear_on_submit=False):
                    ec1, ec2 = st.columns(2)
                    with ec1:
                        edit_name = st.text_input("Item Name", value=item['name'], key=f"edit_store_name_{i}")
                    with ec2:
                        edit_price = st.number_input("Price (£)", min_value=0.01, step=1.0, format="%.2f", value=float(item['price']), key=f"edit_store_price_{i}")
                    btn_col1, btn_col2, _ = st.columns([1, 1, 3])
                    with btn_col1:
                        save_btn = st.form_submit_button("💾 Save", type="primary", width="stretch")
                    with btn_col2:
                        cancel_btn = st.form_submit_button("❌ Cancel", width="stretch")
                    if save_btn:
                        if edit_name.strip():
                            old_data = {"name": item['name'], "price": item['price']}
                            items[i]['name'] = edit_name.strip()
                            items[i]['price'] = float(edit_price)
                            save_store_items(items)
                            log_action("STORE_ITEM_EDITED", old_data=old_data, new_data={"name": items[i]['name'], "price": items[i]['price']})
                            st.session_state.editing_store_item_idx = None
                            st.success(f"✅ Updated item: {edit_name}")
                            st.rerun()
                        else:
                            st.error("Item name cannot be empty.")
                    if cancel_btn:
                        st.session_state.editing_store_item_idx = None
                        st.rerun()
            else:
                c1, c2, c3, c4 = st.columns([3, 2, 1, 1])
                c1.markdown(f"• **{item['name']}**")
                c2.markdown(f"💷 **£{item['price']:.2f}**")
                with c3:
                    if st.button("✏️", key=f"edit_store_item_{i}", help="Edit Item"):
                        st.session_state.editing_store_item_idx = i
                        st.rerun()
                with c4:
                    if st.button("🗑️", key=f"del_store_item_{i}", help="Delete Item"):
                        items.pop(i)
                        save_store_items(items)
                        log_action("STORE_ITEM_DELETED", old_data={"name": item['name']})
                        st.success("Deleted.")
                        st.rerun()


# ============================================================
# 🧰 EMPLOYEE ITEMS — PDF REPORT HELPERS
# ============================================================
def _item_pdf_text(value):
    """Return PDF-safe text and break very long unspaced tokens.

    fpdf2 can raise ``FPDFException: Not enough horizontal space to render
    a single character`` when a single token is wider than the available
    cell width.  Item names, employee IDs, or imported values can contain
    long unbroken strings, so add harmless break opportunities.
    """
    text = str(value if value is not None else "").encode("latin-1", "replace").decode("latin-1")
    chunks = []
    for word in text.split(" "):
        if len(word) <= 45:
            chunks.append(word)
        else:
            chunks.extend(word[i:i + 45] for i in range(0, len(word), 45))
    return " ".join(chunks)


def _pdf_multiline(pdf, text, line_height=5):
    """Write a safe full-width PDF line, resetting x before each multi_cell."""
    width = max(10, pdf.w - pdf.l_margin - pdf.r_margin)
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(width, line_height, _item_pdf_text(text))


def _build_item_holdings_pdf(records, title="Employee Item Holdings"):
    if not PDF_AVAILABLE:
        return None
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()
    pdf.set_font("Arial", "B", 15)
    _pdf_multiline(pdf, title, 9)
    pdf.set_font("Arial", "", 9)
    _pdf_multiline(pdf, f"Generated: {datetime.now():%Y-%m-%d %H:%M:%S}", 6)
    pdf.ln(3)
    for r in records:
        outstanding = float(r.get("qty_outstanding", 0) or 0)
        if outstanding <= 0:
            continue
        pdf.set_font("Arial", "B", 10)
        _pdf_multiline(pdf, f"{r.get('emp_name','')} ({r.get('employee_id','')}) - {r.get('emp_dept','')}", 6)
        pdf.set_font("Arial", "", 9)
        lines = [
            f"Item: {r.get('item_name','')}",
            f"Issued: {r.get('qty_issued',0):g} | Returned: {r.get('qty_returned',0):g} | Outstanding: {outstanding:g}",
            f"Unit price: GBP {float(r.get('unit_price',0) or 0):.2f} | Outstanding value: GBP {outstanding * float(r.get('unit_price',0) or 0):.2f}",
            f"Issue date: {r.get('issue_date','')} | Issued by: {r.get('issued_by','')}",
        ]
        for line in lines:
            _pdf_multiline(pdf, line, 5)
        pdf.ln(2)
    out = pdf.output(dest="S")
    return out.encode("latin-1") if isinstance(out, str) else bytes(out)


def _build_item_checkin_pdf(c, title="Leaver Item Check-in"):
    if not PDF_AVAILABLE:
        return None
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()
    pdf.set_font("Arial", "B", 15)
    _pdf_multiline(pdf, title, 9)
    pdf.set_font("Arial", "", 9)
    _pdf_multiline(pdf, f"Check-in #{c.get('id','-')} | Generated: {datetime.now():%Y-%m-%d %H:%M:%S}", 6)
    for label in ("emp_name", "employee_id", "emp_dept", "leaving_date", "checkin_date", "checked_in_by"):
        _pdf_multiline(pdf, f"{label.replace('_',' ').title()}: {c.get(label,'')}", 5)
    pdf.ln(2)
    pdf.set_font("Arial", "B", 10); pdf.cell(0, 6, "Returned Items", ln=1)
    pdf.set_font("Arial", "", 9)
    for it in c.get("returned_items", []):
        _pdf_multiline(pdf, f"- {it.get('item_name','')} | Qty {it.get('quantity',0):g} | GBP {float(it.get('unit_price',0) or 0):.2f}", 5)
    pdf.ln(2); pdf.set_font("Arial", "B", 10); pdf.cell(0, 6, "Not Returned / Deducted", ln=1); pdf.set_font("Arial", "", 9)
    for it in c.get("not_returned_items", []):
        pdf.cell(0, 5, _item_pdf_text(f"- {it.get('item_name','')} | Qty {it.get('quantity',0):g} | Line total GBP {float(it.get('line_total',0) or 0):.2f}"), ln=1)
    pdf.ln(2); pdf.set_font("Arial", "B", 10); pdf.cell(0, 6, _item_pdf_text(f"Total Deduction: GBP {float(c.get('total_deduction',0) or 0):.2f}"), ln=1)
    if c.get("deduction_request_id"):
        pdf.set_font("Arial", "", 9); pdf.cell(0, 5, _item_pdf_text(f"Store Deduction Request: #{c.get('deduction_request_id')}"), ln=1)
    if c.get("notes"):
        pdf.multi_cell(0, 5, _item_pdf_text(f"Notes: {c.get('notes')}"))
    out = pdf.output(dest="S")
    return out.encode("latin-1") if isinstance(out, str) else bytes(out)


# ============================================================
# 🧰 EMPLOYEE ITEMS — UI
# ============================================================
def render_item_issue_form(user_name, user_dept):
    issue_form_version = int(st.session_state.get("item_issue_form_version", 0))
    st.subheader("🧰 Issue Items to Employee")
    st.caption(
        "Select the employee from the HR employee list, then record the tools, "
        "uniform, PPE and company property being issued. Store staff can issue to "
        "employees in any department."
    )

    employees = _hrp_load_employees()
    if not employees:
        st.warning(
            "No employees found in the HR employee list. HR must add the employee "
            "in **🏢 HR Department → HR Portal → Employee Directory** first."
        )
        return

    active = [e for e in employees
              if str(e.get("status", "Active")).strip().casefold() == "active"]

    if not active:
        st.info("There are no Active employees in the HR list.")
        return

    fc1, fc2 = st.columns([2, 3])
    with fc1:
        departments = sorted({str(e.get("department", "")).strip()
                              for e in active if str(e.get("department", "")).strip()})
        dept_filter = st.selectbox(
            "🏢 Filter by Department",
            ["All Departments"] + departments,
            key="item_issue_dept_filter",
        )
    with fc2:
        search = st.text_input(
            "🔎 Search employee",
            placeholder="Type name, ACE-ID, job title or department…",
            key="item_issue_search",
        )

    filtered = active
    if dept_filter != "All Departments":
        filtered = [e for e in filtered
                    if str(e.get("department", "")).strip() == dept_filter]
    if search.strip():
        q = search.strip().casefold()
        filtered = [e for e in filtered
                    if q in (f"{e.get('name','')} {e.get('emp_id','')} "
                             f"{e.get('job_title','')} {e.get('department','')}").casefold()]

    if not filtered:
        st.warning("No employee matches your search.")
        return

    filtered = sorted(filtered, key=lambda e: (str(e.get("name", "")).casefold(),
                                                str(e.get("emp_id", "")).casefold()))

    labels = [f"{e['name']} — {e['emp_id']}  ·  {e.get('department','')}  ·  {e.get('job_title','')}"
              for e in filtered]
    by_label = {lbl: e for lbl, e in zip(labels, filtered)}

    selected_label = st.selectbox(
        f"👤 Employee ({len(filtered)} shown)",
        labels,
        key="item_issue_emp_selector",
    )
    selected_emp = by_label.get(selected_label)
    if not selected_emp:
        return

    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Employee ID", selected_emp.get("emp_id", ""))
    s2.metric("Department", selected_emp.get("department", ""))
    s3.metric("Job Title", selected_emp.get("job_title", ""))
    s4.metric("Start Date", str(selected_emp.get("start_date", "")) or "-")

    st.divider()
    st.markdown("### ➕ Add Items to Issue")

    store_items_master = load_store_items()
    item_options = [it["name"] for it in store_items_master if it.get("active")]

    if "item_issue_rows" not in st.session_state:
        st.session_state.item_issue_rows = [
            {"item_name": "", "quantity": 1, "unit_price": 0.0, "notes": ""}
        ]

    for i, row in enumerate(st.session_state.item_issue_rows):
        c1, c2, c3, c4, c5 = st.columns([3, 1.1, 1.4, 2.2, 0.4])
        opts = [""] + item_options
        with c1:
            sel_item = st.selectbox(
                "Item", opts,
                index=opts.index(row.get("item_name", "")) if row.get("item_name") in opts else 0,
                key=f"issue_item_sel_{issue_form_version}_{i}",
            )
            if sel_item != row.get("item_name", ""):
                row["item_name"] = sel_item
                match = next((it for it in store_items_master if it["name"] == sel_item), None)
                if match:
                    row["unit_price"] = match["price"]
                st.rerun()
        with c2:
            row["quantity"] = st.number_input("Qty", min_value=1, step=1,
                                              value=int(row.get("quantity", 1)),
                                              key=f"issue_item_qty_{issue_form_version}_{i}")
        with c3:
            row["unit_price"] = st.number_input("Unit £", min_value=0.0, step=1.0,
                                                format="%.2f", value=float(row.get("unit_price", 0.0)),
                                                key=f"issue_item_price_{issue_form_version}_{i}")
        with c4:
            row["notes"] = st.text_input("Notes", value=row.get("notes", ""),
                                          key=f"issue_item_notes_{issue_form_version}_{i}")
        with c5:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑️", key=f"issue_item_del_{issue_form_version}_{i}"):
                st.session_state.item_issue_rows.pop(i)
                st.rerun()

    if st.button("➕ Add Another Item", key=f"issue_add_row_{issue_form_version}"):
        st.session_state.item_issue_rows.append(
            {"item_name": "", "quantity": 1, "unit_price": 0.0, "notes": ""}
        )
        st.rerun()

    total_value = sum(r.get("quantity", 1) * r.get("unit_price", 0.0)
                      for r in st.session_state.item_issue_rows if r.get("item_name"))
    st.markdown(f"**Total Value:** £{total_value:.2f}")
    st.divider()

    with st.form(f"issue_items_form_{issue_form_version}"):
        issue_date = st.date_input("📅 Issue Date", value=date.today())
        general_notes = st.text_area("📝 General Notes",
                                     placeholder="Optional notes about this issue…")
        submitted = st.form_submit_button("✅ Record Issue", type="primary", width="stretch")

    if submitted:
        valid_rows = [r for r in st.session_state.item_issue_rows
                      if r.get("item_name") and r.get("quantity", 0) > 0]
        if not valid_rows:
            st.error("Please add at least one item.")
            return

        all_items = load_employee_items()
        next_id = get_next_employee_item_id(all_items)
        for r in valid_rows:
            existing = next((x for x in all_items
                             if str(x.get("employee_id", "")).strip() == selected_emp["emp_id"]
                             and x.get("item_name") == r["item_name"]
                             and float(x.get("qty_outstanding", 0) or 0) > 0), None)
            if existing:
                existing["qty_issued"] += r["quantity"]
                existing["qty_outstanding"] += r["quantity"]
                existing["total_value"] = existing["qty_issued"] * existing["unit_price"]
                if r.get("notes"):
                    existing["notes"] = (str(existing.get("notes", "")) + " | " + r["notes"]).strip(" |")
            else:
                all_items.append({
                    "id": next_id,
                    "employee_id": selected_emp["emp_id"],
                    "emp_name": selected_emp["name"],
                    "emp_dept": selected_emp.get("department", ""),
                    "item_name": r["item_name"],
                    "qty_issued": float(r["quantity"]),
                    "qty_returned": 0.0,
                    "qty_outstanding": float(r["quantity"]),
                    "unit_price": float(r["unit_price"]),
                    "total_value": float(r["quantity"]) * float(r["unit_price"]),
                    "issue_date": str(issue_date),
                    "issued_by": user_name,
                    "status": "Issued",
                    "notes": r.get("notes", ""),
                })
                next_id += 1

        save_all_employee_items(all_items)
        log_action("EMPLOYEE_ITEMS_ISSUED", selected_emp["emp_id"],
                   new_data={"employee": selected_emp["name"],
                             "items": valid_rows,
                             "total": total_value})
        st.session_state.item_issue_rows = [
            {"item_name": "", "quantity": 1, "unit_price": 0.0, "notes": ""}
        ]
        st.session_state.item_issue_form_version = int(st.session_state.get("item_issue_form_version", 0)) + 1
        st.success(f"✅ {len(valid_rows)} item(s) recorded as issued to "
                   f"{selected_emp['name']} ({selected_emp['emp_id']}).")
        st.rerun()


def render_employee_holdings_overview(user_role=""):
    st.subheader("📊 Employee Item Holdings")
    st.caption("Current outstanding company items held by employees. Search by employee or department and download single/group PDF reports.")
    items = load_employee_items(force=True)
    outstanding_items = [r for r in items if float(r.get("qty_outstanding", 0) or 0) > 0]
    if not outstanding_items:
        st.info("No items are currently outstanding with employees.")
        return
    total_value = sum(float(r.get("qty_outstanding", 0) or 0) * float(r.get("unit_price", 0) or 0) for r in outstanding_items)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Employees Holding", len({str(r.get('employee_id','')) for r in outstanding_items}))
    c2.metric("Outstanding Lines", len(outstanding_items))
    c3.metric("Outstanding Qty", sum(float(r.get('qty_outstanding',0) or 0) for r in outstanding_items))
    c4.metric("Outstanding Value", f"£{total_value:,.2f}")
    st.divider()
    departments = sorted({str(r.get("emp_dept", "")).strip() for r in outstanding_items if str(r.get("emp_dept", "")).strip()})
    employees = sorted({f"{r.get('emp_name','')} — {r.get('employee_id','')}" for r in outstanding_items}, key=str.casefold)
    f1, f2, f3 = st.columns([2, 2, 3])
    with f1:
        dept_filter = st.selectbox("🏢 Department", ["All Departments"] + departments, key="holdings_dept_filter")
    with f2:
        employee_filter = st.selectbox("👤 Employee", ["All Employees"] + employees, key="holdings_employee_filter")
    with f3:
        search = st.text_input("🔎 Search", placeholder="Employee, ACE-ID, item, department...", key="holdings_search")
    filtered = list(outstanding_items)
    if dept_filter != "All Departments": filtered = [r for r in filtered if str(r.get("emp_dept","")) == dept_filter]
    if employee_filter != "All Employees":
        eid = employee_filter.rsplit(" — ", 1)[-1].strip()
        filtered = [r for r in filtered if str(r.get("employee_id","")) == eid]
    if search.strip():
        q=search.casefold().strip(); filtered=[r for r in filtered if q in " ".join(str(v) for v in r.values()).casefold()]
    if not filtered:
        st.info("No holdings match the selected filters.")
        return
    st.dataframe(pd.DataFrame([{
        "ID": r["id"], "Employee ID": r["employee_id"], "Employee": r["emp_name"], "Department": r["emp_dept"],
        "Item": r["item_name"], "Issued": r["qty_issued"], "Returned": r["qty_returned"], "Outstanding": r["qty_outstanding"],
        "Unit £": f"£{float(r['unit_price']):.2f}", "Outstanding Value": f"£{float(r['qty_outstanding'])*float(r['unit_price']):.2f}",
        "Issue Date": r["issue_date"], "Issued By": r["issued_by"], "Status": r["status"]
    } for r in filtered]), width="stretch", hide_index=True)
    st.divider()
    p1,p2,p3 = st.columns(3)
    with p1:
        pdf = _build_item_holdings_pdf(filtered, "Employee Item Holdings - Selected Filters")
        st.download_button("📄 Download Selected Holdings PDF", data=pdf or b"", file_name="Employee_Item_Holdings_Selected.pdf", mime="application/pdf", disabled=not bool(pdf), key="holdings_pdf_selected")
    with p2:
        if employee_filter != "All Employees":
            emp_pdf = _build_item_holdings_pdf(filtered, f"Employee Item Holdings - {employee_filter}")
            st.download_button("👤 Download Single Employee PDF", data=emp_pdf or b"", file_name="Employee_Item_Holdings_Employee.pdf", mime="application/pdf", disabled=not bool(emp_pdf), key="holdings_pdf_employee")
        else:
            st.caption("Select one employee for a single-employee PDF.")
    with p3:
        if dept_filter != "All Departments":
            dept_pdf = _build_item_holdings_pdf(filtered, f"Employee Item Holdings - {dept_filter}")
            st.download_button("🏢 Download Department PDF", data=dept_pdf or b"", file_name="Employee_Item_Holdings_Department.pdf", mime="application/pdf", disabled=not bool(dept_pdf), key="holdings_pdf_department")
        else:
            st.caption("Select a department for a department-wise PDF.")


def render_item_checkin_form(user_name):
    st.subheader("📋 Leaver Item Check-in")
    st.caption(
        "Confirm which items a leaving employee has returned. Items not returned "
        "automatically generate a Store Deduction request for Director approval."
    )

    employees = _hrp_load_employees()
    if not employees:
        st.warning("No employees found in the HR employee list.")
        return

    all_items = load_employee_items()
    outstanding_ids = {str(r.get("employee_id", "")).strip()
                       for r in all_items
                       if float(r.get("qty_outstanding", 0) or 0) > 0}
    candidates = [e for e in employees
                  if str(e.get("emp_id", "")).strip() in outstanding_ids]

    if not candidates:
        st.info("No employees currently hold any outstanding items.")
        return

    fc1, fc2 = st.columns([2, 3])
    with fc1:
        departments = sorted({str(e.get("department", "")).strip()
                              for e in candidates if str(e.get("department", "")).strip()})
        dept_filter = st.selectbox(
            "🏢 Filter by Department",
            ["All Departments"] + departments,
            key="checkin_dept_filter",
        )
    with fc2:
        search = st.text_input(
            "🔎 Search employee",
            placeholder="Type name, ACE-ID, department…",
            key="checkin_search",
        )

    filtered = candidates
    if dept_filter != "All Departments":
        filtered = [e for e in filtered
                    if str(e.get("department", "")).strip() == dept_filter]
    if search.strip():
        q = search.strip().casefold()
        filtered = [e for e in filtered
                    if q in (f"{e.get('name','')} {e.get('emp_id','')} "
                             f"{e.get('department','')}").casefold()]

    if not filtered:
        st.warning("No employee matches your search.")
        return

    def _outstanding_count(eid):
        return sum(1 for r in all_items
                   if str(r.get("employee_id", "")).strip() == str(eid).strip()
                   and float(r.get("qty_outstanding", 0) or 0) > 0)

    filtered = sorted(filtered, key=lambda e: (str(e.get("name", "")).casefold(),
                                                str(e.get("emp_id", "")).casefold()))
    labels = [f"{e['name']} — {e['emp_id']} · {e.get('department','')} "
              f"({_outstanding_count(e['emp_id'])} outstanding)"
              for e in filtered]
    by_label = {lbl: e for lbl, e in zip(labels, filtered)}

    selected = st.selectbox(f"👤 Employee ({len(filtered)} shown)", labels,
                            key="checkin_emp_selector")
    selected_emp = by_label.get(selected)
    if not selected_emp:
        return

    outstanding = get_employee_outstanding_items(selected_emp["emp_id"])
    st.info(f"Checking in items for **{selected_emp['name']}** "
            f"({selected_emp['emp_id']}) · {selected_emp.get('department','')}")
    st.markdown(f"### 📦 Items Held ({len(outstanding)})")
    st.caption("Enter the quantity returned for each item. Leave at 0 if not "
               "returned — those items will be deducted.")

    key = f"checkin_qty_{selected_emp['emp_id']}"
    if key not in st.session_state:
        st.session_state[key] = {r["id"]: 0.0 for r in outstanding}
    for r in outstanding:
        st.session_state[key].setdefault(r["id"], 0.0)

    returned_items, not_returned_items = [], []
    total_deduction = 0.0

    for r in outstanding:
        outstanding_qty = float(r["qty_outstanding"])
        unit_price = float(r["unit_price"])
        c1, c2, c3 = st.columns([4, 2, 3])
        with c1:
            st.markdown(
                f"**{r['item_name']}**  \n"
                f"Unit £{unit_price:.2f} · Issued {r['qty_issued']:.0f} · "
                f"Already returned {r['qty_returned']:.0f} · Outstanding {outstanding_qty:.0f}"
            )
        with c2:
            returned_qty = st.number_input(
                "Qty Returned", min_value=0.0, max_value=outstanding_qty, step=1.0,
                value=float(st.session_state[key][r["id"]]),
                key=f"checkin_ret_{r['id']}",
            )
            st.session_state[key][r["id"]] = returned_qty
        with c3:
            not_ret = outstanding_qty - returned_qty
            deduction = not_ret * unit_price
            if not_ret > 0:
                st.markdown(
                    f"<div style='color:#ef4444;font-weight:bold;padding-top:6px;'>"
                    f"Not returned: {not_ret:.0f}<br>Deduction: £{deduction:.2f}</div>",
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    "<div style='color:#10b981;font-weight:bold;padding-top:6px;'>"
                    "✅ Fully returned</div>",
                    unsafe_allow_html=True,
                )
        if returned_qty > 0:
            returned_items.append({"item_id": r["id"], "item_name": r["item_name"],
                                   "quantity": returned_qty, "unit_price": unit_price})
        if not_ret > 0:
            not_returned_items.append({"item_id": r["id"], "item_name": r["item_name"],
                                       "quantity": not_ret, "unit_price": unit_price,
                                       "line_total": deduction})
            total_deduction += deduction

    st.divider()
    st.markdown(
        f"<h3 style='text-align:right;'>Total Deduction for Unreturned Items: "
        f"<span style='color:#ef4444;'>£{total_deduction:.2f}</span></h3>",
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)
    with col1:
        leaving_date = st.date_input(
            "📅 Leaving Date",
            value=selected_emp.get("leaving_date") or date.today(),
            key="checkin_leaving_date",
        )
    with col2:
        checkin_date = st.date_input("📅 Check-in Date", value=date.today(),
                                     key="checkin_date")
    notes = st.text_area("📝 Check-in Notes", key="checkin_notes")
    current_checkin_pdf = _build_item_checkin_pdf({
        "id": "Draft", "employee_id": selected_emp.get("emp_id", ""), "emp_name": selected_emp.get("name", ""),
        "emp_dept": selected_emp.get("department", ""), "leaving_date": str(leaving_date), "checkin_date": str(checkin_date),
        "checked_in_by": user_name, "returned_items": returned_items, "not_returned_items": not_returned_items,
        "total_deduction": total_deduction, "notes": notes, "deduction_request_id": ""
    }, "Leaver Item Check-in - Draft")
    st.download_button("📄 Download Check-in PDF", data=current_checkin_pdf or b"", file_name="Leaver_Item_Checkin_Draft.pdf", mime="application/pdf", disabled=not bool(current_checkin_pdf), key="checkin_draft_pdf")

    if not returned_items and not not_returned_items:
        st.warning("Nothing recorded yet — enter the returned quantities above.")
        return

    st.divider()
    b1, b2 = st.columns(2)
    with b1:
        if st.button("💾 Save Check-in (Returned Only)",
                     disabled=not returned_items, key="save_checkin_only",
                     width="stretch"):
            _save_item_checkin(user_name, selected_emp, returned_items,
                               not_returned_items, total_deduction,
                               leaving_date, checkin_date, notes,
                               create_deduction=False)
    with b2:
        if st.button(f"📤 Save & Create Store Deduction (£{total_deduction:.2f})",
                     disabled=total_deduction <= 0, type="primary",
                     key="save_checkin_and_deduct", width="stretch"):
            _save_item_checkin(user_name, selected_emp, returned_items,
                               not_returned_items, total_deduction,
                               leaving_date, checkin_date, notes,
                               create_deduction=True)


def _save_item_checkin(user_name, employee, returned_items, not_returned_items,
                       total_deduction, leaving_date, checkin_date, notes,
                       create_deduction=False):
    """Persist the leaver check-in and create linked Store transactions when needed."""
    checkins = load_item_checkins()
    new_id = get_next_item_checkin_id(checkins)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    deduction_req_id = ""
    return_req_id = ""  # kept for compatibility with older check-in records; no new returns are created

    store_records = load_store_deductions()

    # Items not returned can create a Store Deduction for Director approval.
    if create_deduction and not_returned_items:
        ded_id = get_next_store_deduction_id(store_records)
        ded_items = [{
            "item_name": it["item_name"],
            "quantity": int(it["quantity"]),
            "price": float(it["unit_price"]),
        } for it in not_returned_items]
        ded_desc = (f"Auto-generated from Leaver Item Check-in #{new_id}. "
                    f"Items not returned by {employee['name']} ({employee['emp_id']}). "
                    f"{(notes or '').strip()}").strip()
        ded_rec = {
            "id": ded_id,
            "emp_name": employee["name"],
            "date_leaving": str(leaving_date),
            "emp_dept": employee.get("department", ""),
            "manager": user_name,
            "date_submit": str(checkin_date),
            "type": "Deduction",
            "items": ded_items,
            "total_deduction": float(total_deduction),
            "desc": ded_desc,
            "attachment_name": "None",
            "status": "pending",
            "director_comments": "",
            "rejection_reason": "",
            "decision_date": "",
            "decision_by": "",
            "submitted_by": user_name,
            "submitted_date": now,
            "pdf_path": "",
        }
        store_records.append(ded_rec)
        deduction_req_id = str(ded_id)
        log_action("STORE_DEDUCTION_CREATED", ded_id, new_data=ded_rec)

    # Returned company property is simply cleared from the employee's holdings.
    # IMPORTANT: a returned item does NOT create a Store Addition transaction.
    # Only items that remain unreturned can create a Store Deduction for Director approval.
    # Save Store data only when a non-returned item created a deduction.
    if deduction_req_id:
        save_all_store_deductions(store_records)

    checkins.append({
        "id": new_id,
        "employee_id": employee["emp_id"],
        "emp_name": employee["name"],
        "emp_dept": employee.get("department", ""),
        "leaving_date": str(leaving_date),
        "checkin_date": str(checkin_date),
        "checked_in_by": user_name,
        "returned_items": returned_items,
        "not_returned_items": not_returned_items,
        "total_deduction": float(total_deduction),
        "deduction_request_id": deduction_req_id,
        "return_request_id": return_req_id,
        "status": "Completed",
        "notes": (notes or "").strip(),
    })
    save_all_item_checkins(checkins)

    items = load_employee_items()
    for r in items:
        if str(r.get("employee_id", "")).strip() != str(employee["emp_id"]).strip():
            continue
        ret = next((x for x in returned_items if x["item_id"] == r["id"]), None)
        not_ret = next((x for x in not_returned_items if x["item_id"] == r["id"]), None)
        if ret:
            r["qty_returned"] = float(r.get("qty_returned", 0)) + float(ret["quantity"])
            r["qty_outstanding"] = float(r.get("qty_outstanding", 0)) - float(ret["quantity"])
        if not_ret:
            r["qty_outstanding"] = float(r.get("qty_outstanding", 0)) - float(not_ret["quantity"])
            tag = f"Deducted {not_ret['quantity']:.0f} via Deduction #{deduction_req_id}" if deduction_req_id \
                  else f"Not returned {not_ret['quantity']:.0f}"
            r["notes"] = (str(r.get("notes", "")) + f" | {tag}").strip(" |")
        if float(r.get("qty_outstanding", 0)) <= 0:
            r["status"] = "Returned" if float(r.get("qty_returned", 0)) > 0 else "Deducted"
        elif float(r.get("qty_returned", 0)) > 0:
            r["status"] = "Partially Returned"
    save_all_employee_items(items)

    log_action("ITEM_CHECKIN_COMPLETED", new_id,
               new_data={"employee": employee["name"],
                         "returned": len(returned_items),
                         "not_returned": len(not_returned_items),
                         "deduction": total_deduction,
                         "deduction_id": deduction_req_id,
                         "return_id": return_req_id})

    st.session_state.pop(f"checkin_qty_{employee['emp_id']}", None)

    if deduction_req_id and return_req_id:
        st.success(
            f"✅ Check-in #{new_id} saved. "
            f"Store Deduction #{deduction_req_id} and Store Return #{return_req_id} "
            f"created — both sent to Director for approval."
        )
    elif deduction_req_id:
        st.success(
            f"✅ Check-in #{new_id} saved. Store Deduction #{deduction_req_id} created for "
            f"£{total_deduction:.2f} — sent to Director for approval."
        )
    else:
        st.success(f"✅ Check-in #{new_id} saved. Holdings updated.")
    st.rerun()


def render_item_checkin_history():
    st.subheader("📚 Leaver Check-in History")
    checkins = load_item_checkins()
    if not checkins:
        st.info("No check-in sessions recorded yet.")
        return

    store_records = load_store_deductions(force=True)

    def _linked_store_status(request_id):
        if not request_id:
            return None
        record = next(
            (r for r in store_records if str(r.get("id", "")) == str(request_id)),
            None,
        )
        if not record:
            return None
        status = str(record.get("status", "pending")).strip().lower()
        if status == "approved":
            decision_by = str(record.get("decision_by", "Director")).strip() or "Director"
            decision_date = str(record.get("decision_date", "")).strip()
            suffix = f" on {decision_date}" if decision_date else ""
            return f"🟢 Approved by {decision_by}{suffix}", "success"
        if status == "rejected":
            decision_by = str(record.get("decision_by", "Director")).strip() or "Director"
            decision_date = str(record.get("decision_date", "")).strip()
            suffix = f" on {decision_date}" if decision_date else ""
            return f"🔴 Rejected by {decision_by}{suffix}", "error"
        return "🟡 Pending Director approval", "warning"

    for c in reversed(checkins):
        returned_count = len(c.get("returned_items", []))
        not_returned_count = len(c.get("not_returned_items", []))
        with st.expander(
            f"#{c['id']} | {c['emp_name']} | {c['emp_dept']} | "
            f"↩️ {returned_count} returned | ❌ {not_returned_count} not returned | {c['checkin_date']}"
        ):
            st.write(f"👤 {c['emp_name']} ({c['employee_id']}) · 🏢 {c['emp_dept']}")
            st.write(f"📅 Leaving: {c['leaving_date']} | Check-in: {c['checkin_date']} | By: {c['checked_in_by']}")

            if c.get("return_request_id"):
                linked = _linked_store_status(c["return_request_id"])
                if linked:
                    message, kind = linked
                    text = f"🔗 Linked Store Return (Addition) #{c['return_request_id']} — {message}"
                    if kind == "success":
                        st.success(text)
                    elif kind == "error":
                        st.error(text)
                    else:
                        st.warning(text)
                else:
                    st.info(f"🔗 Linked Store Return (Addition) #{c['return_request_id']} — record not found")

            if c.get("deduction_request_id"):
                linked = _linked_store_status(c["deduction_request_id"])
                if linked:
                    message, kind = linked
                    text = f"🔗 Linked Store Deduction #{c['deduction_request_id']} — {message}"
                    if kind == "success":
                        st.success(text)
                    elif kind == "error":
                        st.error(text)
                    else:
                        st.warning(text)
                else:
                    st.info(f"🔗 Linked Store Deduction #{c['deduction_request_id']} — record not found")

            st.markdown(f"**Returned Items ({returned_count}):**")
            if c.get("returned_items"):
                st.dataframe(pd.DataFrame([{
                    "Item": it["item_name"], "Qty": it["quantity"], "Unit £": it["unit_price"]
                } for it in c["returned_items"]]), width="stretch", hide_index=True)
            else:
                st.caption("None")

            st.markdown(f"**Not Returned — Deducted ({not_returned_count}):**")
            if c.get("not_returned_items"):
                st.dataframe(pd.DataFrame([{
                    "Item": it["item_name"], "Qty": it["quantity"],
                    "Unit £": it["unit_price"], "Line Total £": it.get("line_total", 0)
                } for it in c["not_returned_items"]]), width="stretch", hide_index=True)
            else:
                st.caption("None")

            if c.get("notes"):
                st.info(f"📝 {c['notes']}")
            history_pdf = _build_item_checkin_pdf(c, f"Leaver Item Check-in #{c.get('id','')}")
            st.download_button("📄 Download Check-in PDF", data=history_pdf or b"", file_name=f"Leaver_Item_Checkin_{c.get('id','')}.pdf", mime="application/pdf", disabled=not bool(history_pdf), key=f"checkin_history_pdf_{c.get('id','')}")


def initialise_excel():
    safe_init_excel(EXCEL_PATH, EXCEL_COLUMNS)

initialise_excel()
initialise_work_orders()
initialise_inspector_bonus()
initialise_hr_leave()
initialise_store_deduction()
initialise_employee_items()

def load_records_from_excel(force=False):
    if not force and "_records_cache" in st.session_state:
        return list(st.session_state["_records_cache"])
    try:
        if not os.path.exists(EXCEL_PATH): return _set_data_cache("_records_cache", []).copy()
        df = _read_excel_records(EXCEL_PATH)
        if df.empty: return _set_data_cache("_records_cache", []).copy()
        parsed = []
        for r in df.to_dict(orient="records"):
            try: record_id = int(r.get("ID", 0))
            except: record_id = 0
            try: amount = float(r.get("Amount (£)", 0))
            except: amount = 0.0
            parsed.append({"id": record_id, "emp_name": str(r.get("Employee Name", "Not Specified")).strip(), "dept": str(r.get("Department", "Not Specified")).strip(), "type": str(r.get("Transaction Type", "Not Specified")).strip(), "category": str(r.get("Category Reason", "Not Specified")).strip(), "date": str(r.get("Date", "")).strip(), "amount": amount, "manager": str(r.get("Line Manager", "Not Specified")).strip(), "desc": str(r.get("Description", "")).strip(), "attachment_name": str(r.get("Attachment Name", "None")).strip(), "status": str(r.get("Status", "pending")).strip().lower(), "director_comments": str(r.get("Director Comments", "")).strip(), "decision_date": str(r.get("Decision Date", "")).strip(), "decision_by": str(r.get("Decision By", "")).strip(), "submitted_by": str(r.get("Submitted By", "")).strip(), "pdf_path": str(r.get("PDF File Path", "")).strip(), "edited_from_id": str(r.get("Edited From ID", "")).strip(), "old_data": str(r.get("Old Data", "")).strip()})
        _set_data_cache("_records_cache", parsed)
        return list(parsed)
    except Exception as e:
        st.error(f"Load Error: {e}")
        return []

def save_all_records(records):
    export = []
    for r in records:
        export.append({"ID": int(r.get("id", 0)), "Employee Name": str(r.get("emp_name", "")), "Department": str(r.get("dept", "")), "Transaction Type": str(r.get("type", "")), "Category Reason": str(r.get("category", "")), "Date": str(r.get("date", "")), "Amount (£)": float(r.get("amount", 0.0)), "Line Manager": str(r.get("manager", "")), "Description": str(r.get("desc", "")), "Attachment Name": str(r.get("attachment_name", "None")), "Status": str(r.get("status", "pending")).lower(), "Director Comments": str(r.get("director_comments", "")), "Decision Date": str(r.get("decision_date", "")), "Decision By": str(r.get("decision_by", "")), "Submitted By": str(r.get("submitted_by", "")), "PDF File Path": str(r.get("pdf_path", "")), "Edited From ID": str(r.get("edited_from_id", "")), "Old Data": str(r.get("old_data", ""))})
    pd.DataFrame(export, columns=EXCEL_COLUMNS).to_excel(EXCEL_PATH, index=False, engine="openpyxl")
    _set_data_cache("_records_cache", list(records))
    sync_saved_file_to_drive(EXCEL_PATH)

def save_record_to_excel(new_record):
    current = load_records_from_excel()
    current.append(new_record)
    save_all_records(current)

def get_submitted_by(req):
    value = str(req.get("submitted_by", "") or "").strip()
    if value and value.lower() not in ("nan", "none", "-"): return value
    try:
        req_id = str(req.get("id", "")).strip()
        if req_id:
            for entry in reversed(load_audit_log()):
                action = str(entry.get("Action", "")).strip().upper()
                audit_req_id = str(entry.get("Request_ID", "")).strip()
                if action == "CREATED" and audit_req_id == req_id:
                    creator = str(entry.get("User_Name", "")).strip()
                    if creator and creator.lower() not in ("nan", "none", "-"): return creator
    except Exception: pass
    return ""

def get_completed_by(req):
    for key in ("completed_by", "decision_by", "approved_by"):
        value = str(req.get(key, "") or "").strip()
        if value and value.lower() not in ("nan", "none", "-", "director"): return value
    return ""

def generate_approval_pdf(request_data):
    if not PDF_AVAILABLE: return False, None, "Install fpdf2: pip install fpdf2"
    try:
        req_id = request_data.get("id")
        all_recs = load_records_from_excel()
        fresh_data = next((r for r in all_recs if int(str(r.get("id", "0"))) == int(str(req_id))), request_data)
        def clean_text(t):
            if t is None: return ""
            t = str(t).replace("\u2013", "-").replace("\u2014", "-").replace("\u2212", "-").replace("\u25b3", "(triangle)")
            t = t.encode("latin-1", "ignore").decode("latin-1")
            for c in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']: t = t.replace(c, " ")
            return t.strip()
        emp_name = clean_text(fresh_data.get("emp_name", "Unknown"))
        dept = clean_text(fresh_data.get("dept", ""))
        category = clean_text(fresh_data.get("category", ""))
        amount_val = fresh_data.get("amount", "0")
        amount = f"{float(amount_val):.2f}" if str(amount_val).replace('.','').isdigit() else str(amount_val)
        req_date = format_date(fresh_data.get("date", ""))
        desc = clean_text(fresh_data.get("desc", ""))
        manager = clean_text(fresh_data.get("manager", ""))
        status = str(fresh_data.get("status", "pending")).strip().lower()
        dir_approve = format_date(fresh_data.get("decision_date", ""))
        dir_name = clean_text(fresh_data.get("decision_by", "Director"))
        dir_comments = clean_text(fresh_data.get("director_comments", ""))
        att_names = ""
        for key in ["attachment_name", "Attachment Name", "attachment", "Attachment"]:
            val = str(fresh_data.get(key, "")).strip()
            if val and val.lower() not in ["none", "nan", ""]:
                att_names = val; break
        att_names = clean_text(att_names)
        display_files = []
        if att_names:
            for name in att_names.split(","):
                n = name.strip()
                if n and n.lower() not in ["none", ""]: display_files.append(n)
        pdf = FPDF()
        pdf.add_page()
        if os.path.exists(LOGO_PATH): pdf.image(LOGO_PATH, x=75, y=10, w=60)
        pdf.ln(22)
        pdf.set_font("Courier", "", 11)
        pdf.cell(0, 5, txt="Addition & Deduction Approval Form", ln=True, align="C")
        pdf.ln(3)
        line_y = pdf.get_y()
        pdf.line(10, line_y, 200, line_y); pdf.line(10, line_y + 1.5, 200, line_y + 1.5)
        pdf.ln(12)
        pdf.set_font("Courier", "B", 10)
        pdf.cell(0, 5, txt="REQUEST DETAILS", ln=True); pdf.ln(2)
        pdf.set_font("Courier", "", 9)
        pdf.cell(52, 5, "Request ID:", 0, 0); pdf.cell(0, 5, str(fresh_data.get("id", "")), ln=True)
        pdf.cell(52, 5, "Employee Name:", 0, 0); pdf.cell(0, 5, emp_name, ln=True)
        pdf.cell(52, 5, "Department:", 0, 0); pdf.cell(0, 5, dept, ln=True)
        pdf.cell(52, 5, "Transaction Type:", 0, 0); pdf.cell(0, 5, clean_text(fresh_data.get("type", "")), ln=True)
        pdf.cell(52, 5, "Category / Reason:", 0, 0); pdf.cell(0, 5, category, ln=True)
        pdf.cell(52, 5, "Request Date:", 0, 0); pdf.cell(0, 5, req_date, ln=True)
        pdf.cell(52, 5, "Amount:", 0, 0); pdf.cell(0, 5, f"£{amount}", ln=True)
        pdf.cell(52, 5, "Line Manager:", 0, 0); pdf.cell(0, 5, manager, ln=True)
        submitted_by = clean_text(get_submitted_by(fresh_data))
        if submitted_by:
            pdf.cell(52, 5, "Submitted By:", 0, 0); pdf.cell(0, 5, submitted_by, ln=True)
        pdf.ln(6)
        pdf.set_font("Courier", "B", 10)
        pdf.cell(0, 5, txt="DESCRIPTION / JUSTIFICATION", ln=True); pdf.ln(2)
        pdf.set_font("Courier", "", 9)
        pdf.multi_cell(0, 5, desc); pdf.ln(8)
        pdf.set_font("Courier", "B", 10)
        pdf.cell(0, 5, txt="DIRECTOR APPROVAL", ln=True); pdf.ln(2)
        pdf.set_font("Courier", "", 9)
        if status == "approved":
            pdf.cell(52, 5, "Decision:", 0, 0)
            pdf.set_font("Courier", "B", 9); pdf.set_text_color(0, 128, 0)
            pdf.cell(0, 5, "APPROVED", ln=True); pdf.set_text_color(0, 0, 0); pdf.set_font("Courier", "", 9)
            pdf.cell(52, 5, "Approved By:", 0, 0); pdf.cell(0, 5, dir_name, ln=True)
            pdf.cell(52, 5, "Approval Date / Time:", 0, 0); pdf.cell(0, 5, dir_approve if dir_approve != "-" else "-", ln=True)
            if dir_comments and dir_comments not in ["None", ""]:
                pdf.ln(2); pdf.set_font("Courier", "B", 9); pdf.cell(52, 5, "Director Comments:", 0, 0)
                pdf.set_font("Courier", "", 9); pdf.ln(5); pdf.multi_cell(0, 5, dir_comments)
        elif status == "rejected":
            pdf.cell(52, 5, "Decision:", 0, 0)
            pdf.set_font("Courier", "B", 9); pdf.set_text_color(200, 0, 0)
            pdf.cell(0, 5, "REJECTED", ln=True); pdf.set_text_color(0, 0, 0); pdf.set_font("Courier", "", 9)
            pdf.cell(52, 5, "Rejected By:", 0, 0); pdf.cell(0, 5, dir_name, ln=True)
            pdf.cell(52, 5, "Rejection Date / Time:", 0, 0); pdf.cell(0, 5, dir_approve if dir_approve != "-" else "-", ln=True)
            if dir_comments and dir_comments not in ["None", ""]:
                pdf.ln(2); pdf.set_font("Courier", "B", 9); pdf.cell(52, 5, "Reason for Rejection:", 0, 0)
                pdf.set_font("Courier", "", 9); pdf.ln(5); pdf.multi_cell(0, 5, dir_comments)
        else:
            pdf.cell(52, 5, "Decision:", 0, 0); pdf.cell(0, 5, "Pending", ln=True)
        pdf.ln(12)
        dash_y = pdf.get_y()
        for x in range(10, 200, 4): pdf.line(x, dash_y, x + 2, dash_y)
        if status == "approved" and os.path.exists(APPROVED_STAMP_PATH): pdf.image(APPROVED_STAMP_PATH, x=75, y=dash_y - 6, w=60)
        elif status == "rejected" and os.path.exists(REJECTED_STAMP_PATH): pdf.image(REJECTED_STAMP_PATH, x=75, y=dash_y - 6, w=60)
        pdf.ln(8)
        pdf.set_font("Courier", "", 8)
        pdf.cell(0, 5, txt="Authorised Signature / Director", ln=True)
        pdf.add_page()
        pdf.set_font("Courier", "B", 12)
        pdf.cell(0, 8, txt="ATTACHMENTS", ln=True); pdf.ln(6)
        pdf.set_font("Courier", "", 9)
        if len(display_files) > 0:
            for idx, fname in enumerate(display_files, 1):
                file_path = os.path.join(UPLOAD_DIR, fname)
                if os.path.exists(file_path):
                    if fname.lower().endswith((".png", ".jpg", ".jpeg")):
                        pdf.ln(2)
                        try: pdf.image(file_path, x=10, w=190); pdf.ln(70)
                        except: pdf.cell(0, 5, "     Warning: Preview could not be displayed", ln=True); pdf.ln(3)
                    else: pdf.cell(0, 5, "     Non-image file - see original upload", ln=True); pdf.ln(3)
                else: pdf.cell(0, 5, "     Warning: File not found on server", ln=True); pdf.ln(3)
        else: pdf.cell(0, 6, "- No files were attached to this request", ln=True)
        safe_id = clean_text(str(req_id))
        safe_name = emp_name
        safe_category = category
        safe_date = datetime.now().strftime("%Y-%m-%d")
        filename = f"{safe_id}# {safe_name} - {safe_category} - {safe_date}.pdf"
        pdf_bytes = bytes(pdf.output())
        os.makedirs(PDF_DIR, exist_ok=True)
        full_pdf_path = os.path.join(PDF_DIR, filename)
        with open(full_pdf_path, "wb") as f: f.write(pdf_bytes)
        _upload_to_drive_bg(full_pdf_path, filename)
        return True, pdf_bytes, filename
    except Exception as e:
        return False, None, f"PDF Error: {str(e)}"

def display_pdf_button(req, can_generate=True, key_suffix=""):
    if not PDF_AVAILABLE:
        st.warning("⚠️ PDF generation is unavailable. Please install fpdf2.")
        return False
    req_id = str(req.get("id", "unknown"))
    status = str(req.get("status", "pending")).strip().lower()
    suffix = str(key_suffix or "default").replace(" ", "_")
    state_key = f"pdf_result_{req_id}_{status}_{suffix}"
    button_key = f"generate_pdf_{req_id}_{status}_{suffix}"
    download_key = f"download_pdf_{req_id}_{status}_{suffix}"
    if can_generate:
        if st.button(f"📄 Generate PDF for ID #{req_id}", type="primary", key=button_key):
            with st.spinner("Generating PDF..."):
                ok, pdf_bytes, filename = generate_approval_pdf(req)
            if ok and pdf_bytes:
                st.session_state[state_key] = {"data": bytes(pdf_bytes), "filename": filename}
                st.success("✅ PDF generated successfully. Download it below.")
            else:
                st.session_state.pop(state_key, None)
                st.error(f"❌ {filename}")
    result = st.session_state.get(state_key)
    if result:
        st.download_button("📥 Download PDF", data=result["data"], file_name=result["filename"], mime="application/pdf", type="primary", key=download_key)
        return True
    return False

def ensure_attachment_local(filename):
    if not filename or str(filename).strip().lower() in ("none", "nan", ""): return None
    filename = os.path.basename(str(filename).strip())
    local_path = os.path.join(UPLOAD_DIR, filename)
    if os.path.exists(local_path): return local_path
    if drive_service is not None:
        remote = _drive_find_file(filename)
        if remote and _drive_download_file(remote["id"], local_path): return local_path
    return None

def display_attachments(req):
    user_info = st.session_state.get("user_info", {})
    user_role = str(user_info.get("role", "")).strip().lower()
    is_director = (user_role == "director")
    att = req.get("attachment_name", "None")
    if not att or str(att).strip().lower() in ["none", "nan", ""]:
        st.info("📎 No attachments.")
        return
    try:
        attached_files = [n.strip() for n in str(att).split(",") if n.strip()]
        found_any = False
        for idx, name in enumerate(attached_files):
            path = ensure_attachment_local(name)
            if not path: continue
            found_any = True
            is_image = name.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".webp"))
            if is_director and is_image:
                st.markdown(f"### 🖼️ {name}")
                st.image(path, caption=name, width="stretch")
                with open(path, "rb") as f:
                    st.download_button(label=f"⬇️ Download {name}", data=f.read(), file_name=name, mime="image/*", key=f"director_att_{req.get('id', idx)}_{idx}")
            else:
                with open(path, "rb") as f: file_data = f.read()
                st.download_button(label=f"⬇️ Download {name}", data=file_data, file_name=name, key=f"attachment_{req.get('id', idx)}_{idx}")
        if not found_any: st.info("📎 Attachments referenced but files are not available.")
    except Exception as e:
        st.error(f"❌ Could not display attachments: {e}")

if "logged_in" not in st.session_state: st.session_state.logged_in = False
if "user_info" not in st.session_state: st.session_state.user_info = {}
if "editing_request_id" not in st.session_state: st.session_state.editing_request_id = None
if "editing_inspector_bonus_id" not in st.session_state: st.session_state.editing_inspector_bonus_id = None
if "editing_work_order_id" not in st.session_state: st.session_state.editing_work_order_id = None
if "editing_work_order_id_emp" not in st.session_state: st.session_state.editing_work_order_id_emp = None

def display_company_header():
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        if os.path.exists(LOGO_PATH): st.image(LOGO_PATH, width=300)
        else: st.title("⚡ ACOOLE ELECTRICAL LTD")
        st.caption("Acoole Operations & Authorisation Portal")
        st.divider()

def change_my_password_form():
    if not st.session_state.get("logged_in") or not st.session_state.get("user_info"): return
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
                if USERS[current_username]["password"] != old_pass: st.error("❌ Current password is NOT correct!"); return
                if new_pass1 != new_pass2: st.error("❌ New passwords do NOT match!"); return
                if len(new_pass1) < 4: st.error("❌ New password must be at least 4 characters!"); return
                USERS[current_username]["password"] = new_pass1
                save_users(USERS)
                st.session_state.user_info["password"] = new_pass1
                st.success("✅ Password changed successfully!"); st.balloons(); st.rerun()

def settings_management_panel():
    st.subheader("⚙️ System Settings — Categories, Departments & Roles")
    st.info("🛡️ Super Admin Only — Add, Edit, Delete Categories, Departments & Roles."); st.divider()
    cats_tab, dept_tab, roles_tab = st.tabs(["🏷️ Manage Categories", "🏢 Manage Departments", "🎖️ Manage Roles / Permissions"])
    with cats_tab:
        st.markdown("### 🏷️ Request Categories")
        st.info("These appear in the request form dropdown."); st.divider()
        current_cats = load_categories()
        with st.form("add_category_form", clear_on_submit=True):
            new_cat = st.text_input("➕ Add New Category", placeholder="e.g. Travel Allowance")
            if st.form_submit_button("✅ Add Category"):
                if new_cat.strip() and new_cat.strip() not in current_cats:
                    current_cats.append(new_cat.strip()); save_categories(current_cats)
                    log_action("CATEGORY_ADDED", new_data={"name": new_cat.strip()})
                    st.success(f"✅ Added: {new_cat}"); st.rerun()
                elif new_cat.strip() in current_cats: st.warning("⚠️ Category already exists!")
        st.divider()
        for i, cat in enumerate(current_cats):
            c1, c2, c3 = st.columns([4, 1, 1])
            with c1: st.markdown(f"• **{cat}**")
            with c2:
                if st.button(f"✏️ Edit", key=f"edit_cat_{i}"): st.session_state[f"editing_cat_{i}"] = True
            with c3:
                if len(current_cats) > 1 and st.button(f"🗑️ Delete", key=f"del_cat_{i}"):
                    current_cats.pop(i); save_categories(current_cats)
                    log_action("CATEGORY_DELETED", old_data={"name": cat})
                    st.success(f"✅ Deleted: {cat}"); st.rerun()
            if st.session_state.get(f"editing_cat_{i}", False):
                with st.form(f"save_cat_form_{i}", clear_on_submit=True):
                    renamed = st.text_input("Rename Category", value=cat)
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.form_submit_button("💾 Save"):
                            current_cats[i] = renamed.strip(); save_categories(current_cats)
                            log_action("CATEGORY_EDITED", old_data={"name": cat}, new_data={"name": renamed.strip()})
                            st.session_state[f"editing_cat_{i}"] = False
                            st.success(f"✅ Renamed to: {renamed}"); st.rerun()
                    with col2:
                        if st.form_submit_button("❌ Cancel"): st.session_state[f"editing_cat_{i}"] = False; st.rerun()
    with dept_tab:
        st.markdown("### 🏢 Manage Departments")
        st.info("Create, rename or remove departments."); st.divider()
        current_depts = load_departments()
        with st.form("add_dept_form", clear_on_submit=True):
            new_dept = st.text_input("➕ Add New Department", placeholder="e.g. HR, Maintenance")
            if st.form_submit_button("✅ Add Department"):
                if new_dept.strip() and new_dept.strip() not in current_depts:
                    current_depts.append(new_dept.strip()); save_departments(current_depts)
                    log_action("DEPARTMENT_ADDED", new_data={"name": new_dept.strip()})
                    st.success(f"✅ Added: {new_dept}"); st.rerun()
                elif new_dept.strip() in current_depts: st.warning("⚠️ Department already exists!")
        st.divider()
        for i, dept_name in enumerate(current_depts):
            c1, c2, c3 = st.columns([4, 1, 1])
            with c1: st.markdown(f"• **{dept_name}**")
            with c2:
                if st.button(f"✏️ Edit", key=f"edit_dept_{i}"): st.session_state[f"editing_dept_{i}"] = True
            with c3:
                if len(current_depts) > 1 and st.button(f"🗑️ Delete", key=f"del_dept_{i}"):
                    current_depts.pop(i); save_departments(current_depts)
                    log_action("DEPARTMENT_DELETED", old_data={"name": dept_name})
                    st.success(f"✅ Deleted: {dept_name}"); st.rerun()
            if st.session_state.get(f"editing_dept_{i}", False):
                with st.form(f"save_dept_form_{i}", clear_on_submit=True):
                    renamed = st.text_input("Rename Department", value=dept_name)
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.form_submit_button("💾 Save"):
                            current_depts[i] = renamed.strip(); save_departments(current_depts)
                            log_action("DEPARTMENT_EDITED", old_data={"name": dept_name}, new_data={"name": renamed.strip()})
                            st.session_state[f"editing_dept_{i}"] = False
                            st.success(f"✅ Renamed to: {renamed}"); st.rerun()
                    with col2:
                        if st.form_submit_button("❌ Cancel"): st.session_state[f"editing_dept_{i}"] = False; st.rerun()
    with roles_tab:
        st.markdown("### 🎖️ User Roles / Permission Levels")
        st.info("⚠️ 'Super Admin' cannot be deleted."); st.divider()
        current_roles = load_roles()
        with st.form("add_role_form", clear_on_submit=True):
            new_role = st.text_input("➕ Add New Role", placeholder="e.g. HR Manager")
            if st.form_submit_button("✅ Add Role"):
                if new_role.strip() and new_role.strip() not in current_roles:
                    current_roles.append(new_role.strip()); save_roles(current_roles)
                    default_perms = PERMISSION_DEFAULTS.get(new_role.strip(), {})
                    log_action("ROLE_ADDED", new_data={"role": new_role.strip(), "permissions": default_perms})
                    st.success(f"✅ Added: {new_role}"); st.rerun()
                elif new_role.strip() in current_roles: st.warning("⚠️ Role already exists!")
        st.divider()
        for i, role in enumerate(current_roles):
            c1, c2, c3 = st.columns([4, 1, 1])
            with c1: st.markdown(f"• **{role}**")
            with c2:
                if st.button(f"✏️ Edit", key=f"edit_role_{i}"): st.session_state[f"editing_role_{i}"] = True
            with c3:
                if role != "Super Admin" and len(current_roles) > 1 and st.button(f"🗑️ Delete", key=f"del_role_{i}"):
                    current_roles.pop(i); save_roles(current_roles)
                    log_action("ROLE_DELETED", old_data={"role": role})
                    st.success(f"✅ Deleted: {role}"); st.rerun()
            if st.session_state.get(f"editing_role_{i}", False):
                with st.form(f"save_role_form_{i}", clear_on_submit=True):
                    renamed = st.text_input("Rename Role", value=role)
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.form_submit_button("💾 Save"):
                            current_roles[i] = renamed.strip(); save_roles(current_roles)
                            log_action("ROLE_EDITED", old_data={"role": role}, new_data={"role": renamed.strip()})
                            st.session_state[f"editing_role_{i}"] = False
                            st.success(f"✅ Renamed to: {renamed}"); st.rerun()
                    with col2:
                        if st.form_submit_button("❌ Cancel"): st.session_state[f"editing_role_{i}"] = False; st.rerun()

def refresh_authenticated_user_session():
    """Refresh the logged-in user's permissions from the persistent user database.

    User permissions are stored in USER_DB_PATH, but Streamlit session state normally
    keeps the login snapshot. Refreshing it on each rerun makes permission changes
    take effect without requiring the user to log out and back in.
    """
    if not st.session_state.get("logged_in"):
        return
    current_username = str(st.session_state.get("user_info", {}).get("username", "")).strip().lower()
    if not current_username:
        return
    try:
        users = load_users(force=True)
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


def user_management_panel():
    st.subheader("👤 User Management — Create & Manage System Users")
    st.info("🛡️ Super Admin Only — Create, edit, or delete user accounts."); st.divider()
    # Read the latest user record so permission edits are never based on a stale cache.
    USERS = load_users(force=True); ROLES = load_roles()
    tab1, tab2, tab3 = st.tabs(["➕ Create New User", "✏️ Edit User", "🗑️ Delete User"])
    with tab1:
        st.markdown("### ➕ Create New System User")
        with st.form("create_user_form", border=True, clear_on_submit=True):
            new_full_name = st.text_input("👤 Full Name", placeholder="e.g. John Smith")
            new_username = st.text_input("🔐 Username", placeholder="e.g. john_smith").lower().strip()
            new_password = st.text_input("🔑 Password", type="password")
            new_role = st.selectbox("🎖️ Role / Permission Level", ROLES)
            st.markdown("### ✅ Extra Permissions")
            st.caption(f"Defaults for **{new_role}** are pre-selected")
            defaults = PERMISSION_DEFAULTS.get(new_role, PERMISSION_DEFAULTS["Staff"])
            col1, col2 = st.columns(2)
            perm_view_all = col1.checkbox(PERMISSION_LABELS["can_view_all_dept"], value=defaults.get("can_view_all_dept", False))
            perm_pdf = col1.checkbox(PERMISSION_LABELS["can_generate_pdf"], value=defaults.get("can_generate_pdf", False))
            perm_download = col2.checkbox(PERMISSION_LABELS["can_download_data"], value=defaults.get("can_download_data", False))
            perm_approve = col2.checkbox(PERMISSION_LABELS["can_approve_requests"], value=defaults.get("can_approve_requests", False))
            st.markdown("### 🧩 Module Access")
            mcol1, mcol2 = st.columns(2)
            perm_ad = mcol1.checkbox(PERMISSION_LABELS["can_access_addition_deduction"], value=defaults.get("can_access_addition_deduction", True))
            perm_wo = mcol2.checkbox(PERMISSION_LABELS["can_access_work_orders"], value=defaults.get("can_access_work_orders", False))
            perm_wo_total = st.checkbox(PERMISSION_LABELS["can_access_wo_total"], value=defaults.get("can_access_wo_total", False))
            perm_hr = st.checkbox(PERMISSION_LABELS["can_access_hr_leave"], value=defaults.get("can_access_hr_leave", False))
            perm_leave_req = st.checkbox(PERMISSION_LABELS["can_access_leave_request"], value=defaults.get("can_access_leave_request", False))
            perm_employee_reports = st.checkbox(PERMISSION_LABELS["can_access_employee_hr_reports"], value=defaults.get("can_access_employee_hr_reports", False))
            st.markdown("### 📊 HR / Report Access (View Only unless stated)")
            rcol1, rcol2 = st.columns(2)
            perm_holiday_calendar = rcol1.checkbox(PERMISSION_LABELS["can_access_holiday_calendar"], value=defaults.get("can_access_holiday_calendar", False))
            perm_hr_reports = rcol2.checkbox(PERMISSION_LABELS["can_access_hr_reports"], value=defaults.get("can_access_hr_reports", False))
            perm_employee_overview = rcol1.checkbox(PERMISSION_LABELS["can_access_employee_overview"], value=defaults.get("can_access_employee_overview", False))
            perm_holiday_calculator = rcol2.checkbox(PERMISSION_LABELS["can_access_holiday_calculator"], value=defaults.get("can_access_holiday_calculator", False))
            perm_employee_directory = rcol1.checkbox(PERMISSION_LABELS["can_access_employee_directory"], value=defaults.get("can_access_employee_directory", False))
            perm_store = st.checkbox(PERMISSION_LABELS["can_access_store_deduction"], value=defaults.get("can_access_store_deduction", False))
            perm_employee_items = st.checkbox(PERMISSION_LABELS["can_access_employee_items"], value=defaults.get("can_access_employee_items", False))
            perm_inspector = st.checkbox(PERMISSION_LABELS["can_access_inspector_bonus"], value=defaults.get("can_access_inspector_bonus", False))
            new_dept = st.selectbox("🏢 Department", load_departments())
            employee_options = [""] + [f"{e.get('emp_id')} — {e.get('name')}" for e in _hrp_load_employees()]
            selected_employee_link = st.selectbox("👤 Link to Employee (required for Employee HR Reports)", employee_options)
            new_employee_id = selected_employee_link.split(" — ", 1)[0].strip() if " — " in selected_employee_link else ""
            new_active = st.checkbox("✅ Account Active", value=True, help="Uncheck to block this user from logging in.")
            if st.form_submit_button("✅ Create User Account", type="primary"):
                if not new_full_name.strip() or not new_username or not new_password: st.error("❌ All fields required!")
                elif new_username in USERS: st.error(f"❌ Username '{new_username}' already exists!")
                else:
                    USERS[new_username] = {"full_name": new_full_name.strip(), "password": new_password, "role": new_role, "dept": new_dept, "can_view_all_dept": perm_view_all, "can_generate_pdf": perm_pdf, "can_download_data": perm_download, "can_approve_requests": perm_approve, "can_access_inspector_bonus": perm_inspector, "can_access_addition_deduction": perm_ad, "can_access_work_orders": perm_wo, "can_access_wo_total": perm_wo_total, "can_access_hr_leave": perm_hr, "can_access_leave_request": perm_leave_req, "can_access_employee_hr_reports": perm_employee_reports, "can_access_holiday_calendar": perm_holiday_calendar, "can_access_hr_reports": perm_hr_reports, "can_access_employee_overview": perm_employee_overview, "can_access_holiday_calculator": perm_holiday_calculator, "can_access_employee_directory": perm_employee_directory, "employee_id": new_employee_id, "can_access_store_deduction": perm_store, "can_access_employee_items": perm_employee_items, "is_active": new_active}
                    save_users(USERS)
                    log_action("USER_CREATED", new_data={"username": new_username, "full_name": new_full_name.strip(), "role": new_role, "department": new_dept, "is_active": new_active})
                    st.success(f"✅ User **'{new_full_name}'** created!"); st.balloons()
    with tab2:
        st.markdown("### ✏️ Edit User")
        edit_user_sel = st.selectbox("Select User to Edit", list(USERS.keys()), key="edit_user_selector")
        if edit_user_sel:
            curr = USERS[edit_user_sel]
            st.info(f"Current: **{curr.get('full_name', edit_user_sel)}** | {curr['role']} | {curr['dept']} | {'🟢 Active' if curr.get('is_active', True) else '🔴 Inactive'}")
            with st.form(f"edit_user_form_{edit_user_sel}", border=True, clear_on_submit=True):
                upd_full_name = st.text_input("👤 Full Name", value=curr.get("full_name", edit_user_sel))
                upd_username_new = st.text_input("🔐 Change Username", value=edit_user_sel).lower().strip()
                upd_password = st.text_input("🔑 New Password (leave blank to keep)", type="password")
                upd_role = st.selectbox("🎖️ Role", ROLES, index=ROLES.index(curr["role"]) if curr["role"] in ROLES else 0)
                dept_list = load_departments()
                upd_dept = st.selectbox("🏢 Department", dept_list, index=dept_list.index(curr["dept"]) if curr["dept"] in dept_list else 0)
                employee_options_edit = [""] + [f"{e.get('emp_id')} — {e.get('name')}" for e in _hrp_load_employees()]
                curr_emp_id = str(curr.get("employee_id", "")).strip()
                curr_emp_label = next((x for x in employee_options_edit if x.startswith(curr_emp_id + " — ")), "") if curr_emp_id else ""
                edit_employee_link = st.selectbox("👤 Link to Employee (required for Employee HR Reports)", employee_options_edit, index=employee_options_edit.index(curr_emp_label) if curr_emp_label in employee_options_edit else 0)
                edit_employee_id = edit_employee_link.split(" — ", 1)[0].strip() if " — " in edit_employee_link else ""
                st.markdown("### ✅ Update Permissions")
                curr_perm_view = bool(curr.get("can_view_all_dept", False))
                curr_perm_pdf = bool(curr.get("can_generate_pdf", False))
                curr_perm_dl = bool(curr.get("can_download_data", False))
                curr_perm_app = bool(curr.get("can_approve_requests", False))
                curr_perm_ib = bool(curr.get("can_access_inspector_bonus", False))
                curr_perm_ad = bool(curr.get("can_access_addition_deduction", False))
                curr_perm_wo = bool(curr.get("can_access_work_orders", False))
                curr_perm_wo_total = bool(curr.get("can_access_wo_total", False))
                curr_perm_hr = bool(curr.get("can_access_hr_leave", False))
                curr_perm_leave_req = bool(curr.get("can_access_leave_request", False))
                curr_perm_holiday_calendar = bool(curr.get("can_access_holiday_calendar", False))
                curr_perm_hr_reports = bool(curr.get("can_access_hr_reports", False))
                curr_perm_employee_overview = bool(curr.get("can_access_employee_overview", False))
                curr_perm_holiday_calculator = bool(curr.get("can_access_holiday_calculator", False))
                curr_perm_employee_directory = bool(curr.get("can_access_employee_directory", False))
                curr_perm_store = bool(curr.get("can_access_store_deduction", False))
                curr_perm_employee_items = bool(curr.get("can_access_employee_items", False))
                ecol1, ecol2 = st.columns(2)
                edit_view = ecol1.checkbox(PERMISSION_LABELS["can_view_all_dept"], value=curr_perm_view)
                edit_pdf = ecol1.checkbox(PERMISSION_LABELS["can_generate_pdf"], value=curr_perm_pdf)
                edit_dl = ecol2.checkbox(PERMISSION_LABELS["can_download_data"], value=curr_perm_dl)
                edit_app = ecol2.checkbox(PERMISSION_LABELS["can_approve_requests"], value=curr_perm_app)
                st.markdown("### 🧩 Module Access")
                mcol1, mcol2 = st.columns(2)
                edit_ad = mcol1.checkbox(PERMISSION_LABELS["can_access_addition_deduction"], value=curr_perm_ad)
                edit_wo = mcol2.checkbox(PERMISSION_LABELS["can_access_work_orders"], value=curr_perm_wo)
                edit_wo_total = st.checkbox(PERMISSION_LABELS["can_access_wo_total"], value=curr_perm_wo_total)
                edit_hr = st.checkbox(PERMISSION_LABELS["can_access_hr_leave"], value=curr_perm_hr)
                edit_leave_req = st.checkbox(PERMISSION_LABELS["can_access_leave_request"], value=curr_perm_leave_req)
                edit_employee_reports = st.checkbox(PERMISSION_LABELS["can_access_employee_hr_reports"], value=bool(curr.get("can_access_employee_hr_reports", False)))
                st.markdown("### 📊 HR / Report Access (View Only unless stated)")
                ercol1, ercol2 = st.columns(2)
                edit_holiday_calendar = ercol1.checkbox(PERMISSION_LABELS["can_access_holiday_calendar"], value=curr_perm_holiday_calendar)
                edit_hr_reports = ercol2.checkbox(PERMISSION_LABELS["can_access_hr_reports"], value=curr_perm_hr_reports)
                edit_employee_overview = ercol1.checkbox(PERMISSION_LABELS["can_access_employee_overview"], value=curr_perm_employee_overview)
                edit_holiday_calculator = ercol2.checkbox(PERMISSION_LABELS["can_access_holiday_calculator"], value=curr_perm_holiday_calculator)
                edit_employee_directory = ercol1.checkbox(PERMISSION_LABELS["can_access_employee_directory"], value=curr_perm_employee_directory)
                edit_store = st.checkbox(PERMISSION_LABELS["can_access_store_deduction"], value=curr_perm_store)
                edit_employee_items = st.checkbox(PERMISSION_LABELS["can_access_employee_items"], value=curr_perm_employee_items)
                edit_ib = st.checkbox(PERMISSION_LABELS["can_access_inspector_bonus"], value=curr_perm_ib)
                edit_active = st.checkbox("✅ Account Active", value=curr.get("is_active", True), help="Uncheck to block this user from logging in.")
                if st.form_submit_button("🔄 Update User", type="primary"):
                    USERS = load_users()
                    if upd_username_new != edit_user_sel:
                        if upd_username_new in USERS: st.error(f"❌ Username '{upd_username_new}' already exists!"); return
                        USERS[upd_username_new] = {"full_name": upd_full_name.strip(), "password": upd_password if upd_password else curr["password"], "role": upd_role, "dept": upd_dept, "can_view_all_dept": edit_view, "can_generate_pdf": edit_pdf, "can_download_data": edit_dl, "can_approve_requests": edit_app, "can_access_inspector_bonus": edit_ib, "can_access_addition_deduction": edit_ad, "can_access_work_orders": edit_wo, "can_access_wo_total": edit_wo_total, "can_access_hr_leave": edit_hr, "can_access_leave_request": edit_leave_req, "can_access_employee_hr_reports": edit_employee_reports, "can_access_holiday_calendar": edit_holiday_calendar, "can_access_hr_reports": edit_hr_reports, "can_access_employee_overview": edit_employee_overview, "can_access_holiday_calculator": edit_holiday_calculator, "can_access_employee_directory": edit_employee_directory, "employee_id": edit_employee_id, "can_access_store_deduction": edit_store, "can_access_employee_items": edit_employee_items, "is_active": edit_active}
                        del USERS[edit_user_sel]
                    else:
                        USERS[edit_user_sel]["full_name"] = upd_full_name.strip()
                        if upd_password: USERS[edit_user_sel]["password"] = upd_password
                        USERS[edit_user_sel]["role"] = upd_role
                        USERS[edit_user_sel]["dept"] = upd_dept
                        USERS[edit_user_sel]["can_view_all_dept"] = edit_view
                        USERS[edit_user_sel]["can_generate_pdf"] = edit_pdf
                        USERS[edit_user_sel]["can_download_data"] = edit_dl
                        USERS[edit_user_sel]["can_approve_requests"] = edit_app
                        USERS[edit_user_sel]["can_access_inspector_bonus"] = edit_ib
                        USERS[edit_user_sel]["can_access_addition_deduction"] = edit_ad
                        USERS[edit_user_sel]["can_access_work_orders"] = edit_wo
                        USERS[edit_user_sel]["can_access_wo_total"] = edit_wo_total
                        USERS[edit_user_sel]["can_access_hr_leave"] = edit_hr
                        USERS[edit_user_sel]["can_access_leave_request"] = edit_leave_req
                        USERS[edit_user_sel]["can_access_employee_hr_reports"] = edit_employee_reports
                        USERS[edit_user_sel]["can_access_holiday_calendar"] = edit_holiday_calendar
                        USERS[edit_user_sel]["can_access_hr_reports"] = edit_hr_reports
                        USERS[edit_user_sel]["can_access_employee_overview"] = edit_employee_overview
                        USERS[edit_user_sel]["can_access_holiday_calculator"] = edit_holiday_calculator
                        USERS[edit_user_sel]["can_access_employee_directory"] = edit_employee_directory
                        USERS[edit_user_sel]["employee_id"] = edit_employee_id
                        USERS[edit_user_sel]["can_access_store_deduction"] = edit_store
                        USERS[edit_user_sel]["can_access_employee_items"] = edit_employee_items
                        USERS[edit_user_sel]["is_active"] = edit_active
                    save_users(USERS)
                    log_action("USER_EDITED", old_data=curr, new_data={"full_name": upd_full_name.strip(), "username": upd_username_new, "role": upd_role, "department": upd_dept, "is_active": edit_active})
                    st.success(f"✅ User updated: **{upd_full_name}**"); st.rerun()
    with tab3:
        st.markdown("### ⚠️ Delete User Account")
        st.warning("Existing requests remain safe.")
        del_user_sel = st.selectbox("Select User to DELETE", [u for u in USERS.keys() if u != st.session_state.user_info["username"]])
        if del_user_sel:
            del_name = USERS[del_user_sel].get("full_name", del_user_sel)
            if st.button(f"🗑️ DELETE: {del_name} ({del_user_sel})", type="secondary"):
                del USERS[del_user_sel]
                save_users(USERS)
                log_action("USER_DELETED", old_data={"username": del_user_sel, "full_name": del_name})
                st.success(f"✅ User **{del_name}** deleted!"); st.rerun()

# ============================================================
# 🔐 LOGIN
# ============================================================
if not st.session_state.logged_in:
    display_company_header()
    with st.form("login_form", border=True):
        st.markdown("### 🔒 Secure Gateway Login")
        st.caption("Enter your credentials to access the system"); st.divider()
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
            else: st.error("❌ Invalid Username or Password. Please try again.")
    st.stop()

# Always refresh the authenticated user's role/permissions from USER_DB_PATH before
# rendering the role-specific portal. This makes admin permission changes effective
# on the next app rerun instead of only after a fresh login.
refresh_authenticated_user_session()

col_left, col_right = st.columns([4, 1])
with col_left: refresh_data_button()
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

all_live_requests = load_records_from_excel()
CATEGORIES = load_categories()

st.divider()


# ============================================================
# 👤 EMPLOYEE HR REPORTS — VIEW ONLY
# ============================================================
def render_employee_hr_reports(current_user_info):
    """Read-only HR dashboard restricted to the Employee ID linked to the login."""
    _hr_portal_init()
    linked_id = str((current_user_info or {}).get("employee_id", "")).strip()
    if not linked_id:
        st.subheader("👤 My HR Reports")
        st.error("❌ Your account is not linked to an Employee ID. Please contact Super Admin.")
        return

    employees = st.session_state.get("hrp_employees")
    if employees is None:
        employees = _hrp_load_employees()
        st.session_state.hrp_employees = employees
    employees = _hrp_normalize_employee_records(employees)
    st.session_state.hrp_employees = employees
    leave_records = st.session_state.get("hrp_leave_records")
    if leave_records is None:
        leave_records = _hrp_load_leave_records()
        st.session_state.hrp_leave_records = leave_records
    employee = next((e for e in employees if str(e.get("emp_id", "")).strip().casefold() == linked_id.casefold()), None)
    if not employee:
        st.subheader("👤 My HR Reports")
        st.error(f"❌ Employee ID **{linked_id}** could not be found in the HR employee records.")
        return

    st.subheader("👤 My HR Reports")
    st.caption("View only — you can see your own employee details, holiday position, absence records and calendar.")

    entitlement_result = _hrp_get_employee_entitlement(employee)
    entitlement = float(entitlement_result[0]) if isinstance(entitlement_result, tuple) else float(entitlement_result)
    entitlement_note = str(entitlement_result[1]) if isinstance(entitlement_result, tuple) and len(entitlement_result) > 1 else ""
    approved_holiday = _hrp_get_approved_holiday_days(linked_id)

    today = date.today()
    try:
        employee_start = employee.get("start_date")
        if not isinstance(employee_start, date):
            employee_start = pd.to_datetime(employee_start).date()
    except Exception:
        employee_start = today
    passed_bank_holiday_end = today
    try:
        leaving_date = employee.get("leaving_date")
        if leaving_date:
            if not isinstance(leaving_date, date):
                leaving_date = pd.to_datetime(leaving_date).date()
            passed_bank_holiday_end = min(today, leaving_date)
    except Exception:
        pass

    passed_bank_holiday_days = _hrp_get_bank_holiday_days(
        employee, employee_start, passed_bank_holiday_end
    )
    company_closure_days = _hrp_get_company_closure_holiday_days(
        employee, employee_start, passed_bank_holiday_end if passed_bank_holiday_end >= employee_start else employee_start
    )
    company_closure_future = _hrp_get_company_closure_holiday_days(
        employee,
        max(today + timedelta(days=1), employee_start),
        date(today.year, 12, 31),
    )
    company_closure_reserved = round(company_closure_days + company_closure_future, 1)
    approved_holiday_with_closure = round(approved_holiday + company_closure_reserved, 1)
    pure_balance = entitlement - approved_holiday_with_closure

    upcoming_bank_holidays = _hrp_get_upcoming_bank_holidays(today + timedelta(days=1), limit=5)
    upcoming_bank_holiday_days = _hrp_get_upcoming_bank_holiday_days_for_employee(employee, today + timedelta(days=1))
    bank_holidays_reserved = round(passed_bank_holiday_days + upcoming_bank_holiday_days, 1)
    balance = round(pure_balance - bank_holidays_reserved, 1)

    employee_is_left = (
        str(employee.get("status", "")).strip().casefold() == "left"
        or bool(employee.get("leaving_date"))
    )
    if employee_is_left:
        entitlement = 0.0
        approved_holiday = 0.0
        company_closure_reserved = 0.0
        approved_holiday_with_closure = 0.0
        passed_bank_holiday_days = 0.0
        upcoming_bank_holiday_days = 0.0
        bank_holidays_reserved = 0.0
        balance = 0.0
        upcoming_bank_holidays = []

    d1, d2, d3, d4, d5, d6 = st.columns(6)
    d1.metric("Holiday Entitlement", f"{entitlement:g} days")
    d2.metric("Holiday Used / Pre-booked", f"{approved_holiday_with_closure:g} days")
    d3.metric("Bank Holidays Passed", f"{passed_bank_holiday_days:g} days")
    d4.metric("Holiday Balance", f"{balance:g} days")
    if upcoming_bank_holidays:
        next_bank_date, next_bank_name = upcoming_bank_holidays[0]
        d5.metric("Upcoming Bank Holiday", next_bank_date.strftime("%d %b %Y"))
        d5.caption(next_bank_name)
    else:
        d5.metric("Upcoming Bank Holiday", "None")
    d6.metric("Department", employee.get("department", ""))

    if employee_is_left:
        final_records = _hrp_get_final_settlement_records(linked_id)
        active_final = [
            r for r in final_records
            if str(r.get("status", "")).strip().casefold() in {"pending", "approved"}
        ]
        if active_final:
            latest_final = active_final[0]
            final_status = str(latest_final.get("status", "")).strip().title()
            final_type = str(latest_final.get("type", "")).strip()
            final_days = float(latest_final.get("days", 0) or 0)
            final_amount = float(latest_final.get("amount", 0) or 0)
            if final_status == "Approved":
                st.success(
                    f"💷 Final holiday settlement: **{final_type} — {final_days:g} days (£{final_amount:,.2f}) — Approved**. "
                    "The employee's live holiday balance remains **0 days**."
                )
            else:
                st.info(
                    f"⏳ Final holiday settlement: **{final_type} — {final_days:g} days (£{final_amount:,.2f}) — Pending Director approval**. "
                    "The employee's live holiday balance remains **0 days**."
                )

    st.caption(
        f"🏦 Bank holidays passed since your start date: **{passed_bank_holiday_days:g} days** · "
        f"Bank holidays still to come: **{upcoming_bank_holiday_days:g} days** · "
        f"Bank holidays reserved/deducted from balance: **{bank_holidays_reserved:g} days** · "
        f"Pre-booked company closure (29–31 Dec): **{company_closure_reserved:g} days** · "
        f"Holiday used including pre-booked closure: **{approved_holiday_with_closure:g} days** · "
        f"Current available holiday balance: **{balance:g} days**"
    )
    if upcoming_bank_holidays:
        with st.expander("📅 Upcoming Bank Holidays", expanded=False):
            st.dataframe(
                pd.DataFrame([
                    {"Date": bank_date.strftime("%d/%m/%Y"), "Bank Holiday": bank_name}
                    for bank_date, bank_name in upcoming_bank_holidays
                ]),
                width="stretch",
                hide_index=True,
            )
    if "HR-adjusted" in entitlement_note:
        st.caption(f"⚙️ HR-adjusted holiday entitlement: **{entitlement:g} days**")
        adjustment_note = str(employee.get("adjustment_note", "")).strip()
        if adjustment_note:
            st.caption(f"HR adjustment reason: {adjustment_note}")

    st.markdown("### 📋 My Employee Details")
    details = pd.DataFrame([{
        "Employee ID": employee.get("emp_id", ""), "Full Name": employee.get("name", ""),
        "Department": employee.get("department", ""), "Position": employee.get("job_title", ""),
        "Start Date": employee.get("start_date", ""), "Agreement": employee.get("agreement_type", ""),
        "Status": employee.get("status", ""), "Working Pattern": employee.get("working_pattern", ""),
        "Days Worked / Week": employee.get("days_per_week", 5),
    }])
    st.dataframe(details, width="stretch", hide_index=True)

    st.markdown("### 📅 My Holiday & Absence Calendar")
    year = st.number_input("Year", min_value=2020, max_value=2100, value=date.today().year, step=1, key="employee_hr_report_year")
    first = date(int(year), 1, 1); last = date(int(year), 12, 31)
    dates = []
    cur = first
    while cur <= last:
        dates.append(cur); cur += timedelta(days=1)
    own_records = [r for r in leave_records if str(r.get("employee_id", "")).strip().casefold() == linked_id.casefold()
                   and str(r.get("status", "")).strip().casefold() != "rejected"]
    lookup = {}
    for r in own_records:
        try:
            a = r.get("date_from"); b = r.get("date_to")
            if isinstance(a, str): a = pd.to_datetime(a).date()
            if isinstance(b, str): b = pd.to_datetime(b).date()
            code = _hrp_calendar_leave_code(r.get("type", ""), r.get("status", "Approved"))
        except Exception:
            continue
        cur = a
        while cur <= b:
            if cur.weekday() < 5 and cur not in _hrp_non_working_dates(cur.year):
                existing = lookup.get(cur, "")
                if code.rstrip("*") == "S":
                    lookup[cur] = code
                elif existing.rstrip("*") == "S":
                    pass
                elif not existing:
                    lookup[cur] = code
                elif existing.endswith("*") and not code.endswith("*"):
                    lookup[cur] = code
                elif code not in existing.split("/"):
                    lookup[cur] = f"{existing}/{code}"
            cur += timedelta(days=1)
    emp_start = _hrp_normalize_employee_date(employee.get("start_date"), None)
    emp_leaving = _hrp_normalize_employee_date(employee.get("leaving_date"), None)
    row = {}
    for d in dates:
        col = d.strftime("%d %b")
        if emp_start and d < emp_start:
            row[col] = "NA"
        elif emp_leaving is not None and d > emp_leaving:
            row[col] = "LEFT"
        else:
            nonwork_code, _ = _hrp_non_working_reason(d)
            row[col] = lookup.get(d, nonwork_code or "")
    cal_df = pd.DataFrame([row])
    def style_my_calendar(dataframe):
        styles = pd.DataFrame("", index=dataframe.index, columns=dataframe.columns)
        for c in dataframe.columns:
            code = str(dataframe.at[0, c] or "").rstrip("*")
            bg = HRP_CALENDAR_COLOURS.get(code)
            if bg:
                styles.at[0, c] = f"background-color: {bg}; color: #000000; font-weight: 800; text-align: center; vertical-align: middle;"
        return styles
    calendar_style = (
        cal_df.style
        .apply(style_my_calendar, axis=None)
        .set_properties(**{"text-align": "center", "vertical-align": "middle"})
        .set_table_styles([
            {"selector": "th", "props": [("text-align", "center"), ("vertical-align", "middle")]},
            {"selector": "td", "props": [("text-align", "center"), ("vertical-align", "middle")]},
        ])
    )
    st.dataframe(calendar_style, width="stretch", hide_index=True)
    st.caption("H Holiday · HD Half Day · BH Bank Holiday · UH Unpaid Holiday · C College · NA Closed / Not yet started · T Training · M Maternity · UA Unpaid Absence · S Sick · FE Family/Emergency · P Paternity · O Other · LEFT Employee Left · * Pending")

    st.markdown("### 📝 My Leave & Absence Records")
    visible = []
    for r in sorted(own_records, key=lambda x: x.get("date_from", date.min), reverse=True):
        visible.append({
            "Leave ID": r.get("leave_id", ""), "Date From": r.get("date_from", ""), "Date To": r.get("date_to", ""),
            "Leave Type": r.get("type", ""), "Days": r.get("days", 0), "Status": r.get("status", ""),
            "Request Source": r.get("request_source", ""), "Notes": r.get("notes", ""),
        })
    if visible:
        st.dataframe(pd.DataFrame(visible), width="stretch", hide_index=True)
    else:
        st.info("No holiday or absence records have been recorded for you yet.")

if user_info.get("can_access_employee_hr_reports", False) and role not in ["Manager", "Staff", "Team Member"]:
    render_employee_hr_reports(user_info)


# ============================================================
# 📊 DIRECTOR — GRANULAR HR / REPORT ACCESS
# ============================================================
def render_director_hr_access_portal(current_user_info):
    """Render only the read-only HR/report modules explicitly granted by Super Admin."""
    try:
        _hr_portal_init()
    except Exception as exc:
        print(f"Director HR access initialisation warning: {exc}")
    permissions = {
        "holiday_calendar": bool(current_user_info.get("can_access_holiday_calendar", False)),
        "hr_reports": bool(current_user_info.get("can_access_hr_reports", False)),
        "employee_overview": bool(current_user_info.get("can_access_employee_overview", False)),
        "holiday_calculator": bool(current_user_info.get("can_access_holiday_calculator", False)),
        "employee_directory": bool(current_user_info.get("can_access_employee_directory", False)),
    }
    labels = {
        "holiday_calendar": "📅 Holiday Calendar",
        "hr_reports": "📥 HR Reports",
        "employee_overview": "📊 Employee Overview",
        "holiday_calculator": "🧮 Holiday Calculator",
        "employee_directory": "👥 Employee Directory",
    }
    enabled = [key for key, allowed in permissions.items() if allowed]
    if not enabled:
        st.info("No HR/report access has been granted to this Director. A Super Admin can enable individual modules in User Management → Edit User.")
        return

    st.subheader("📊 HR & Reports")
    st.caption("Read-only access granted by Super Admin. Each module is controlled independently.")
    tabs = st.tabs([labels[key] for key in enabled])
    for key, tab in zip(enabled, tabs):
        with tab:
            if key == "holiday_calendar":
                _hrp_render_holiday_calendar()
            elif key == "hr_reports":
                render_hr_download_reports()
            elif key == "employee_overview":
                employees = _hrp_load_employees()
                if not employees:
                    st.info("No employees are currently recorded.")
                else:
                    rows = []
                    for e in employees:
                        pos = _hrp_get_holiday_position(e.get("emp_id", ""))
                        rows.append({
                            "Employee ID": e.get("emp_id", ""),
                            "Name": e.get("name", ""),
                            "Department": e.get("department", ""),
                            "Job Title": e.get("job_title", ""),
                            "Status": e.get("status", "Active"),
                            "Start Date": e.get("start_date", ""),
                            "Holiday Entitlement": pos.get("entitlement", 0.0),
                            "Holiday Used / Booked": pos.get("used", 0.0),
                            "Bank Holidays Passed": pos.get("bank_holidays_passed", 0.0),
                            "Upcoming Bank Holidays": pos.get("upcoming_bank_holidays", 0.0),
                            "Holiday Balance": pos.get("balance", 0.0),
                        })
                    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
                    st.caption(f"Total employees: {len(rows)}")
            elif key == "holiday_calculator":
                employees = _hrp_load_employees()
                if not employees:
                    st.info("No employees are currently recorded.")
                else:
                    rows = []
                    for e in employees:
                        pos = _hrp_get_holiday_position(e.get("emp_id", ""))
                        rows.append({
                            "Employee ID": e.get("emp_id", ""),
                            "Employee": e.get("name", ""),
                            "Department": e.get("department", ""),
                            "Entitlement": pos.get("entitlement", 0.0),
                            "Booked / Used": pos.get("used", 0.0),
                            "Bank Holidays Passed": pos.get("bank_holidays_passed", 0.0),
                            "Upcoming Bank Holidays": pos.get("upcoming_bank_holidays", 0.0),
                            "Remaining": pos.get("balance", 0.0),
                        })
                    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
                    st.info("Remaining is calculated from entitlement minus booked/used holiday, passed bank holidays and upcoming bank holidays, using the same live HR calculation as Employee Management.")
            elif key == "employee_directory":
                employees = _hrp_load_employees()
                if not employees:
                    st.info("No employees are currently recorded.")
                else:
                    rows = [{
                        "Employee ID": e.get("emp_id", ""),
                        "Full Name": e.get("name", ""),
                        "Department": e.get("department", ""),
                        "Position / Job Title": e.get("job_title", ""),
                        "Agreement Type": e.get("agreement_type", ""),
                        "Status": e.get("status", "Active"),
                        "Start Date": e.get("start_date", ""),
                        "Working Pattern": e.get("working_pattern", ""),
                        "Days Worked / Week": e.get("days_per_week", 5),
                    } for e in employees]
                    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
                    st.caption(f"Total employees: {len(rows)}")


# ============================================================
# 📋 ROLE-BASED PORTALS
# ============================================================
if role == "Employee":
    if user_info.get("can_access_employee_items", False):
        st.subheader(f"🧰 My Employee Items — {full_name}")
        itab1, itab2 = st.tabs(["📊 My Holdings", "📚 My Check-in History"])
        with itab1:
            linked_id = str(user_info.get("employee_id", "")).strip()
            all_holdings = load_employee_items(force=True)
            own = [r for r in all_holdings if str(r.get("employee_id", "")).strip() == linked_id and float(r.get("qty_outstanding",0) or 0) > 0]
            if own:
                st.dataframe(pd.DataFrame([{"Employee":r.get("emp_name"),"Department":r.get("emp_dept"),"Item":r.get("item_name"),"Outstanding":r.get("qty_outstanding"),"Unit £":r.get("unit_price"),"Issue Date":r.get("issue_date")} for r in own]), width="stretch", hide_index=True)
                pdf=_build_item_holdings_pdf(own, f"My Employee Item Holdings - {full_name}")
                st.download_button("📄 Download My Holdings PDF", data=pdf or b"", file_name="My_Employee_Item_Holdings.pdf", mime="application/pdf", disabled=not bool(pdf), key="employee_my_holdings_pdf")
            else:
                st.info("No outstanding company items are currently recorded against your linked HR employee account.")
        with itab2:
            linked_id = str(user_info.get("employee_id", "")).strip()
            mine=[c for c in load_item_checkins(force=True) if str(c.get("employee_id","")).strip()==linked_id]
            if not mine: st.info("No check-in history found.")
            for c in reversed(mine):
                st.write(f"#{c.get('id')} | {c.get('checkin_date')} | Deduction £{float(c.get('total_deduction',0) or 0):.2f}")
                pdf=_build_item_checkin_pdf(c, f"My Leaver Item Check-in #{c.get('id')}")
                st.download_button("📄 PDF", data=pdf or b"", file_name=f"My_Checkin_{c.get('id')}.pdf", mime="application/pdf", disabled=not bool(pdf), key=f"employee_checkin_pdf_{c.get('id')}")

elif role == "Work Order Employee":
    render_work_order_employee_portal(full_name, dept)

elif role == "Payroll":
    st.subheader("🧾 Payroll Portal")
    st.info("✅ View all requests and Download PDFs.")
    st.divider()
    tab_add_ded, tab_hr_leave, tab_store_ded, tab_store_ret, tab_leaver, tab_work_orders, tab_inspector_bonus = st.tabs([
        "➕ Addition & Deduction",
        "👥 HR Leave Settlement",
        "📦 Store Deductions",
        "📦 Store Returns (Additions)",
        "🧾 Leaver Clearance",
        "🛠️ Work Orders",
        "💰 National Grid Inspector Bonus"
    ])
    with tab_add_ded:
        tab_pending, tab_approved, tab_rejected = st.tabs(["⏳ Pending Requests", "✅ Approved Requests", "❌ Rejected Requests"])
        with tab_pending:
            pending = [r for r in all_live_requests if r.get("status") == "pending"]
            if not pending: st.success("✅ No pending requests!")
            else:
                st.metric("⏳ Pending", len(pending)); st.divider()
                for req in reversed(pending):
                    with st.expander(f"🟡 ID #{req.get('id')} | {req.get('emp_name')} | £{float(req.get('amount',0)):.2f} | {req.get('dept')}"):
                        st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept')}")
                        st.write(f"🔄 Type: {req.get('type')} | 🏷️ Category: {req.get('category')}")
                        st.write(f"💷 Amount: £{float(req.get('amount',0)):.2f}")
                        st.write(f"👔 **Line Manager:** {req.get('manager')}")
                        submitted_by = get_submitted_by(req)
                        if submitted_by: st.write(f"📝 **Submitted by:** {submitted_by}")
                        st.write(f"📅 **Date:** {req.get('date')}")
                        st.info(f"📝 **Description:** {req.get('desc')}")
                        display_attachments(req)
                        st.divider(); display_pdf_button(req, can_generate=True)
        with tab_approved:
            approved = [r for r in all_live_requests if r.get("status") == "approved"]
            if not approved: st.info("📋 No approved requests.")
            else:
                st.metric("✅ Approved", len(approved)); st.divider()
                for req in reversed(approved):
                    dec_by = req.get('decision_by', 'Director')
                    with st.expander(f"🟢 ID #{req.get('id')} | {req.get('emp_name')} | £{float(req.get('amount',0)):.2f} | ✅ {dec_by}"):
                        st.write(f"👤 Employee: {req.get('emp_name')}")
                        submitted_by = get_submitted_by(req)
                        if submitted_by: st.write(f"📝 **Submitted by:** {submitted_by}")
                        st.write(f"💷 Amount: £{float(req.get('amount',0)):.2f}")
                        st.info(f"💬 Comments: {req.get('director_comments', 'None')}")
                        display_attachments(req)
                        st.divider(); display_pdf_button(req, can_generate=True)
        with tab_rejected:
            rejected = [r for r in all_live_requests if r.get("status") == "rejected"]
            if not rejected: st.success("✅ No rejected requests!")
            else:
                st.metric("❌ Rejected", len(rejected)); st.divider()
                for req in reversed(rejected):
                    with st.expander(f"🔴 ID #{req.get('id')} | {req.get('emp_name')} | £{float(req.get('amount',0)):.2f}"):
                        st.write(f"👤 Employee: {req.get('emp_name')}")
                        st.error(f"❌ Rejected By: {req.get('decision_by', '—')} on {format_date(req.get('decision_date', ''))}")
                        st.error(f"💬 Reason: {req.get('director_comments', 'None')}")
                        display_attachments(req)
                        st.divider(); display_pdf_button(req, can_generate=True)
    with tab_hr_leave: render_hr_leave_payroll_portal()
    with tab_store_ded: render_store_payroll_portal(type_filter="Deduction")
    with tab_store_ret: render_store_payroll_portal(type_filter="Addition")
    with tab_leaver: render_leaver_clearance_payroll(full_name)
    with tab_work_orders: render_work_order_payroll_portal(full_name)
    with tab_inspector_bonus: render_inspector_bonus_payroll_portal(full_name)

elif role == "Work Order Manager":
    dept_name = dept
    request_tab, work_order_tab = st.tabs(["➕ Addition & Deduction", "🛠️ Work Orders"])
    with request_tab:
        dept_name = dept
        if st.session_state.get("editing_request_id"):
            eid = st.session_state.editing_request_id
            rec = next((r for r in all_live_requests if int(r.get("id", 0)) == int(eid)), None)
            if rec:
                st.subheader(f"✏️ Edit Request #{eid}")
                show_old_new_comparison("{}", rec)
                st.markdown("### 📎 Manage Attachments")
                att_name_raw = rec.get("attachment_name", "None")
                existing_files = []
                if att_name_raw and str(att_name_raw).strip().lower() != "none":
                    existing_files = [n.strip() for n in str(att_name_raw).split(",") if n.strip()]
                files_to_keep = []; files_to_remove = []
                if existing_files:
                    st.info(f"📋 **{len(existing_files)} attachment(s) currently attached:**")
                    for fname in existing_files:
                        file_path = os.path.join(UPLOAD_DIR, fname)
                        safe_key = f"keep_{eid}_{fname.replace(' ','_').replace('.','_')}"
                        col_check, col_name, col_dl = st.columns([1, 5, 2])
                        keep = col_check.checkbox("✅ Keep", value=True, key=safe_key)
                        col_name.markdown(f"📄 `{fname}`")
                        if os.path.exists(file_path):
                            with open(file_path, "rb") as f: col_dl.download_button("⬇️", f.read(), file_name=fname, key=f"dl_{safe_key}")
                        else: col_dl.caption("⚠️ Missing")
                        if keep: files_to_keep.append(fname)
                        else: files_to_remove.append(fname)
                    if files_to_remove: st.warning(f"🗑️ Will remove: {', '.join(files_to_remove)}")
                else: st.info("📋 No attachments currently attached.")
                st.markdown("#### ➕ Attach New Files")
                new_files_upload = st.file_uploader("Upload additional files", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key=f"new_upload_{eid}")
                with st.form("edit_form"):
                    c1, c2 = st.columns(2)
                    with c1:
                        en = st.text_input("👤 Employee Name", rec.get("emp_name", ""))
                        rt = st.selectbox("🔄 Transaction Type", ["Addition", "Deduction"], index=["Addition", "Deduction"].index(rec.get("type", "Addition")))
                        cat_idx = CATEGORIES.index(rec.get("category")) if rec.get("category") in CATEGORIES else 0
                        ct = st.selectbox("🏷️ Category / Reason", CATEGORIES, index=cat_idx)
                        amt = st.number_input("💷 Amount (£)", min_value=0.01, step=10.0, value=float(rec.get("amount", 0.01)))
                    with c2:
                        from datetime import datetime as dt
                        try: d = dt.strptime(str(rec.get("date", ""))[:10], "%Y-%m-%d")
                        except: d = dt.today()
                        dt_val = st.date_input("📅 Date", d)
                        mgr = st.text_input("👔 Line Manager", rec.get("manager", ""))
                        desc = st.text_area("📝 Description / Justification", rec.get("desc", ""))
                    if st.form_submit_button("✅ Submit Edit", type="primary"):
                        final_attachments = list(files_to_keep)
                        if new_files_upload:
                            for idx, f in enumerate(new_files_upload, start=len(final_attachments)+1):
                                fn = f"ID_{eid}_EDIT_F{idx}_{f.name}"
                                edit_path = os.path.join(UPLOAD_DIR, fn)
                                with open(edit_path, "wb") as outfile: outfile.write(f.getbuffer())
                                _upload_to_drive_bg(edit_path, fn)
                                final_attachments.append(fn)
                        records = load_records_from_excel()
                        old_data_dict = {"emp_name": rec.get("emp_name"), "dept": rec.get("dept"), "type": rec.get("type"), "category": rec.get("category"), "date": rec.get("date"), "amount": rec.get("amount"), "manager": rec.get("manager"), "desc": rec.get("desc")}
                        new_data_dict = {"emp_name": en.strip(), "dept": dept_name, "type": rt, "category": ct, "date": str(dt_val), "amount": amt, "manager": mgr.strip(), "desc": desc.strip()}
                        for r in records:
                            if int(r.get("id", 0)) == int(eid):
                                r["emp_name"] = en.strip(); r["type"] = rt; r["category"] = ct; r["amount"] = amt; r["date"] = str(dt_val); r["manager"] = mgr.strip(); r["desc"] = desc.strip(); r["status"] = "pending"; r["attachment_name"] = ", ".join(final_attachments) or "None"; r["old_data"] = json.dumps(old_data_dict, default=str)
                                break
                        log_action("EDITED", eid, old_data=old_data_dict, new_data=new_data_dict)
                        save_all_records(records)
                        st.success(f"✅ Updated!"); st.session_state.editing_request_id = None; st.rerun()
                if st.button("❌ Cancel", key=f"cancel_edit_{eid}"):
                    st.session_state.editing_request_id = None; st.rerun()
        else:
            st.subheader(f"➕ New Request — {dept_name}")
            nid = get_next_id(all_live_requests)
            form_version = st.session_state.get("wo_mgr_new_req_form_version", 0)
            # Keep Employee Source outside the form so switching between
            # Registered Employee and Manual Employee immediately reruns the
            # app and reveals the appropriate employee input.
            employee_source = st.radio(
                "👤 Employee Source",
                ["Registered Employee", "Manual Employee"],
                horizontal=True,
                key=f"wo_mgr_employee_source_v{form_version}",
            )
            with st.form(f"wo_mgr_new_req_v{form_version}", clear_on_submit=False):
                c1, c2 = st.columns(2)
                with c1:
                    registered_employee = None
                    if employee_source == "Registered Employee":
                        registered = [
                            e for e in _hrp_load_employees()
                            if str(e.get("status", "Active")).strip().casefold() == "active"
                            and str(e.get("name", "")).strip()
                        ]
                        dept_filtered = [
                            e for e in registered
                            if str(e.get("department", "")).strip().casefold() == str(dept_name).strip().casefold()
                        ]
                        if dept_filtered:
                            registered = dept_filtered
                        registered = sorted(registered, key=lambda e: (str(e.get("name", "")).casefold(), str(e.get("emp_id", "")).casefold()))
                        if registered:
                            labels = [f"{e.get('name','')} — {e.get('emp_id','')} · {e.get('department','')}" for e in registered]
                            selected_label = st.selectbox("Registered Employee", labels, key=f"wo_mgr_registered_emp_v{form_version}")
                            registered_employee = {label: emp for label, emp in zip(labels, registered)}.get(selected_label)
                            en = str((registered_employee or {}).get("name", "")).strip()
                            request_dept = str((registered_employee or {}).get("department", dept_name)).strip() or dept_name
                            if registered_employee:
                                st.caption(f"Employee ID: {registered_employee.get('emp_id', '-')} · Department: {request_dept}")
                        else:
                            st.warning("No active registered employees found for this department. You can use Manual Employee instead.")
                            en = ""
                            request_dept = dept_name
                    else:
                        en = st.text_input("👤 Employee Name", key=f"wo_mgr_en_v{form_version}")
                        request_dept = dept_name
                    rt = st.selectbox("🔄 Transaction Type", ["Addition", "Deduction"], key=f"wo_mgr_rt_v{form_version}")
                    ct = st.selectbox("🏷️ Category / Reason", CATEGORIES, key=f"wo_mgr_ct_v{form_version}")
                    amt = st.number_input("💷 Amount (£)", 0.01, step=10.0, key=f"wo_mgr_req_amt_v{form_version}")
                with c2:
                    from datetime import datetime as dt
                    dt_val = st.date_input("📅 Date", value=dt.today(), key=f"wo_mgr_dt_v{form_version}")
                    mgr = st.text_input("👔 Line Manager", key=f"wo_mgr_mgr_v{form_version}")
                    files = st.file_uploader("📎 Attachments", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key=f"wo_mgr_files_v{form_version}")
                    desc = st.text_area("📝 Description / Justification", key=f"wo_mgr_desc_v{form_version}")
                if st.form_submit_button("📤 Send to Director", type="primary"):
                    if en.strip() and mgr.strip() and desc.strip():
                        att_list = []
                        if files:
                            for i, f in enumerate(files, 1):
                                fn = f"ID_{nid}_F{i}_{f.name}"
                                file_path = os.path.join(UPLOAD_DIR, fn)
                                with open(file_path, "wb") as out: out.write(f.getbuffer())
                                att_list.append(fn)
                                _upload_to_drive_bg(file_path, fn)
                        payload = {"id": nid, "emp_name": en.strip(), "dept": request_dept, "type": rt, "category": ct, "date": str(dt_val), "amount": amt, "manager": mgr.strip(), "desc": desc.strip(), "attachment_name": ", ".join(att_list) or "None", "status": "pending", "director_comments": "", "decision_date": "", "decision_by": "", "submitted_by": full_name, "pdf_path": "", "edited_from_id": "", "old_data": ""}
                        save_record_to_excel(payload)
                        log_action("CREATED", nid)
                        st.session_state["wo_mgr_new_req_form_version"] = form_version + 1
                        st.success(f"✅ Request #{nid} sent for approval!"); st.rerun()
                    else: st.error("⚠️ Please fill in: Employee Name, Line Manager, and Description")
            st.divider()
            st.subheader("📋 My Department Requests")
            my_reqs = [r for r in all_live_requests if r.get("dept") == dept_name]
            if not my_reqs: st.info("📋 No requests yet.")
            else:
                for req in reversed(my_reqs):
                    status = req.get("status", "pending").lower()
                    icon = "🟡" if status == "pending" else ("🟢" if status == "approved" else "🔴")
                    with st.expander(f"{icon} ID #{req.get('id')} | {req.get('emp_name')} | {status.upper()} | £{float(req.get('amount',0)):.2f}"):
                        st.write(f"👤 {req.get('emp_name')} | 👔 {req.get('manager')}")
                        st.write(f"🔄 {req.get('type')} | 🏷️ {req.get('category')}")
                        st.info(f"📝 {req.get('desc')}")
                        display_attachments(req)
                        if status == "approved": display_pdf_button(req, can_generate=True)
                        if status in ["pending", "rejected"]:
                            if st.button(f"✏️ Edit Request #{req.get('id')}", key=f"edit_{req.get('id')}"):
                                st.session_state.editing_request_id = req.get("id"); st.rerun()
    with work_order_tab:
        render_work_order_manager_portal(full_name, dept_name, show_total=True)

elif role in ["Manager", "Staff", "Team Member"]:
    dept_name = dept
    has_inspector_bonus = user_info.get("can_access_inspector_bonus", False)
    has_addition_deduction = user_info.get("can_access_addition_deduction", True)
    has_work_orders = user_info.get("can_access_work_orders", False)
    has_hr_leave = user_info.get("can_access_hr_leave", False)
    has_leave_request = user_info.get("can_access_leave_request", False)
    has_employee_hr_reports = user_info.get("can_access_employee_hr_reports", False)
    has_store_deduction = user_info.get("can_access_store_deduction", False)
    has_employee_items = user_info.get("can_access_employee_items", False)
    labels = []
    if has_addition_deduction: labels.append("➕ Addition & Deduction")
    if has_hr_leave:
        labels.append("🏢 HR Department")
    if has_leave_request:
        labels.append("📝 Leave Request")
    if has_employee_hr_reports:
        labels.append("👤 My HR Report")
    if has_store_deduction: labels.append("📦 Store Deduction")
    if has_store_deduction: labels.append("📦 Store Return (Addition)")
    if has_store_deduction: labels.append("📋 My Submitted Store Requests")
    if has_employee_items: labels.append("🧰 Issue Items")
    if has_employee_items: labels.append("📊 Employee Holdings")
    is_store_user = str(dept_name or "").strip().casefold() == "store"
    if has_employee_items: labels.append("📋 Leaver Item Check-in")
    if is_store_user: labels.append("🔗 Leaver Clearance")
    if has_employee_items: labels.append("📚 Check-in History")
    if has_work_orders: labels.append("🛠️ Work Orders")
    if has_inspector_bonus: labels.append("💰 National Grid Inspector Bonus")
    if not labels:
        st.subheader("🔐 Access Restricted")
        st.error("❌ No modules have been enabled for your account. Please contact your Super Admin.")
    else:
        tabs = st.tabs(labels)
        tab_idx = 0
        if has_addition_deduction:
            with tabs[tab_idx]:
                if st.session_state.get("editing_request_id"):
                    eid = st.session_state.editing_request_id
                    rec = next((r for r in all_live_requests if int(r.get("id", 0)) == int(eid)), None)
                    if rec:
                        st.subheader(f"✏️ Edit Request #{eid}")
                        st.markdown("### 📎 Manage Attachments")
                        att_name_raw = rec.get("attachment_name", "None")
                        existing_files = []
                        if att_name_raw and str(att_name_raw).strip().lower() != "none":
                            existing_files = [n.strip() for n in str(att_name_raw).split(",") if n.strip()]
                        files_to_keep = []; files_to_remove = []
                        if existing_files:
                            st.info(f"📋 **{len(existing_files)} attachment(s) currently attached:**")
                            for fname in existing_files:
                                file_path = os.path.join(UPLOAD_DIR, fname)
                                safe_key = f"keep_{eid}_{fname.replace(' ','_').replace('.','_')}"
                                col_check, col_name, col_dl = st.columns([1, 5, 2])
                                keep = col_check.checkbox("✅ Keep", value=True, key=safe_key)
                                col_name.markdown(f"📄 `{fname}`")
                                if os.path.exists(file_path):
                                    with open(file_path, "rb") as f: col_dl.download_button("⬇️", f.read(), file_name=fname, key=f"dl_{safe_key}")
                                else: col_dl.caption("⚠️ Missing")
                                if keep: files_to_keep.append(fname)
                                else: files_to_remove.append(fname)
                            if files_to_remove: st.warning(f"🗑️ Will remove: {', '.join(files_to_remove)}")
                        else: st.info("📋 No attachments currently attached.")
                        st.markdown("#### ➕ Attach New Files")
                        new_files_upload = st.file_uploader("Upload additional files", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key=f"new_upload_{eid}")
                        with st.form("edit_form"):
                            c1, c2 = st.columns(2)
                            with c1:
                                en = st.text_input("👤 Employee Name", rec.get("emp_name", ""))
                                rt = st.selectbox("🔄 Transaction Type", ["Addition", "Deduction"], index=["Addition", "Deduction"].index(rec.get("type", "Addition")))
                                cat_idx = CATEGORIES.index(rec.get("category")) if rec.get("category") in CATEGORIES else 0
                                ct = st.selectbox("🏷️ Category / Reason", CATEGORIES, index=cat_idx)
                                amt = st.number_input("💷 Amount (£)", min_value=0.01, step=10.0, value=float(rec.get("amount", 0.01)))
                            with c2:
                                from datetime import datetime as dt
                                try: d = dt.strptime(str(rec.get("date", ""))[:10], "%Y-%m-%d")
                                except: d = dt.today()
                                dt_val = st.date_input("📅 Date", d)
                                mgr = st.text_input("👔 Line Manager", rec.get("manager", ""))
                                desc = st.text_area("📝 Description / Justification", rec.get("desc", ""))
                            if st.form_submit_button("✅ Submit Edit", type="primary"):
                                final_attachments = list(files_to_keep)
                                if new_files_upload:
                                    for idx, f in enumerate(new_files_upload, start=len(final_attachments)+1):
                                        fn = f"ID_{eid}_EDIT_F{idx}_{f.name}"
                                        edit_path = os.path.join(UPLOAD_DIR, fn)
                                        with open(edit_path, "wb") as outfile: outfile.write(f.getbuffer())
                                        _upload_to_drive_bg(edit_path, fn)
                                        final_attachments.append(fn)
                                records = load_records_from_excel()
                                old_data_dict = {"emp_name": rec.get("emp_name"), "dept": rec.get("dept"), "type": rec.get("type"), "category": rec.get("category"), "date": rec.get("date"), "amount": rec.get("amount"), "manager": rec.get("manager"), "desc": rec.get("desc")}
                                new_data_dict = {"emp_name": en.strip(), "dept": dept_name, "type": rt, "category": ct, "date": str(dt_val), "amount": amt, "manager": mgr.strip(), "desc": desc.strip()}
                                for r in records:
                                    if int(r.get("id", 0)) == int(eid):
                                        r["emp_name"] = en.strip(); r["type"] = rt; r["category"] = ct; r["amount"] = amt; r["date"] = str(dt_val); r["manager"] = mgr.strip(); r["desc"] = desc.strip(); r["status"] = "pending"; r["attachment_name"] = ", ".join(final_attachments) or "None"; r["old_data"] = json.dumps(old_data_dict, default=str)
                                        break
                                log_action("EDITED", eid, old_data=old_data_dict, new_data=new_data_dict)
                                save_all_records(records)
                                st.success(f"✅ Updated!"); st.session_state.editing_request_id = None; st.rerun()
                        if st.button("❌ Cancel", key=f"cancel_edit_{eid}"):
                            st.session_state.editing_request_id = None; st.rerun()
                else:
                    st.subheader(f"➕ New Request — {dept_name}")
                    nid = get_next_id(all_live_requests)
                    form_version = st.session_state.get("new_req_form_version", 0)
                    # Keep Employee Source outside the form so switching between
                    # Registered Employee and Manual Employee immediately reruns the
                    # app and reveals the appropriate employee input.
                    employee_source = st.radio(
                        "👤 Employee Source",
                        ["Registered Employee", "Manual Employee"],
                        horizontal=True,
                        key=f"employee_source_v{form_version}",
                    )
                    with st.form(f"new_req_v{form_version}", clear_on_submit=False):
                        c1, c2 = st.columns(2)
                        with c1:
                            registered_employee = None
                            if employee_source == "Registered Employee":
                                registered = [
                                    e for e in _hrp_load_employees()
                                    if str(e.get("status", "Active")).strip().casefold() == "active"
                                    and str(e.get("name", "")).strip()
                                ]
                                dept_filtered = [
                                    e for e in registered
                                    if str(e.get("department", "")).strip().casefold() == str(dept_name).strip().casefold()
                                ]
                                if dept_filtered:
                                    registered = dept_filtered
                                registered = sorted(registered, key=lambda e: (str(e.get("name", "")).casefold(), str(e.get("emp_id", "")).casefold()))
                                if registered:
                                    labels = [f"{e.get('name','')} — {e.get('emp_id','')} · {e.get('department','')}" for e in registered]
                                    selected_label = st.selectbox("Registered Employee", labels, key=f"registered_emp_v{form_version}")
                                    registered_employee = {label: emp for label, emp in zip(labels, registered)}.get(selected_label)
                                    en = str((registered_employee or {}).get("name", "")).strip()
                                    request_dept = str((registered_employee or {}).get("department", dept_name)).strip() or dept_name
                                    if registered_employee:
                                        st.caption(f"Employee ID: {registered_employee.get('emp_id', '-')} · Department: {request_dept}")
                                else:
                                    st.warning("No active registered employees found for this department. You can use Manual Employee instead.")
                                    en = ""
                                    request_dept = dept_name
                            else:
                                en = st.text_input("👤 Employee Name", key=f"en_v{form_version}")
                                request_dept = dept_name
                            rt = st.selectbox("🔄 Transaction Type", ["Addition", "Deduction"], key=f"rt_v{form_version}")
                            ct = st.selectbox("🏷️ Category / Reason", CATEGORIES, key=f"ct_v{form_version}")
                            amt = st.number_input("💷 Amount (£)", 0.01, step=10.0, key=f"amt_v{form_version}")
                        with c2:
                            from datetime import datetime as dt
                            dt_val = st.date_input("📅 Date", value=dt.today(), key=f"dt_v{form_version}")
                            mgr = st.text_input("👔 Line Manager", key=f"mgr_v{form_version}")
                            files = st.file_uploader("📎 Attachments", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key=f"files_v{form_version}")
                            desc = st.text_area("📝 Description / Justification", key=f"desc_v{form_version}")
                        if st.form_submit_button("📤 Send to Director", type="primary"):
                            if en.strip() and mgr.strip() and desc.strip():
                                att_list = []
                                if files:
                                    for i, f in enumerate(files, 1):
                                        fn = f"ID_{nid}_F{i}_{f.name}"
                                        file_path = os.path.join(UPLOAD_DIR, fn)
                                        with open(file_path, "wb") as out: out.write(f.getbuffer())
                                        att_list.append(fn)
                                        _upload_to_drive_bg(file_path, fn)
                                payload = {"id": nid, "emp_name": en.strip(), "dept": request_dept, "type": rt, "category": ct, "date": str(dt_val), "amount": amt, "manager": mgr.strip(), "desc": desc.strip(), "attachment_name": ", ".join(att_list) or "None", "status": "pending", "director_comments": "", "decision_date": "", "decision_by": "", "submitted_by": full_name, "pdf_path": "", "edited_from_id": "", "old_data": ""}
                                save_record_to_excel(payload)
                                log_action("CREATED", nid)
                                st.session_state["new_req_form_version"] = form_version + 1
                                st.success(f"✅ Request #{nid} sent for approval!"); st.rerun()
                            else: st.error("⚠️ Please fill in: Employee Name, Line Manager, and Description")
                    st.divider()
                    st.subheader("📋 My Department Requests")
                    my_reqs = [r for r in all_live_requests if r.get("dept") == dept_name]
                    if not my_reqs: st.info("📋 No requests yet.")
                    else:
                        for req in reversed(my_reqs):
                            status = req.get("status", "pending").lower()
                            icon = "🟡" if status == "pending" else ("🟢" if status == "approved" else "🔴")
                            with st.expander(f"{icon} ID #{req.get('id')} | {req.get('emp_name')} | {status.upper()} | £{float(req.get('amount',0)):.2f}"):
                                st.write(f"👤 {req.get('emp_name')} | 👔 {req.get('manager')}")
                                st.write(f"🔄 {req.get('type')} | 🏷️ {req.get('category')}")
                                st.info(f"📝 {req.get('desc')}")
                                display_attachments(req)
                                if req.get("director_comments"): st.info(f"💬 Director Comments: {req.get('director_comments')}")
                                if status == "approved": display_pdf_button(req, can_generate=True)
                                if status in ["pending", "rejected"]:
                                    if st.button(f"✏️ Edit Request #{req.get('id')}", key=f"edit_{req.get('id')}"):
                                        st.session_state.editing_request_id = req.get("id"); st.rerun()
            tab_idx += 1
        if has_hr_leave:
            with tabs[tab_idx]:
                render_hr_department(current_user_info=user_info, has_hr_access=True, director_name=full_name)
            tab_idx += 1
        if has_leave_request:
            with tabs[tab_idx]:
                normalized_dept = re.sub(r"\s+", " ", str(dept_name or "")).strip().casefold()
                is_hr_manager_account = normalized_dept in {"hr", "human resource", "human resources", "hr department", "human resource department"} and has_hr_leave
                if is_hr_manager_account:
                    render_hr_leave_approvals()
                else:
                    render_department_manager_leave_request(current_user_info=user_info)
            tab_idx += 1
        if has_employee_hr_reports:
            with tabs[tab_idx]:
                render_employee_hr_reports(user_info)
            tab_idx += 1
        if has_store_deduction:
            with tabs[tab_idx]:
                render_store_deduction_form(full_name, dept_name)
            tab_idx += 1
        if has_store_deduction:
            with tabs[tab_idx]:
                render_store_return_form(full_name, dept_name)
            tab_idx += 1
        if has_store_deduction:
            with tabs[tab_idx]:
                render_store_my_submissions(full_name)
            tab_idx += 1
        if has_employee_items:
            with tabs[tab_idx]:
                render_item_issue_form(full_name, dept_name)
            tab_idx += 1
        if has_employee_items:
            with tabs[tab_idx]:
                render_employee_holdings_overview(role)
            tab_idx += 1
        if has_employee_items:
            with tabs[tab_idx]:
                render_item_checkin_form(full_name)
            tab_idx += 1
        if is_store_user:
            with tabs[tab_idx]:
                render_leaver_clearance_store(full_name)
            tab_idx += 1
        if has_employee_items:
            with tabs[tab_idx]:
                render_item_checkin_history()
            tab_idx += 1
        if has_work_orders:
            with tabs[tab_idx]:
                render_work_order_employee_portal(full_name, dept_name)
            tab_idx += 1
        if has_inspector_bonus:
            with tabs[tab_idx]:
                render_inspector_bonus_portal(full_name, dept_name)
            tab_idx += 1


elif role == "Director":
    director_hr_access_enabled = any(bool(user_info.get(k, False)) for k in (
        "can_access_holiday_calendar", "can_access_hr_reports",
        "can_access_employee_overview", "can_access_holiday_calculator",
        "can_access_employee_directory"
    ))
    director_tab_labels = [
        "➕ Addition & Deduction",
        "👥 HR Leave Settlement",
        "📦 Store Deductions",
        "📦 Store Returns (Additions)",
        "🧰 Employee Items (View)",
        "🛠️ Work Orders",
        "💰 National Grid Inspector Bonus"
    ]
    if director_hr_access_enabled:
        director_tab_labels.insert(1, "📊 HR Reports & Calendar")
    director_tabs = st.tabs(director_tab_labels)
    director_addition_tab = director_tabs[0]
    if director_hr_access_enabled:
        director_hr_access_tab = director_tabs[1]
        director_hr_leave_tab = director_tabs[2]
        director_store_ded_tab = director_tabs[3]
        director_store_ret_tab = director_tabs[4]
        director_employee_items_tab = director_tabs[5]
        director_work_order_tab = director_tabs[6]
        director_inspector_tab = director_tabs[7]
    else:
        director_hr_access_tab = None
        director_hr_leave_tab = director_tabs[1]
        director_store_ded_tab = director_tabs[2]
        director_store_ret_tab = director_tabs[3]
        director_employee_items_tab = director_tabs[4]
        director_work_order_tab = director_tabs[5]
        director_inspector_tab = director_tabs[6]
    with director_addition_tab:
        st.subheader(f"🎛️ Director Approval Portal — {full_name}")
        st.info("✅ Review all requests, Approve, Reject, OR Change Status. Decisions update automatically.")
        st.info("🔄 **Director can change ANY request to ANY status at ANY time.** All changes are logged.")
        st.divider()
        tab_pending, tab_approved, tab_rejected = st.tabs(["⏳ Pending Requests", "✅ Approved Requests", "❌ Rejected Requests"])
        with tab_pending:
            pending = [r for r in all_live_requests if r.get("status") == "pending"]
            if not pending: st.success("✅ No pending requests!")
            else:
                st.metric("⏳ Pending Requests", len(pending)); st.divider()
                for req in reversed(pending):
                    req_id = req.get("id")
                    with st.expander(f"🟡 ID #{req_id} | {req.get('emp_name')} | 🔄 {req.get('type', 'Not specified')} | £{float(req.get('amount',0)):.2f} | {req.get('dept')}"):
                        col_left, col_right = st.columns([2, 1])
                        with col_left:
                            st.write(f"👤 **Employee:** {req.get('emp_name')}")
                            st.write(f"🏢 **Department:** {req.get('dept')}")
                            st.write(f"🔄 **Type:** {req.get('type')}")
                            st.write(f"🏷️ **Category:** {req.get('category')}")
                            st.write(f"💷 **Amount:** £{float(req.get('amount',0)):.2f}")
                            st.write(f"👔 **Line Manager:** {req.get('manager')}")
                            submitted_by = get_submitted_by(req)
                            if submitted_by: st.write(f"📝 **Submitted by:** {submitted_by}")
                            st.info(f"📝 **Description:**\n{req.get('desc','')}")
                            display_attachments(req)
                            old_data_raw = req.get("old_data", "")
                            if old_data_raw and old_data_raw not in ["", "{}", "None"]:
                                st.divider(); st.markdown("### 🔄 What Changed / Edits"); show_old_new_comparison(old_data_raw, req)
                        with col_right:
                            st.markdown("### ✍️ Decision")
                            comments = st.text_area("Director Comments", key=f"comm_{req_id}")
                            approve_btn = st.button("✅ APPROVE", type="primary", key=f"appr_{req_id}")
                            reject_btn = st.button("❌ REJECT", type="secondary", key=f"rejt_{req_id}")
                            if approve_btn:
                                records = load_records_from_excel()
                                for r in records:
                                    if int(r.get("id",0)) == int(req_id):
                                        r["status"] = "approved"; r["approved_by"] = full_name; r["decision_by"] = full_name; r["director_comments"] = comments; r["decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S"); r["approved_date"] = r["decision_date"]
                                        break
                                save_all_records(records); log_action("APPROVED", req_id)
                                st.success(f"✅ Request #{req_id} APPROVED."); st.rerun()
                            if reject_btn:
                                if not comments.strip():
                                    st.error("❌ A rejection reason/comment is required.")
                                else:
                                    records = load_records_from_excel()
                                    for r in records:
                                        if int(r.get("id",0)) == int(req_id):
                                            r["status"] = "rejected"
                                            r["decision_by"] = full_name
                                            r["director_comments"] = comments.strip()
                                            r["rejection_reason"] = comments.strip()
                                            r["decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                            break
                                    save_all_records(records); log_action("REJECTED", req_id)
                                    st.error(f"❌ Request #{req_id} REJECTED."); st.rerun()
        with tab_approved:
            approved = [r for r in all_live_requests if r.get("status") == "approved"]
            if not approved: st.info("📋 No approved requests yet.")
            else:
                st.metric("✅ Approved Requests", len(approved)); st.divider()
                for req in reversed(approved):
                    req_id = req.get("id")
                    with st.expander(f"🟢 ID #{req_id} | {req.get('emp_name')} | 🔄 {req.get('type', 'Not specified')} | £{float(req.get('amount',0)):.2f}"):
                        st.write(f"👤 {req.get('emp_name')} | 💷 £{float(req.get('amount',0)):.2f}")
                        display_attachments(req)
                        st.divider(); display_pdf_button(req, can_generate=True)
                        st.markdown("### 🔄 Change Status")
                        new_comments = st.text_area("Add comment (optional)", key=f"chg_comm_app_{req_id}")
                        col1, col2 = st.columns(2)
                        with col1:
                            if st.button("⏳ Move to Pending", key=f"app_to_pend_{req_id}"):
                                records = load_records_from_excel()
                                for r in records:
                                    if int(r.get("id",0)) == int(req_id):
                                        r["status"] = "pending"; r["decision_date"] = ""; r["decision_by"] = ""
                                        break
                                save_all_records(records); log_action("STATUS_CHANGED", req_id, old_data={"status":"approved"}, new_data={"status":"pending"})
                                st.success(f"✅ Request #{req_id} moved to Pending."); st.rerun()
                        with col2:
                            if st.button("❌ Change to REJECTED", key=f"app_to_rej_{req_id}"):
                                if not new_comments.strip():
                                    st.error("❌ A rejection reason/comment is required.")
                                else:
                                    records = load_records_from_excel()
                                    for r in records:
                                        if int(r.get("id",0)) == int(req_id):
                                            r["status"] = "rejected"
                                            r["decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                            r["decision_by"] = full_name
                                            r["director_comments"] = new_comments.strip()
                                            r["rejection_reason"] = new_comments.strip()
                                            break
                                    save_all_records(records); log_action("STATUS_CHANGED", req_id, old_data={"status":"approved"}, new_data={"status":"rejected"})
                                    st.success(f"✅ Request #{req_id} changed to Rejected."); st.rerun()
        with tab_rejected:
            rejected = [r for r in all_live_requests if r.get("status") == "rejected"]
            if not rejected: st.success("✅ No rejected requests!")
            else:
                st.metric("❌ Rejected Requests", len(rejected)); st.divider()
                for req in reversed(rejected):
                    req_id = req.get("id")
                    with st.expander(f"🔴 ID #{req_id} | {req.get('emp_name')} | 🔄 {req.get('type', 'Not specified')} | £{float(req.get('amount',0)):.2f}"):
                        st.write(f"👤 {req.get('emp_name')} | 🔄 **Type:** {req.get('type', 'Not specified')} | 💷 £{float(req.get('amount',0)):.2f}")
                        st.error(f"❌ **Rejected by:** {req.get('decision_by', '—')} on {format_date(req.get('decision_date', ''))}")
                        reason = req.get('rejection_reason') or req.get('director_comments') or 'No rejection reason recorded.'
                        st.error(f"💬 **Rejection reason:** {reason}")
                        display_attachments(req)
                        st.markdown("### 🔄 Change Status")
                        new_comments = st.text_area("Add comment (optional)", key=f"chg_comm_rej_{req_id}")
                        col1, col2 = st.columns(2)
                        with col1:
                            if st.button("⏳ Move to Pending", key=f"rej_to_pend_{req_id}"):
                                records = load_records_from_excel()
                                for r in records:
                                    if int(r.get("id",0)) == int(req_id):
                                        r["status"] = "pending"; r["decision_date"] = ""; r["decision_by"] = ""
                                        break
                                save_all_records(records); log_action("STATUS_CHANGED", req_id, old_data={"status":"rejected"}, new_data={"status":"pending"})
                                st.success(f"✅ Request #{req_id} moved to Pending."); st.rerun()
                        with col2:
                            if st.button("✅ Change to APPROVED", key=f"rej_to_app_{req_id}"):
                                records = load_records_from_excel()
                                for r in records:
                                    if int(r.get("id",0)) == int(req_id):
                                        r["status"] = "approved"; r["decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S"); r["decision_by"] = full_name
                                        break
                                save_all_records(records); log_action("STATUS_CHANGED", req_id, old_data={"status":"rejected"}, new_data={"status":"approved"})
                                st.success(f"✅ Request #{req_id} changed to Approved."); st.rerun()
    if director_hr_access_enabled and director_hr_access_tab is not None:
        with director_hr_access_tab:
            render_director_hr_access_portal(user_info)

    with director_hr_leave_tab:
        render_hr_leave_director_portal(full_name)

    with director_store_ded_tab:
        render_store_director_portal(full_name, type_filter="Deduction")
    with director_store_ret_tab:
        render_store_director_portal(full_name, type_filter="Addition")
    with director_employee_items_tab:
        sub_holdings, sub_checkins = st.tabs(["📊 Employee Holdings", "📚 Check-in History"])
        with sub_holdings:
            render_employee_holdings_overview("Director")
        with sub_checkins:
            render_item_checkin_history()
    with director_work_order_tab: render_work_order_director_portal(full_name)
    with director_inspector_tab: render_inspector_bonus_director_portal(full_name)

elif role == "Super Admin":
    super_add_ded_tab, super_store_ded_tab, super_store_ret_tab, super_employee_items_tab, super_work_orders_tab, super_inspector_bonus_tab, super_data_control_tab, super_system_mgmt_tab = st.tabs([
        "➕ Addition & Deduction",
        "📦 Store Deductions",
        "📦 Store Returns (Additions)",
        "🧰 Employee Items",
        "🛠️ Work Orders",
        "💰 National Grid Inspector Bonus",
        "🛡️ Data Control",
        "🔧 System Management"
    ])
    with super_add_ded_tab:
        st.subheader("🛡️ Super Admin — All Addition & Deduction Requests")
        st.info("✅ View ALL requests across ALL departments. Download PDFs. **Approval → Director only.**")
        st.divider()
        tab_pending, tab_approved, tab_rejected = st.tabs(["⏳ All Pending", "✅ All Approved", "❌ All Rejected"])
        with tab_pending:
            pending = [r for r in all_live_requests if str(r.get("status", "")).strip().lower() == "pending"]
            if not pending: st.success("✅ No pending requests.")
            else:
                st.metric("⏳ All Pending", len(pending)); st.divider()
                for req in reversed(pending):
                    req_id = req.get("id"); amount = float(req.get("amount", 0))
                    with st.expander(f"🟡 ID #{req_id} | {req.get('emp_name')} | {req.get('dept')} | £{amount:.2f}"):
                        st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept')}")
                        st.write(f"🔄 Type: {req.get('type')} | 🏷️ Category: {req.get('category')}")
                        st.write(f"💷 Amount: £{amount:.2f}")
                        submitted_by = get_submitted_by(req)
                        if submitted_by: st.write(f"📝 **Submitted by:** {submitted_by}")
                        st.info(f"📝 Description: {req.get('desc')}")
                        display_attachments(req)
                        if req.get("director_comments"): st.info(f"💬 Director Comments: {req.get('director_comments')}")
                        st.divider(); display_pdf_button(req, can_generate=True)
        with tab_approved:
            approved = [r for r in all_live_requests if str(r.get("status", "")).strip().lower() == "approved"]
            if not approved: st.info("📋 No approved requests.")
            else:
                st.metric("✅ All Approved", len(approved)); st.divider()
                for req in reversed(approved):
                    req_id = req.get("id"); amount = float(req.get("amount", 0))
                    dec_by = req.get('decision_by', 'Director')
                    with st.expander(f"🟢 ID #{req_id} | {req.get('emp_name')} | {req.get('dept')} | £{amount:.2f} | ✅ {dec_by}"):
                        st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept')}")
                        st.write(f"💷 Amount: £{amount:.2f}")
                        st.success(f"💬 Director Comments: {req.get('director_comments', 'None')}")
                        display_attachments(req)
                        st.divider(); display_pdf_button(req, can_generate=True)
        with tab_rejected:
            rejected = [r for r in all_live_requests if str(r.get("status", "")).strip().lower() == "rejected"]
            if not rejected: st.info("📋 No rejected requests.")
            else:
                st.metric("❌ All Rejected", len(rejected)); st.divider()
                for req in reversed(rejected):
                    req_id = req.get("id"); amount = float(req.get("amount", 0))
                    dec_by = req.get('decision_by', 'Director')
                    with st.expander(f"🔴 ID #{req_id} | {req.get('emp_name')} | {req.get('dept')} | £{amount:.2f}"):
                        st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept')}")
                        st.write(f"💷 Amount: £{amount:.2f}")
                        st.error(f"❌ Rejected By: {dec_by}")
                        st.error(f"💬 Reason: {req.get('director_comments', 'None')}")
                        display_attachments(req)
                        st.divider(); display_pdf_button(req, can_generate=True)
    with super_store_ded_tab:
        render_store_super_admin(type_filter="Deduction")
    with super_store_ret_tab:
        render_store_super_admin(type_filter="Addition")
    with super_employee_items_tab:
        st.subheader("🧰 Employee Items — Super Admin")
        inner_issue, inner_holdings, inner_checkin, inner_history = st.tabs([
            "🧰 Issue Items", "📊 Holdings", "📋 Leaver Check-in", "📚 Check-in History"
        ])
        with inner_issue:
            render_item_issue_form(full_name, "Super Admin")
        with inner_holdings:
            render_employee_holdings_overview("Super Admin")
        with inner_checkin:
            render_item_checkin_form(full_name)
        with inner_history:
            render_item_checkin_history()
    with super_work_orders_tab:
        render_work_orders_super_admin()
    with super_inspector_bonus_tab:
        render_inspector_bonus_super_admin()
    with super_data_control_tab:
        _super_admin_transaction_control()
    with super_system_mgmt_tab:
        st.subheader("🔧 System Management — Super Admin")
        st.info("🛡️ Manage system settings, users, audit history and data reset controls.")
        st.divider()
        tab_settings, tab_hr_settings, tab_store_settings, tab_users, tab_audit = st.tabs([
            "⚙️ System Settings",
            "👥 HR Leave Settings",
            "📦 Store Settings",
            "👤 User Management",
            "📖 Audit History"
        ])
        with tab_settings: settings_management_panel()
        with tab_hr_settings: render_hr_leave_settings()
        with tab_store_settings: render_store_items_settings()
        with tab_users: user_management_panel()
        with tab_audit:
            if "display_audit_log_panel" in globals(): display_audit_log_panel()
            else: st.info("📖 Audit log panel not defined — skipping")
            st.divider()
            st.subheader("⚠️ Super Admin — Data Reset / Live Launch")
            st.warning("These controls are permanent. They are intended for preparing the portal for live use.")
            danger_col1, danger_col2 = st.columns(2)
            with danger_col1:
                if not st.session_state.get("confirm_clear_requests", False):
                    if st.button("🧹 Clear All Submitted Requests", type="secondary", key="clear_all_requests_btn"):
                        st.session_state["confirm_clear_requests"] = True; st.rerun()
                else:
                    st.error("⚠️ This will permanently remove ALL request records.")
                    c1, c2 = st.columns(2)
                    with c1:
                        if st.button("✅ Yes, Clear Requests", type="primary", key="confirm_clear_all_requests_btn"):
                            clear_all_requests_file(); st.session_state["confirm_clear_requests"] = False
                            st.success("✅ Cleared."); st.rerun()
                    with c2:
                        if st.button("↩️ Cancel", key="cancel_clear_all_requests_btn"):
                            st.session_state["confirm_clear_requests"] = False; st.rerun()
            with danger_col2:
                reset_col1, reset_col2, reset_col3, reset_col4 = st.columns(4)
                with reset_col1:
                    if not st.session_state.get("confirm_clear_inspector_bonus", False):
                        if st.button("💰 Clear Inspector Bonuses", key="super_admin_clear_all_inspector_bonus", type="secondary", width="stretch"):
                            st.session_state["confirm_clear_inspector_bonus"] = True
                    else:
                        if st.button("✅ Confirm", key="super_admin_confirm_clear_inspector_bonus", width="stretch"):
                            clear_all_inspector_bonus()
                            log_action("SUPER_ADMIN_CLEAR_INSPECTOR_BONUS", "ALL", decision_by=full_name)
                            st.session_state["confirm_clear_inspector_bonus"] = False
                            st.success("✅ Cleared."); st.rerun()
                with reset_col2:
                    if not st.session_state.get("confirm_clear_all_work_orders", False):
                        if st.button("🛠️ Clear Work Orders", key="super_admin_clear_all_work_orders", type="secondary", width="stretch"):
                            st.session_state["confirm_clear_all_work_orders"] = True
                    else:
                        if st.button("✅ Confirm", key="super_admin_confirm_clear_all_work_orders", width="stretch"):
                            save_all_work_orders([])
                            st.session_state["confirm_clear_all_work_orders"] = False
                            st.success("✅ Cleared."); st.rerun()
                with reset_col3:
                    if not st.session_state.get("confirm_clear_hr_leave", False):
                        if st.button("👥 Clear HR Leave", key="super_admin_clear_all_hr_leave", type="secondary", width="stretch"):
                            st.session_state["confirm_clear_hr_leave"] = True
                    else:
                        if st.button("✅ Confirm", key="super_admin_confirm_clear_all_hr_leave", width="stretch"):
                            clear_all_hr_leave()
                            st.session_state["confirm_clear_hr_leave"] = False
                            st.success("✅ Cleared."); st.rerun()
                with reset_col4:
                    if not st.session_state.get("confirm_clear_store_deduction", False):
                        if st.button("📦 Clear Store Transactions", key="super_admin_clear_all_store_deduction", type="secondary", width="stretch"):
                            st.session_state["confirm_clear_store_deduction"] = True
                    else:
                        if st.button("✅ Confirm", key="super_admin_confirm_clear_all_store_deduction", width="stretch"):
                            clear_all_store_deductions()
                            st.session_state["confirm_clear_store_deduction"] = False
                            st.success("✅ Cleared."); st.rerun()
                with reset_col4:
                    if not st.session_state.get("confirm_clear_audit", False):
                        if st.button("🗑️ Clear Audit", type="secondary", key="clear_audit_history_btn"):
                            st.session_state["confirm_clear_audit"] = True; st.rerun()
                    else:
                        if st.button("✅ Confirm", type="primary", key="confirm_clear_audit_btn"):
                            clear_audit_log_file(); st.session_state["confirm_clear_audit"] = False
                            st.success("✅ Cleared."); st.rerun()
        st.divider()
        st.subheader("📥 Download Data Backups")
        st.caption("Download a copy before using the live-launch reset. User accounts and system settings are kept separately.")

        backup_files = [
            ("📥 Requests", EXCEL_PATH, "requests", "backup_requests"),
            ("📥 Users", USER_DB_PATH, "users", "backup_users"),
            ("📥 Settings", SETTINGS_PATH, "settings", "backup_settings"),
            ("📥 Inspector Bonuses", INSPECTOR_BONUS_PATH, "inspector_bonus", "backup_inspector_bonus"),
            ("📥 Work Orders", WORK_ORDERS_PATH, "work_orders", "backup_work_orders"),
            ("📥 HR Leave Requests", HR_LEAVE_PATH, "hr_leave_requests", "backup_hr_leave"),
            ("📥 HR Daily Rates", HR_DAILY_RATES_PATH, "hr_daily_rates", "backup_hr_rates"),
            ("📥 Store Transactions", STORE_DEDUCTION_PATH, "store_transactions", "backup_store_transactions"),
            ("📥 Store Items", STORE_ITEMS_PATH, "store_items", "backup_store_items"),
            ("📥 HR Employee Records", HR_EMPLOYEES_PATH, "hr_employee_records", "backup_hr_employee_records"),
            ("📥 HR Portal Leave Records", HR_PORTAL_LEAVE_PATH, "hr_portal_leave_records", "backup_hr_portal_leave_records"),
            ("📥 Audit Log", AUDIT_LOG_PATH, "audit_log", "backup_audit_log"),
            ("📥 Employee Items", EMPLOYEE_ITEMS_PATH, "employee_items", "backup_employee_items"),
            ("📥 Item Check-ins", ITEM_CHECKIN_PATH, "item_checkins", "backup_item_checkins"),
            ("📥 Leaver Clearance", LEAVER_CLEARANCE_PATH, "leaver_clearance", "backup_leaver_clearance"),
        ]
        backup_cols = st.columns(3)
        for i, (label, path, stem, key) in enumerate(backup_files):
            with backup_cols[i % 3]:
                if os.path.exists(path):
                    try:
                        with open(path, "rb") as f:
                            st.download_button(
                                label,
                                f.read(),
                                file_name=f"BACKUP_{stem}_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                                type="primary",
                                key=key,
                                width="stretch",
                            )
                    except Exception as e:
                        st.warning(f"Unable to prepare {label}: {e}")
                else:
                    st.button(f"{label} (not available)", disabled=True, key=f"{key}_missing")

        st.divider()
        st.subheader("🚀 Make Software Live")
        st.warning(
            "Use this only after downloading the backups above. "
            "It permanently clears all operational/test data: requests, work orders, "
            "inspector bonuses, HR leave requests and audit history. User accounts and system settings are NOT deleted."
        )

        def _clear_live_launch_data():
            clear_all_requests_file()
            clear_all_inspector_bonus()
            _write_empty_excel(WORK_ORDERS_PATH, WORK_ORDER_COLUMNS)
            _write_empty_excel(HR_LEAVE_PATH, HR_LEAVE_COLUMNS)
            _write_empty_excel(STORE_DEDUCTION_PATH, STORE_DEDUCTION_COLUMNS)
            _write_empty_excel(EMPLOYEE_ITEMS_PATH, EMPLOYEE_ITEMS_COLUMNS)
            _write_empty_excel(ITEM_CHECKIN_PATH, ITEM_CHECKIN_COLUMNS)
            _invalidate_data_cache(
                "_work_orders_cache",
                "_work_order_cache",
                "_inspector_bonus_cache",
                "_records_cache",
                "_hr_leave_cache",
                "_store_deduction_cache",
                "_employee_items_cache",
                "_item_checkins_cache",
            )
            sync_saved_file_to_drive(WORK_ORDERS_PATH)
            sync_saved_file_to_drive(HR_LEAVE_PATH)
            sync_saved_file_to_drive(STORE_DEDUCTION_PATH)
            sync_saved_file_to_drive(EMPLOYEE_ITEMS_PATH)
            sync_saved_file_to_drive(ITEM_CHECKIN_PATH)
            clear_audit_log_file()

            for folder in (PDF_DIR, WORK_ORDER_PDF_DIR, INSPECTOR_BONUS_PDF_DIR, HR_LEAVE_PDF_DIR, STORE_DEDUCTION_PDF_DIR, UPLOAD_DIR):
                if os.path.isdir(folder):
                    for root, dirs, files in os.walk(folder, topdown=False):
                        for filename in files:
                            try:
                                os.remove(os.path.join(root, filename))
                            except OSError:
                                pass
                        for dirname in dirs:
                            try:
                                os.rmdir(os.path.join(root, dirname))
                            except OSError:
                                pass
            st.session_state["_live_data_reset"] = datetime.now().isoformat()

        if not st.session_state.get("confirm_live_launch", False):
            if st.button(
                "🚀 Make Software Live — Clear All Operational Data",
                type="primary",
                key="make_software_live_btn",
                width="stretch",
            ):
                st.session_state["confirm_live_launch"] = True
                st.rerun()
        else:
            st.error(
                "⚠️ FINAL CONFIRMATION: this will permanently remove all requests, "
                "work orders, inspector bonuses, HR leave requests, Store Deductions, audit history, generated PDFs and uploaded attachments. "
                "Users and system settings will remain."
            )
            live_c1, live_c2 = st.columns(2)
            with live_c1:
                if st.button(
                    "🚀 YES — Make Software Live",
                    type="primary",
                    key="confirm_make_software_live_btn",
                    width="stretch",
                ):
                    try:
                        _clear_live_launch_data()
                        st.session_state["confirm_live_launch"] = False
                        st.success("✅ Software is ready for live use. All operational/test data has been cleared.")
                        st.rerun()
                    except Exception as e:
                        st.session_state["confirm_clear_live_launch"] = False
                        st.error(f"❌ Live launch reset failed: {e}")
            with live_c2:
                if st.button("↩️ Cancel", key="cancel_make_software_live_btn", width="stretch"):
                    st.session_state["confirm_live_launch"] = False
                    st.rerun()

else:
    st.subheader("🔐 Access Restricted")
    st.error("❌ Your role does not have a defined portal. Please contact Super Admin.")
# ========================================================
# ✅ END OF ROLE-BASED PORTALS
# ============================================================
# ✅ END OF FILE — NOTHING AFTER THIS!
# ============================================================
