# ============================================================
# 🔄 ACOOLE PORTAL — PROFESSIONAL VERSION v4.19
# ============================================================
# ✅ v4.19 (HR LEAVE SETTLEMENT MODULE):
#    • Added HR Leave Settlement module (holiday payouts/deductions).
#    • Super Admin manages HR categories + per-department daily rates.
#    • Director approves/rejects with mandatory rejection reason.
#    • Auto-calculates amount from days × daily rate (editable override).
#    • Payroll read-only view of approved HR settlements.
#    • New permission flag: can_access_hr_leave.
# ✅ v4.18 (ATTACHMENT PRESERVATION FIX)
# ✅ v4.17 (BASE64 WHITESPACE & PADDING FIX)
# ✅ v4.16 (PERMANENT GOOGLE DRIVE FIX - BASE64 METHOD)
# ✅ v4.15 (WORK ORDER TOTAL TAB RESTRUCTURE)
# ✅ v4.14 (WORK ORDER TOTAL FOR EMPLOYEES/TEAM MEMBERS)
# ✅ v4.13 (GRANULAR WORK ORDER TOTAL PERMISSION)
# ✅ v4.12 (USER ACTIVE / INACTIVE + SUPER ADMIN FULL ACCESS)
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
import requests
import threading
from datetime import datetime, date, timezone

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
os.makedirs(HR_LEAVE_PDF_DIR, exist_ok=True)

HR_LEAVE_COLUMNS = [
    "ID", "Employee Name", "Employee Department", "Transaction Type",
    "Category Reason", "Owe Owed", "Date", "Number of Days", "Amount (£)",
    "Line Manager", "Description", "Attachment Name", "Status",
    "Director Comments", "Rejection Reason", "Decision Date", "Decision By",
    "Submitted By", "Submitted Date", "PDF File Path"
]
HR_DAILY_RATES_COLUMNS = ["Department", "Transaction Type", "Daily Rate (£)", "Active"]
DEFAULT_HR_CATEGORIES = [
    "Annual Leave Balance", "Unused Holiday Payout",
    "Overused Holiday Deduction", "Leave Encashment",
]
DEFAULT_OWE_OWED = ["Company Owes Employee", "Employee Owes Company"]

USER_DB_COLUMNS = [
    "full_name", "username", "password", "role", "dept",
    "can_view_all_dept", "can_generate_pdf", "can_download_data",
    "can_approve_requests", "can_access_inspector_bonus",
    "can_access_addition_deduction", "can_access_work_orders",
    "can_access_wo_total",
    "can_access_hr_leave",
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

_DRIVE_ID_CACHE = {}
_DRIVE_SYNC_FINGERPRINTS = {}
_DRIVE_SYNC_LOCK = threading.RLock()

# ============================================================
# GOOGLE DRIVE CONNECTION — SERVICE ACCOUNT (BASE64 METHOD v4.18)
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
except Exception as e:
    drive_service = None
    print(f"Google Drive initialisation failed; local storage will be used: {e}")

def upload_to_google_drive(local_file_path, display_filename):
    if drive_service is None or not os.path.exists(local_file_path): return None
    cache_key = f"{GOOGLE_DRIVE_FOLDER_ID}::{display_filename}"
    def _do(file_id=None):
        media = MediaFileUpload(local_file_path, resumable=False)
        if file_id:
            return drive_service.files().update(fileId=file_id, media_body=media, fields="id,name,parents").execute()
        metadata = {"name": display_filename, "parents": [GOOGLE_DRIVE_FOLDER_ID]}
        return drive_service.files().create(body=metadata, media_body=media, fields="id,name,parents").execute()
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
        print(f"Google Drive upload failed: {e}")
        return None

def _drive_find_file(filename, parent_id=GOOGLE_DRIVE_FOLDER_ID):
    if drive_service is None: return None
    try:
        safe_name = str(filename).replace("'", "\\'")
        q = (f"name = '{safe_name}' and '{parent_id}' in parents and trashed = false")
        result = drive_service.files().list(q=q, spaces="drive", fields="files(id,name,modifiedTime)", orderBy="modifiedTime desc", pageSize=10).execute()
        files = result.get("files", [])
        return files[0] if files else None
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
            return drive_service.files().update(fileId=file_id, media_body=media, fields="id,name").execute()
        metadata = {"name": filename, "parents": [parent_id]}
        return drive_service.files().create(body=metadata, media_body=media, fields="id,name").execute()
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
        print(f"Drive upload failed for {filename}: {e}")
        return None

def _drive_download_file(file_id, local_path):
    if drive_service is None:
        return False
    tmp_path = f"{local_path}.tmp"
    try:
        request = drive_service.files().get_media(fileId=file_id)
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

def sync_persistent_file(local_path, columns=None):
    if drive_service is None:
        if not os.path.exists(local_path) and columns is not None:
            pd.DataFrame(columns=columns).to_excel(local_path, index=False, engine="openpyxl")
        return

    with _DRIVE_SYNC_LOCK:
        filename = os.path.basename(local_path)
        remote = _drive_find_file(filename)
        if remote:
            try:
                rmt = str(remote.get("modifiedTime", "")).rstrip("Z")
                local_newer = False
                if os.path.exists(local_path) and rmt:
                    try:
                        from datetime import timezone
                        if "." in rmt:
                            remote_dt = datetime.strptime(rmt, "%Y-%m-%dT%H:%M:%S.%f").replace(tzinfo=timezone.utc)
                        else:
                            remote_dt = datetime.strptime(rmt, "%Y-%m-%dT%H:%M:%S").replace(tzinfo=timezone.utc)
                        local_dt = datetime.fromtimestamp(os.path.getmtime(local_path), timezone.utc)
                        local_newer = local_dt > remote_dt
                    except Exception:
                        local_newer = False
                if not local_newer:
                    if not _drive_download_file(remote["id"], local_path):
                        print(f"Using local copy of {filename} (Drive download failed)")
            except Exception as e:
                print(f"Drive sync check failed for {filename}: {e}")
                if not os.path.exists(local_path) and columns is not None:
                    pd.DataFrame(columns=columns).to_excel(local_path, index=False, engine="openpyxl")
        elif os.path.exists(local_path):
            _drive_upload_path(local_path, filename)
        elif columns is not None:
            pd.DataFrame(columns=columns).to_excel(local_path, index=False, engine="openpyxl")
            _drive_upload_path(local_path, filename)

def sync_saved_file_to_drive(local_path):
    if drive_service is None or not os.path.exists(local_path):
        return
    try:
        st_info = os.stat(local_path)
        fingerprint = f"{st_info.st_size}:{int(st_info.st_mtime_ns)}"
    except Exception:
        return
    with _DRIVE_SYNC_LOCK:
        if _DRIVE_SYNC_FINGERPRINTS.get(local_path) == fingerprint:
            return
        try:
            if _drive_upload_path(local_path) is not None:
                _DRIVE_SYNC_FINGERPRINTS[local_path] = fingerprint
        except Exception as e:
            print(f"Drive sync failed for {local_path}: {e}")
            _DRIVE_SYNC_FINGERPRINTS.pop(local_path, None)

def _upload_to_drive_bg(local_path, filename):
    if drive_service is None or not os.path.exists(local_path):
        return
    with _DRIVE_SYNC_LOCK:
        try:
            upload_to_google_drive(local_path, filename)
        except Exception as e:
            print(f"Drive upload failed for {filename}: {e}")

def initialise_drive_storage():
    if drive_service is None or st.session_state.get("drive_storage_initialised"): return
    os.makedirs(APP_FOLDER, exist_ok=True)
    with _DRIVE_SYNC_LOCK:
        targets = [
            (EXCEL_PATH, EXCEL_COLUMNS),
            (USER_DB_PATH, USER_DB_COLUMNS),
            (SETTINGS_PATH, ["setting", "value"]),
            (AUDIT_LOG_PATH, AUDIT_COLUMNS),
            (WORK_ORDERS_PATH, WORK_ORDER_COLUMNS),
            (INSPECTOR_BONUS_PATH, INSPECTOR_BONUS_COLUMNS),
            (HR_LEAVE_PATH, HR_LEAVE_COLUMNS),
            (HR_DAILY_RATES_PATH, HR_DAILY_RATES_COLUMNS),
        ]
        for path, columns in targets:
            sync_persistent_file(path, columns)
        st.session_state["drive_storage_initialised"] = True

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
DEFAULT_ROLES = ["Manager", "Staff", "Team Member", "Work Order Employee", "Work Order Manager", "Director", "Payroll", "Super Admin"]
DEFAULT_DEPARTMENTS = ["National Grid", "Isolator", "Project", "Accounts", "Payroll Department", "ACoole Electrical Ltd"]
EXCEL_COLUMNS = ["ID", "Employee Name", "Department", "Transaction Type", "Category Reason", "Date", "Amount (£)", "Line Manager", "Description", "Attachment Name", "Status", "Director Comments", "Decision Date", "Decision By", "Submitted By", "PDF File Path", "Edited From ID", "Old Data"]
WORK_ORDER_COLUMNS = ["Work Order ID", "Manual Work Order No.", "Employee Name", "Department", "Work Date", "Hours", "Amount (£)", "Manager", "Description", "Attachment Name", "Status", "Site Address", "Customer Job No.", "Manager Comments", "Manager Decision Date", "Manager Decision By", "Director Comments", "Director Decision Date", "Director Decision By", "Submitted By", "Submitted Date", "Payroll Status", "Payroll Date", "Payroll By", "PDF File Path"]
INSPECTOR_BONUS_COLUMNS = ["ID", "Inspector Name", "Month & Year", "Days Absent", "Reasons for Absence", "Total Jobs Completed", "Bonus Amount (£)", "Status", "Director Comments", "Director Decision Date", "Director Decision By", "Submitted By", "Submitted Date", "PDF File Path"]

DEFAULT_USERS = [
    {"full_name": "National Grid Manager", "username": "national_grid", "password": "acoole123", "role": "Manager", "dept": "National Grid", "can_access_inspector_bonus": True, "can_access_addition_deduction": True, "can_access_work_orders": False, "can_access_wo_total": False, "can_access_hr_leave": True, "is_active": True},
    {"full_name": "Isolator Manager", "username": "isolator", "password": "acoole123", "role": "Manager", "dept": "Isolator", "can_access_hr_leave": True, "is_active": True},
    {"full_name": "Project Manager", "username": "project", "password": "acoole123", "role": "Manager", "dept": "Project", "can_access_hr_leave": True, "is_active": True},
    {"full_name": "Accounts Manager", "username": "accounts", "password": "acoole123", "role": "Manager", "dept": "Accounts", "can_access_hr_leave": True, "is_active": True},
    {"full_name": "Andy Acoole", "username": "andy", "password": "andy2026", "role": "Director", "dept": "ACoole Electrical Ltd", "is_active": True},
    {"full_name": "System Administrator", "username": "wais", "password": "superadmin123", "role": "Super Admin", "dept": "System Administration", "is_active": True},
    {"full_name": "Payroll Team", "username": "payroll", "password": "payroll2026", "role": "Payroll", "dept": "Payroll Department", "is_active": True}
]
PERMISSION_DEFAULTS = {
    "Work Order Employee": {"can_view_all_dept": False, "can_generate_pdf": False, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False, "can_access_addition_deduction": False, "can_access_work_orders": True, "can_access_wo_total": False, "can_access_hr_leave": False},
    "Work Order Manager": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False, "can_access_addition_deduction": False, "can_access_work_orders": True, "can_access_wo_total": True, "can_access_hr_leave": False},
    "Staff": {"can_view_all_dept": False, "can_generate_pdf": False, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False, "can_access_addition_deduction": True, "can_access_work_orders": False, "can_access_wo_total": False, "can_access_hr_leave": False},
    "Team Member": {"can_view_all_dept": True, "can_generate_pdf": False, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False, "can_access_addition_deduction": True, "can_access_work_orders": False, "can_access_wo_total": False, "can_access_hr_leave": False},
    "Manager": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": True, "can_access_addition_deduction": True, "can_access_work_orders": False, "can_access_wo_total": False, "can_access_hr_leave": True},
    "Director": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True, "can_approve_requests": True, "can_access_inspector_bonus": True, "can_access_addition_deduction": True, "can_access_work_orders": True, "can_access_wo_total": True, "can_access_hr_leave": True},
    "Payroll": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True, "can_approve_requests": False, "can_access_inspector_bonus": True, "can_access_addition_deduction": True, "can_access_work_orders": True, "can_access_wo_total": True, "can_access_hr_leave": True},
    "Super Admin": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True, "can_approve_requests": True, "can_access_inspector_bonus": True, "can_access_addition_deduction": True, "can_access_work_orders": True, "can_access_wo_total": True, "can_access_hr_leave": True}
}
PERMISSION_LABELS = {
    "can_view_all_dept": "👁️ View All Department Requests",
    "can_generate_pdf": "📄 Generate & Download PDFs",
    "can_download_data": "📥 Download Data Backups",
    "can_approve_requests": "✅ Approve/Reject Requests",
    "can_access_inspector_bonus": "💰 National Grid Inspector Bonus",
    "can_access_addition_deduction": "➕ Addition & Deduction",
    "can_access_work_orders": "🛠️ Work Orders",
    "can_access_wo_total": "💷 Approved Work Order Total",
    "can_access_hr_leave": "👥 HR Leave Settlement"
}

try:
    from fpdf import FPDF
    PDF_AVAILABLE = True
except ImportError:
    FPDF = None
    PDF_AVAILABLE = False

_EXCEL_INIT_LOCK = threading.Lock()

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
        tmp_path = f"{path}.init.tmp"
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
            tmp_path = f"{path}.repair.tmp"
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
        tmp_path = f"{path}.init.tmp"
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
        }
        display_action = action_labels.get(action, action)
        old_val = json.dumps(old_data, ensure_ascii=False)[:300] if old_data else "-"
        new_val = json.dumps(new_data, ensure_ascii=False)[:300] if new_data else "-"
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
        old_v = json.dumps(old_data, ensure_ascii=False)[:300] if old_data else "-"
        new_v = json.dumps(new_data, ensure_ascii=False)[:300] if new_data else "-"
        save_audit_entry({"AuditID": _get_next_audit_id(), "Timestamp": timestamp, "User_Name": username, "User_Role": role, "Action": wo_labels.get(action, action), "Request_ID": str(req_id), "Department": dept, "Amount": amount, "Decision_By": final_decision_by or "-", "Decision_Date": final_decision_date or timestamp, "Field_Changed": "Work Order Status", "Old_Value": old_v if old_data else "-", "New_Value": new_v if new_data else action.replace("WORK_ORDER_", "").replace("_", " ").title(), "IP_Address": "Auto-Logged"})
        return
    if action.startswith("HR_LEAVE_"):
        hr_labels = {
            "HR_LEAVE_CREATED": "👥 HR Leave Settlement Created",
            "HR_LEAVE_APPROVED": "👥 HR Leave Settlement Approved",
            "HR_LEAVE_REJECTED": "👥 HR Leave Settlement Rejected",
            "HR_LEAVE_STATUS_CHANGED": "👥 HR Leave Settlement Status Changed",
        }
        old_v = json.dumps(old_data, ensure_ascii=False)[:300] if old_data else "-"
        new_v = json.dumps(new_data, ensure_ascii=False)[:300] if new_data else "-"
        save_audit_entry({"AuditID": _get_next_audit_id(), "Timestamp": timestamp, "User_Name": username, "User_Role": role, "Action": hr_labels.get(action, action), "Request_ID": str(req_id), "Department": "-", "Amount": "-", "Decision_By": final_decision_by or "-", "Decision_Date": final_decision_date or timestamp, "Field_Changed": "HR Leave Status", "Old_Value": old_v, "New_Value": new_v, "IP_Address": "Auto-Logged"})
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
                for path in (EXCEL_PATH, USER_DB_PATH, SETTINGS_PATH, AUDIT_LOG_PATH, INSPECTOR_BONUS_PATH, WORK_ORDERS_PATH, HR_LEAVE_PATH, HR_DAILY_RATES_PATH):
                    remote = _drive_find_file(os.path.basename(path))
                    if remote: _drive_download_file(remote["id"], path)
            _invalidate_data_cache("_records_cache", "_users_cache", "_settings_cache", "_audit_log_cache", "_work_orders_cache", "_inspector_bonus_cache", "_audit_log_count", "_hr_leave_cache", "_hr_daily_rates_cache")
            st.session_state["_last_refresh"] = datetime.now().isoformat()
        st.rerun()

def make_request_title(req):
    status, amount, dt, dec_by, decision_dt = req["status"].upper(), f"£{req['amount']:.2f}", format_date(req.get("date", "")), req.get("decision_by", ""), format_date(req.get("decision_date", ""))
    if req["status"] == "pending": return f"🟡 ID #{req['id']} | {req['emp_name']} | PENDING | {amount} | 📅 {dt}"
    elif req["status"] == "approved": return f"🟢 ID #{req['id']} | {req['emp_name']} | APPROVED | {amount} | ✅ Approved by {dec_by} on {decision_dt}"
    elif req["status"] == "rejected": return f"🔴 ID #{req['id']} | {req['emp_name']} | REJECTED | {amount} | ❌ Rejected by {dec_by} on {decision_dt}"
    else: return f"⚪ ID #{req['id']} | {req['emp_name']} | {status} | {amount} | 📅 {dt}"

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

# HR Leave Settings
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
                "is_active": _active_or_default(r.get("is_active", ""))
            }
            if user_role == "Super Admin":
                users[username].update({
                    "can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True,
                    "can_approve_requests": True, "can_access_inspector_bonus": True,
                    "can_access_addition_deduction": True, "can_access_work_orders": True,
                    "can_access_wo_total": True, "can_access_hr_leave": True, "is_active": True
                })
        _set_data_cache("_users_cache", users)
        return dict(users)
    except Exception as e:
        st.error(f"User DB Load Error: {e}")
        return {}

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
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
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
                    with csave: save_edit = st.form_submit_button("💾 Save Changes & Resubmit", type="primary", use_container_width=True)
                    with ccancel: cancel_edit = st.form_submit_button("❌ Cancel", use_container_width=True)
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
            submitted = st.form_submit_button("📤 Submit Work Order to Manager", type="primary", use_container_width=True)
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
                st.dataframe(df_view, use_container_width=True, hide_index=True)
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
                with csave: save_edit = st.form_submit_button("💾 Save Changes", type="primary", use_container_width=True)
                with ccancel: cancel_edit = st.form_submit_button("❌ Cancel", use_container_width=True)
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
            submitted = st.form_submit_button("📤 Submit Work Order", type="primary", use_container_width=True)
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
                f"🟡 {get_work_order_number(r)} | {r.get('emp_name')} | "
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
                f"🟢 {get_work_order_number(r)} | {r.get('emp_name')} | "
                f"£{r.get('amount',0):.2f} | Approved by {r.get('director_decision_by','')}"
            ):
                show_full_details(r)
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
                f"🔴 {get_work_order_number(r)} | {r.get('emp_name')} | "
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
                with c_save: save_edit = st.form_submit_button("💾 Save Changes & Resubmit", type="primary", use_container_width=True)
                with c_cancel: cancel_edit = st.form_submit_button("❌ Cancel", use_container_width=True)
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
        submitted = st.form_submit_button("📤 Submit for Director Approval", type="primary", use_container_width=True)
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
            with st.expander(f"🟡 {rec_id} | {r.get('inspector_name')} | {r.get('month_year')} | £{r.get('bonus_amount',0):.2f}"):
                show_details(r); st.divider()
                comments = st.text_area("Director Comments", key=f"ib_dir_comm_{rec_id}")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("✅ Approve", key=f"ib_dir_app_{rec_id}", type="primary"):
                        apply_decision(rec_id, "approved", comments, "pending_director"); st.rerun()
                with c2:
                    if st.button("❌ Reject", key=f"ib_dir_rej_{rec_id}"):
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
            })
        _set_data_cache("_hr_leave_cache", records)
        return list(records)
    except Exception as e:
        st.error(f"HR Leave Load Error: {e}")
        return []

def save_all_hr_leave(records, sync=True):
    rows = [{
        "ID": int(r.get("id", 0)),
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
    } for r in records]
    pd.DataFrame(rows, columns=HR_LEAVE_COLUMNS).to_excel(HR_LEAVE_PATH, index=False, engine="openpyxl")
    _set_data_cache("_hr_leave_cache", list(records))
    if sync: sync_saved_file_to_drive(HR_LEAVE_PATH)

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
# 👥 HR LEAVE SETTLEMENT — UI
# ============================================================
def render_hr_leave_form(user_name):
    st.subheader("👥 New Request — HR Leave Settlement")
    st.caption("Submit a holiday settlement (payout / deduction) for Director approval.")
    hr_records = load_hr_leave()
    hr_cats = load_hr_categories()
    departments = load_departments()

    form_version = st.session_state.get("hr_leave_form_version", 0)
    with st.form(f"hr_leave_form_v{form_version}", clear_on_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            emp_name = st.text_input("👤 Employee Name", key=f"hr_emp_v{form_version}")
            owe_owed = st.selectbox("⚖️ Owe / Owed", DEFAULT_OWE_OWED, key=f"hr_owe_v{form_version}")
            category = st.selectbox("🏷️ Category / Reason", hr_cats, key=f"hr_cat_v{form_version}")
            transaction_type = "Addition" if owe_owed == "Company Owes Employee" else "Deduction"
            st.text_input("🔄 Transaction Type (auto-linked)", value=transaction_type, disabled=True, key=f"hr_type_v{form_version}")
        with c2:
            dt_val = st.date_input("📅 Date", value=date.today(), key=f"hr_date_v{form_version}")
            emp_dept = st.selectbox("🏢 Employee Department", departments, key=f"hr_dept_v{form_version}")
            manager = st.text_input("👔 Line Manager", key=f"hr_mgr_v{form_version}")
            files = st.file_uploader("📎 Attachments", type=["pdf", "png", "jpg", "jpeg"],
                                     accept_multiple_files=True, key=f"hr_files_v{form_version}")

        c3, c4 = st.columns(2)
        with c3:
            num_days = st.number_input("🔢 Number of Days", min_value=0.0, step=0.5,
                                       format="%.2f", key=f"hr_days_v{form_version}")
        with c4:
            rate = get_hr_daily_rate(emp_dept, transaction_type)
            suggested = round((rate or 0) * float(num_days), 2)
            amount = st.number_input("💷 Amount (£)", min_value=0.0, step=1.0, format="%.2f",
                                     value=suggested, key=f"hr_amt_v{form_version}")
            if rate:
                st.caption(f"ℹ️ Auto-calc: {num_days} × £{rate:.2f} = £{suggested:.2f} (editable)")
            else:
                st.caption("⚠️ No daily rate configured by Super Admin — please enter manually.")
        desc = st.text_area("📝 Description / Justification", key=f"hr_desc_v{form_version}")
        submitted = st.form_submit_button("📤 Send to Director", type="primary", use_container_width=True)

    if submitted:
        if not emp_name.strip() or not manager.strip() or not desc.strip():
            st.error("⚠️ Employee Name, Line Manager and Description are required.")
        elif float(num_days) <= 0:
            st.error("⚠️ Number of Days must be greater than 0.")
        elif float(amount) <= 0:
            st.error("⚠️ Amount must be greater than 0.")
        else:
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
                "id": new_id, "emp_name": emp_name.strip(), "emp_dept": emp_dept,
                "type": transaction_type, "category": category, "owe_owed": owe_owed,
                "date": str(dt_val), "days": float(num_days), "amount": float(amount),
                "manager": manager.strip(), "desc": desc.strip(),
                "attachment_name": ", ".join(attachments) or "None",
                "status": "pending", "director_comments": "", "rejection_reason": "",
                "decision_date": "", "decision_by": "",
                "submitted_by": user_name, "submitted_date": now, "pdf_path": "",
            }
            hr_records.append(rec)
            save_all_hr_leave(hr_records)
            log_action("HR_LEAVE_CREATED", new_id, new_data=rec)
            st.session_state["hr_leave_form_version"] = form_version + 1
            st.success(f"✅ HR Leave Settlement #{new_id} sent to Director for approval.")
            st.rerun()

    st.divider()
    st.subheader("📋 My Submitted HR Leave Requests")
    mine = [r for r in hr_records if r.get("submitted_by") == user_name]
    if not mine:
        st.info("📋 You have not submitted any HR Leave requests yet.")
    else:
        for r in reversed(mine):
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

def render_hr_leave_director_portal(director_name):
    st.subheader("👥 HR Leave Settlement — Director Approval")
    st.info("Review HR leave settlement requests. Rejection requires a mandatory reason.")
    st.divider()
    records = load_hr_leave()
    pending  = [r for r in records if r["status"] == "pending"]
    approved = [r for r in records if r["status"] == "approved"]
    rejected = [r for r in records if r["status"] == "rejected"]
    t1, t2, t3 = st.tabs([f"⏳ Pending ({len(pending)})", f"✅ Approved ({len(approved)})", f"❌ Rejected ({len(rejected)})"])

    def show_details(r):
        st.write(f"👤 **Employee:** {r['emp_name']} | 🏢 **Department:** {r['emp_dept']}")
        st.write(f"🔄 **Transaction Type:** {r['type']} | ⚖️ **Owe / Owed:** {r['owe_owed']}")
        st.write(f"🏷️ **Category / Reason:** {r['category']} | 📅 **Date:** {r['date']}")
        st.write(f"🔢 **Days:** {r['days']} | 💷 **Amount:** £{r['amount']:.2f}")
        st.write(f"👔 **Line Manager:** {r['manager']}")
        st.write(f"📝 **Submitted by:** {r['submitted_by']} on {r['submitted_date']}")
        st.info(f"📝 **Description:**\n{r['desc']}")
        st.divider(); st.markdown("#### 📎 Attachments"); display_attachments(r)

    with t1:
        if not pending: st.success("✅ No pending HR Leave requests.")
        for r in reversed(pending):
            rid = r["id"]
            with st.expander(f"🟡 #{rid} | {r['emp_name']} | {r['emp_dept']} | £{r['amount']:.2f} | {r['type']}"):
                show_details(r)
                st.markdown("### ✍️ Director Decision")
                comments = st.text_area("Director Comments (optional)", key=f"hr_dir_comm_{rid}")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("✅ Approve", key=f"hr_dir_app_{rid}", type="primary", use_container_width=True):
                        for x in records:
                            if x["id"] == rid:
                                x["status"] = "approved"; x["director_comments"] = comments.strip()
                                x["decision_by"] = director_name
                                x["decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                break
                        save_all_hr_leave(records)
                        log_action("HR_LEAVE_APPROVED", rid, decision_by=director_name)
                        st.success(f"✅ HR Leave #{rid} Approved."); st.rerun()
                with c2:
                    if st.button("❌ Reject", key=f"hr_dir_rej_{rid}", use_container_width=True):
                        st.session_state[f"hr_reject_modal_{rid}"] = True
                if st.session_state.get(f"hr_reject_modal_{rid}"):
                    st.warning("⚠️ A rejection reason is required.")
                    reason = st.text_area("Reason for Rejection (required)", key=f"hr_rej_reason_{rid}")
                    rc1, rc2 = st.columns(2)
                    with rc1:
                        if st.button("Confirm Rejection", key=f"hr_rej_confirm_{rid}", type="primary", use_container_width=True):
                            if not reason.strip():
                                st.error("❌ Rejection reason cannot be empty.")
                            else:
                                for x in records:
                                    if x["id"] == rid:
                                        x["status"] = "rejected"; x["rejection_reason"] = reason.strip()
                                        x["director_comments"] = comments.strip()
                                        x["decision_by"] = director_name
                                        x["decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                        break
                                save_all_hr_leave(records)
                                log_action("HR_LEAVE_REJECTED", rid, decision_by=director_name)
                                st.session_state[f"hr_reject_modal_{rid}"] = False
                                st.success(f"❌ HR Leave #{rid} Rejected."); st.rerun()
                    with rc2:
                        if st.button("Cancel", key=f"hr_rej_cancel_{rid}", use_container_width=True):
                            st.session_state[f"hr_reject_modal_{rid}"] = False; st.rerun()

    with t2:
        if not approved: st.info("✅ No approved HR Leave requests.")
        for r in reversed(approved):
            with st.expander(f"🟢 #{r['id']} | {r['emp_name']} | £{r['amount']:.2f} | ✅ {r['decision_by']}"):
                show_details(r)
                if r.get("director_comments"): st.info(f"💬 Director Comments: {r['director_comments']}")
                st.write(f"📅 Decision Date: {r['decision_date']}")

    with t3:
        if not rejected: st.info("❌ No rejected HR Leave requests.")
        for r in reversed(rejected):
            with st.expander(f"🔴 #{r['id']} | {r['emp_name']} | £{r['amount']:.2f} | ❌ {r['decision_by']}"):
                show_details(r)
                st.error(f"❌ Rejection Reason: {r['rejection_reason']}")
                if r.get("director_comments"): st.info(f"💬 Director Comments: {r['director_comments']}")

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
    with t3:
        if not rejected: st.info("❌ No rejected HR Leave requests.")
        for r in reversed(rejected):
            with st.expander(f"🔴 #{r['id']} | {r['emp_name']} | £{r['amount']:.2f}"):
                show(r); st.error(f"❌ Reason: {r['rejection_reason']}")

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

def render_hr_leave_settings():
    """Super Admin settings: HR categories + Daily Holiday Rates."""
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

def initialise_excel():
    safe_init_excel(EXCEL_PATH, EXCEL_COLUMNS)

initialise_excel()
initialise_work_orders()
initialise_inspector_bonus()
initialise_hr_leave()

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
                st.image(path, caption=name, use_container_width=True)
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

def user_management_panel():
    st.subheader("👤 User Management — Create & Manage System Users")
    st.info("🛡️ Super Admin Only — Create, edit, or delete user accounts."); st.divider()
    USERS = load_users(); ROLES = load_roles()
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
            perm_inspector = st.checkbox(PERMISSION_LABELS["can_access_inspector_bonus"], value=defaults.get("can_access_inspector_bonus", False))
            new_dept = st.selectbox("🏢 Department", load_departments())
            new_active = st.checkbox("✅ Account Active", value=True, help="Uncheck to block this user from logging in.")
            if st.form_submit_button("✅ Create User Account", type="primary"):
                if not new_full_name.strip() or not new_username or not new_password: st.error("❌ All fields required!")
                elif new_username in USERS: st.error(f"❌ Username '{new_username}' already exists!")
                else:
                    USERS[new_username] = {"full_name": new_full_name.strip(), "password": new_password, "role": new_role, "dept": new_dept, "can_view_all_dept": perm_view_all, "can_generate_pdf": perm_pdf, "can_download_data": perm_download, "can_approve_requests": perm_approve, "can_access_inspector_bonus": perm_inspector, "can_access_addition_deduction": perm_ad, "can_access_work_orders": perm_wo, "can_access_wo_total": perm_wo_total, "can_access_hr_leave": perm_hr, "is_active": new_active}
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
                edit_ib = st.checkbox(PERMISSION_LABELS["can_access_inspector_bonus"], value=curr_perm_ib)
                edit_active = st.checkbox("✅ Account Active", value=curr.get("is_active", True), help="Uncheck to block this user from logging in.")
                if st.form_submit_button("🔄 Update User", type="primary"):
                    USERS = load_users()
                    if upd_username_new != edit_user_sel:
                        if upd_username_new in USERS: st.error(f"❌ Username '{upd_username_new}' already exists!"); return
                        USERS[upd_username_new] = {"full_name": upd_full_name.strip(), "password": upd_password if upd_password else curr["password"], "role": upd_role, "dept": upd_dept, "can_view_all_dept": edit_view, "can_generate_pdf": edit_pdf, "can_download_data": edit_dl, "can_approve_requests": edit_app, "can_access_inspector_bonus": edit_ib, "can_access_addition_deduction": edit_ad, "can_access_work_orders": edit_wo, "can_access_wo_total": edit_wo_total, "can_access_hr_leave": edit_hr, "is_active": edit_active}
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
        if st.form_submit_button("🔐 Authenticate Portal", type="primary", use_container_width=True):
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
# 📋 ROLE-BASED PORTALS
# ============================================================
if role == "Work Order Employee":
    render_work_order_employee_portal(full_name, dept)

elif role == "Payroll":
    st.subheader("🧾 Payroll Portal")
    st.info("✅ View all requests and Download PDFs.")
    st.divider()
    tab_add_ded, tab_hr_leave, tab_work_orders, tab_inspector_bonus = st.tabs([
        "➕ Addition & Deduction",
        "👥 HR Leave Settlement",
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
                                r["emp_name"] = en.strip(); r["type"] = rt; r["category"] = ct; r["amount"] = amt; r["date"] = str(dt_val); r["manager"] = mgr.strip(); r["desc"] = desc.strip(); r["status"] = "pending"; r["attachment_name"] = ", ".join(final_attachments) or "None"; r["old_data"] = json.dumps(old_data_dict)
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
            with st.form(f"wo_mgr_new_req_v{form_version}", clear_on_submit=False):
                c1, c2 = st.columns(2)
                with c1:
                    en = st.text_input("👤 Employee Name", key=f"wo_mgr_en_v{form_version}")
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
                        payload = {"id": nid, "emp_name": en.strip(), "dept": dept_name, "type": rt, "category": ct, "date": str(dt_val), "amount": amt, "manager": mgr.strip(), "desc": desc.strip(), "attachment_name": ", ".join(att_list) or "None", "status": "pending", "director_comments": "", "decision_date": "", "decision_by": "", "submitted_by": full_name, "pdf_path": "", "edited_from_id": "", "old_data": ""}
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
    labels = []
    if has_addition_deduction: labels.append("➕ Addition & Deduction")
    if has_hr_leave: labels.append("👥 HR Leave Settlement")
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
                                        r["emp_name"] = en.strip(); r["type"] = rt; r["category"] = ct; r["amount"] = amt; r["date"] = str(dt_val); r["manager"] = mgr.strip(); r["desc"] = desc.strip(); r["status"] = "pending"; r["attachment_name"] = ", ".join(final_attachments) or "None"; r["old_data"] = json.dumps(old_data_dict)
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
                    with st.form(f"new_req_v{form_version}", clear_on_submit=False):
                        c1, c2 = st.columns(2)
                        with c1:
                            en = st.text_input("👤 Employee Name", key=f"en_v{form_version}")
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
                                payload = {"id": nid, "emp_name": en.strip(), "dept": dept_name, "type": rt, "category": ct, "date": str(dt_val), "amount": amt, "manager": mgr.strip(), "desc": desc.strip(), "attachment_name": ", ".join(att_list) or "None", "status": "pending", "director_comments": "", "decision_date": "", "decision_by": "", "submitted_by": full_name, "pdf_path": "", "edited_from_id": "", "old_data": ""}
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
                render_hr_leave_form(full_name)
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
    director_addition_tab, director_hr_tab, director_work_order_tab, director_inspector_tab = st.tabs([
        "➕ Addition & Deduction",
        "👥 HR Leave Settlement",
        "🛠️ Work Orders",
        "💰 National Grid Inspector Bonus"
    ])
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
                    with st.expander(f"🟡 ID #{req_id} | {req.get('emp_name')} | £{float(req.get('amount',0)):.2f} | {req.get('dept')}"):
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
                                records = load_records_from_excel()
                                for r in records:
                                    if int(r.get("id",0)) == int(req_id):
                                        r["status"] = "rejected"; r["decision_by"] = full_name; r["director_comments"] = comments; r["decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
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
                    with st.expander(f"🟢 ID #{req_id} | {req.get('emp_name')} | £{float(req.get('amount',0)):.2f}"):
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
                                records = load_records_from_excel()
                                for r in records:
                                    if int(r.get("id",0)) == int(req_id):
                                        r["status"] = "rejected"; r["decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S"); r["decision_by"] = full_name
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
                    with st.expander(f"🔴 ID #{req_id} | {req.get('emp_name')} | £{float(req.get('amount',0)):.2f}"):
                        st.write(f"👤 {req.get('emp_name')} | 💷 £{float(req.get('amount',0)):.2f}")
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
    with director_hr_tab:
        render_hr_leave_director_portal(full_name)
    with director_work_order_tab: render_work_order_director_portal(full_name)
    with director_inspector_tab: render_inspector_bonus_director_portal(full_name)

elif role == "Super Admin":
    super_add_ded_tab, super_hr_tab, super_work_orders_tab, super_inspector_bonus_tab, super_system_mgmt_tab = st.tabs([
        "➕ Addition & Deduction",
        "👥 HR Leave Settlement",
        "🛠️ Work Orders",
        "💰 National Grid Inspector Bonus",
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
    with super_hr_tab:
        render_hr_leave_super_admin()
    with super_work_orders_tab:
        render_work_orders_super_admin()
    with super_inspector_bonus_tab:
        render_inspector_bonus_super_admin()
    with super_system_mgmt_tab:
        st.subheader("🔧 System Management — Super Admin")
        st.info("🛡️ Manage system settings, users, audit history and data reset controls.")
        st.divider()
        tab_settings, tab_hr_settings, tab_users, tab_audit = st.tabs([
            "⚙️ System Settings",
            "👥 HR Leave Settings",
            "👤 User Management",
            "📖 Audit History"
        ])
        with tab_settings: settings_management_panel()
        with tab_hr_settings: render_hr_leave_settings()
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
                        if st.button("💰 Clear Inspector Bonuses", key="super_admin_clear_all_inspector_bonus", type="secondary", use_container_width=True):
                            st.session_state["confirm_clear_inspector_bonus"] = True
                    else:
                        if st.button("✅ Confirm", key="super_admin_confirm_clear_inspector_bonus", use_container_width=True):
                            clear_all_inspector_bonus()
                            log_action("SUPER_ADMIN_CLEAR_INSPECTOR_BONUS", "ALL", decision_by=full_name)
                            st.session_state["confirm_clear_inspector_bonus"] = False
                            st.success("✅ Cleared."); st.rerun()
                with reset_col2:
                    if not st.session_state.get("confirm_clear_all_work_orders", False):
                        if st.button("🛠️ Clear Work Orders", key="super_admin_clear_all_work_orders", type="secondary", use_container_width=True):
                            st.session_state["confirm_clear_all_work_orders"] = True
                    else:
                        if st.button("✅ Confirm", key="super_admin_confirm_clear_all_work_orders", use_container_width=True):
                            save_all_work_orders([])
                            st.session_state["confirm_clear_all_work_orders"] = False
                            st.success("✅ Cleared."); st.rerun()
                with reset_col3:
                    if not st.session_state.get("confirm_clear_hr_leave", False):
                        if st.button("👥 Clear HR Leave", key="super_admin_clear_all_hr_leave", type="secondary", use_container_width=True):
                            st.session_state["confirm_clear_hr_leave"] = True
                    else:
                        if st.button("✅ Confirm", key="super_admin_confirm_clear_all_hr_leave", use_container_width=True):
                            clear_all_hr_leave()
                            st.session_state["confirm_clear_hr_leave"] = False
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
            ("📥 Audit Log", AUDIT_LOG_PATH, "audit_log", "backup_audit_log"),
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
            _invalidate_data_cache(
                "_work_orders_cache",
                "_work_order_cache",
                "_inspector_bonus_cache",
                "_records_cache",
                "_hr_leave_cache",
            )
            sync_saved_file_to_drive(WORK_ORDERS_PATH)
            sync_saved_file_to_drive(HR_LEAVE_PATH)
            clear_audit_log_file()

            for folder in (PDF_DIR, WORK_ORDER_PDF_DIR, INSPECTOR_BONUS_PDF_DIR, HR_LEAVE_PDF_DIR, UPLOAD_DIR):
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
                "work orders, inspector bonuses, HR leave requests, audit history, generated PDFs and uploaded attachments. "
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
                        st.session_state["confirm_live_launch"] = False
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
