# ============================================================
# 🔄 ACOOLE PORTAL — PROFESSIONAL VERSION v4.9
# ============================================================
# ✅ v4.9 (PERFORMANCE):
#    • Google Drive uploads now run in the BACKGROUND — UI no longer waits
#      for the network on post / edit / approve / PDF generation.
#    • Drive file-ID cache skips the slow "list files" query on every upload.
#    • Sync fingerprinting skips uploading files that haven't changed.
#    • Startup download skips files whose LOCAL copy is newer than Drive's.
# ✅ v4.8:
#    • Employee Work Order portal now has TWO tabs:
#         📋 My Work Orders   |   🏢 Department Work Orders (own dept, incl. manager-submitted)
#    • Employee can see their department's work orders before submitting,
#      preventing duplicate Work Order No. submissions.
#    • On duplicate Work Order No. submission, the form data is KEPT
#      (no clearing) and the message "Work Order No. already Exist." is shown,
#      so the user can change only the Work Order No. and resubmit.
#    • Form is only cleared on SUCCESSFUL submission (not on duplicate).
# ✅ v4.7:
#    • Employee Work Order form matches Manager's layout
#    • Employee can edit pending / returned / rejected work orders
#    • Manager's Pending tab now shows "Submitted by" for each work order
# ✅ v4.6:
#    • Work Order Manager can APPROVE / REJECT employee-submitted work orders
#    • Edit Work Order button now works
#    • Super Admin can Clear All Inspector Bonus records
# ✅ All prior fixes retained
# ============================================================
import streamlit as st
import os
import sys
import json
import shutil
import subprocess
import pandas as pd
import io
import requests
import threading
from datetime import datetime, date

from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2.credentials import Credentials

# ============================================================
# ─── PAGE CONFIG — MUST BE FIRST! ───
# ============================================================
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

# ============================================================
# ✅ ALL CONFIGURATION — DEFINED FIRST!
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
APP_FOLDER = os.path.join(BASE_DIR, "Acoole_App_Uploads")
os.makedirs(APP_FOLDER, exist_ok=True)
UPLOAD_DIR = os.path.join(APP_FOLDER, "uploaded_attachments")
AUDIT_LOG_PATH = os.path.join(APP_FOLDER, "audit_log.xlsx")
AUDIT_LOG_FILE = AUDIT_LOG_PATH
AUDIT_COLUMNS = [
    "AuditID", "Timestamp", "User_Name", "User_Role",
    "Action", "Request_ID", "Department", "Amount",
    "Decision_By", "Decision_Date", "Field_Changed",
    "Old_Value", "New_Value", "IP_Address"
]
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
# ─── GOOGLE DRIVE ───
GOOGLE_DRIVE_FOLDER_ID = "1g3DsqT_w_tU0QBnrXcZqYjp51SokH4hG"

# ─── ONEDRIVE / MICROSOFT GRAPH ───
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

# ============================================================
# GOOGLE DRIVE — YOUR PERSONAL GOOGLE ACCOUNT
# ============================================================
SCOPES = ["https://www.googleapis.com/auth/drive"]
drive_service = None

# ─── Background sync caches (module-level, safe to share across threads) ───
_DRIVE_ID_CACHE = {}            # "parent_id::filename"  ->  Drive file_id
_DRIVE_SYNC_FINGERPRINTS = {}   # local_path             ->  "size:mtime"
_DRIVE_SYNC_LOCK = threading.Lock()

try:
    gdrive = st.secrets["gdrive"]
    credentials = Credentials(
        token=None,
        refresh_token=gdrive["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=gdrive["client_id"],
        client_secret=gdrive["client_secret"],
        scopes=SCOPES
    )
    drive_service = build("drive", "v3", credentials=credentials, cache_discovery=False)
    about = drive_service.about().get(fields="user").execute()
    st.success("✅ Google Drive connected")
except Exception as e:
    drive_service = None
    st.error(f"❌ Google Drive connection failed: {e}")

# ============================================================
# TEST GOOGLE DRIVE FOLDER
# ============================================================
if drive_service:
    try:
        folder = drive_service.files().get(
            fileId=GOOGLE_DRIVE_FOLDER_ID,
            fields="id,name,mimeType"
        ).execute()
        st.success("✅ Google Drive folder accessible")
    except Exception as e:
        drive_service = None
        st.error(f"❌ Google Drive folder access failed: {e}")

# ============================================================
# GOOGLE DRIVE UPLOAD
# ============================================================
def upload_to_google_drive(local_file_path, display_filename):
    """Upload/update a file on Google Drive. Reuses a cached file ID so we
    don't have to run a slow 'list files' query on every upload."""
    if drive_service is None or not os.path.exists(local_file_path):
        return None
    cache_key = f"{GOOGLE_DRIVE_FOLDER_ID}::{display_filename}"

    def _do(file_id=None):
        media = MediaFileUpload(local_file_path, resumable=False)
        if file_id:
            return drive_service.files().update(
                fileId=file_id, media_body=media, fields="id,name,parents"
            ).execute()
        metadata = {"name": display_filename, "parents": [GOOGLE_DRIVE_FOLDER_ID]}
        return drive_service.files().create(
            body=metadata, media_body=media, fields="id,name,parents"
        ).execute()

    try:
        cached_id = _DRIVE_ID_CACHE.get(cache_key)
        if cached_id:
            try:
                return _do(cached_id).get("id")
            except Exception:
                _DRIVE_ID_CACHE.pop(cache_key, None)

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

# ============================================================
# GOOGLE DRIVE — PERMANENT APPLICATION DATA STORAGE
# ============================================================
def _drive_find_file(filename, parent_id=GOOGLE_DRIVE_FOLDER_ID):
    if drive_service is None:
        return None
    try:
        safe_name = str(filename).replace("'", "\\'")
        q = (f"name = '{safe_name}' and '{parent_id}' in parents and trashed = false")
        result = drive_service.files().list(
            q=q, spaces="drive",
            fields="files(id,name,modifiedTime)",
            orderBy="modifiedTime desc", pageSize=10
        ).execute()
        files = result.get("files", [])
        return files[0] if files else None
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
            return drive_service.files().update(
                fileId=file_id, media_body=media, fields="id,name"
            ).execute()
        metadata = {"name": filename, "parents": [parent_id]}
        return drive_service.files().create(
            body=metadata, media_body=media, fields="id,name"
        ).execute()

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
        print(f"Drive upload failed for {filename}: {e}")
        return None

def _drive_download_file(file_id, local_path):
    if drive_service is None:
        return False
    try:
        request = drive_service.files().get_media(fileId=file_id)
        with open(local_path, "wb") as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
        return True
    except Exception as e:
        print(f"Drive download failed for {local_path}: {e}")
        return False

def sync_persistent_file(local_path, columns=None):
    if drive_service is None:
        return
    filename = os.path.basename(local_path)
    remote = _drive_find_file(filename)
    if remote:
        try:
            rmt = str(remote.get("modifiedTime", "")).rstrip("Z")
            local_newer = False
            if os.path.exists(local_path) and rmt:
                if "." in rmt:
                    remote_dt = datetime.strptime(rmt, "%Y-%m-%dT%H:%M:%S.%f")
                else:
                    remote_dt = datetime.strptime(rmt, "%Y-%m-%dT%H:%M:%S")
                local_dt = datetime.utcfromtimestamp(os.path.getmtime(local_path))
                if local_dt > remote_dt:
                    local_newer = True
            if not local_newer:
                if not _drive_download_file(remote["id"], local_path):
                    print(f"Using local copy of {filename} (Drive download failed)")
        except Exception as e:
            print(f"mtime compare failed for {filename} ({e}); downloading anyway")
            _drive_download_file(remote["id"], local_path)
    elif os.path.exists(local_path):
        _drive_upload_path(local_path, filename)
    elif columns is not None:
        pd.DataFrame(columns=columns).to_excel(local_path, index=False, engine="openpyxl")
        _drive_upload_path(local_path, filename)

def sync_saved_file_to_drive(local_path):
    """Non-blocking upload — spawns a background thread so the UI is never
    blocked waiting on Google Drive. Also skips the upload if the file has
    not changed since the last successful sync."""
    if drive_service is None or not os.path.exists(local_path):
        return
    try:
        st_info = os.stat(local_path)
        fingerprint = f"{st_info.st_size}:{int(st_info.st_mtime)}"
    except Exception:
        return
    if _DRIVE_SYNC_FINGERPRINTS.get(local_path) == fingerprint:
        return
    _DRIVE_SYNC_FINGERPRINTS[local_path] = fingerprint

    def _worker(path):
        with _DRIVE_SYNC_LOCK:
            try:
                if _drive_upload_path(path) is None:
                    _DRIVE_SYNC_FINGERPRINTS.pop(path, None)
            except Exception as e:
                print(f"Background drive sync failed for {path}: {e}")
                _DRIVE_SYNC_FINGERPRINTS.pop(path, None)

    threading.Thread(target=_worker, args=(local_path,), daemon=True).start()

def _upload_to_drive_bg(local_path, filename):
    """Fire-and-forget upload for one-off files (attachments, generated PDFs).
    The local file is kept on disk so downloads still work immediately."""
    if drive_service is None or not os.path.exists(local_path):
        return
    def _worker(path, name):
        try:
            upload_to_google_drive(path, name)
        except Exception as e:
            print(f"Background upload failed for {name}: {e}")
    threading.Thread(target=_worker, args=(local_path, filename), daemon=True).start()

def initialise_drive_storage():
    if drive_service is None or st.session_state.get("drive_storage_initialised"):
        return
    targets = [
        (EXCEL_PATH, EXCEL_COLUMNS),
        (USER_DB_PATH, ["full_name","username","password","role","dept",
            "can_view_all_dept","can_generate_pdf","can_download_data","can_approve_requests","can_access_inspector_bonus"]),
        (SETTINGS_PATH, ["setting", "value"]),
        (AUDIT_LOG_PATH, AUDIT_COLUMNS),
        (WORK_ORDERS_PATH, WORK_ORDER_COLUMNS),
        (INSPECTOR_BONUS_PATH, INSPECTOR_BONUS_COLUMNS),
    ]
    for path, columns in targets:
        sync_persistent_file(path, columns)
    st.session_state["drive_storage_initialised"] = True

# ============================================================
# ✅ ONEDRIVE UPLOAD FUNCTIONS
# ============================================================
def get_onedrive_token():
    if not ONEDRIVE_CLIENT_ID or not ONEDRIVE_CLIENT_SECRET:
        return None
    try:
        url = f"https://login.microsoftonline.com/{ONEDRIVE_TENANT_ID}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials", "client_id": ONEDRIVE_CLIENT_ID,
            "client_secret": ONEDRIVE_CLIENT_SECRET, "scope": "https://graph.microsoft.com/.default"
        }
        res = requests.post(url, data=data, timeout=30)
        if res.status_code == 200:
            return res.json().get("access_token")
    except Exception as e:
        st.warning(f"⚠️ OneDrive connection: {e}")
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
        if res.status_code in (200, 201):
            return True
    except Exception:
        pass
    return False

# ============================================================
# DEFAULTS — ROLES, DEPARTMENTS, USERS, PERMISSIONS
# ============================================================
DEFAULT_CATEGORIES = ["Food Allowance", "Others", "Parking", "Parking Fine", "GYM Membership", "Item Not Returned", "Item Missing"]
DEFAULT_ROLES = ["Manager", "Staff", "Team Member", "Work Order Employee", "Work Order Manager", "Director", "Payroll", "Super Admin"]
DEFAULT_DEPARTMENTS = ["National Grid", "Isolator", "Project", "Accounts", "Payroll Department", "ACoole Electrical Ltd"]
EXCEL_COLUMNS = [
    "ID", "Employee Name", "Department", "Transaction Type", "Category Reason",
    "Date", "Amount (£)", "Line Manager", "Description", "Attachment Name",
    "Status", "Director Comments", "Decision Date", "Decision By",
    "Submitted By", "PDF File Path", "Edited From ID", "Old Data"
]
WORK_ORDER_COLUMNS = [
    "Work Order ID", "Manual Work Order No.", "Employee Name", "Department", "Work Date", "Hours",
    "Amount (£)", "Manager", "Description", "Attachment Name", "Status",
    "Site Address", "Customer Job No.",
    "Manager Comments", "Manager Decision Date", "Manager Decision By",
    "Director Comments", "Director Decision Date", "Director Decision By",
    "Submitted By", "Submitted Date", "Payroll Status", "Payroll Date",
    "Payroll By", "PDF File Path"
]
INSPECTOR_BONUS_COLUMNS = [
    "ID", "Inspector Name", "Month & Year", "Days Absent", "Reasons for Absence",
    "Total Jobs Completed", "Bonus Amount (£)",
    "Status",
    "Director Comments", "Director Decision Date", "Director Decision By",
    "Submitted By", "Submitted Date", "PDF File Path"
]
DEFAULT_USERS = [
    {"full_name": "National Grid Manager", "username": "national_grid", "password": "acoole123", "role": "Manager", "dept": "National Grid", "can_access_inspector_bonus": True},
    {"full_name": "Isolator Manager", "username": "isolator", "password": "acoole123", "role": "Manager", "dept": "Isolator"},
    {"full_name": "Project Manager", "username": "project", "password": "acoole123", "role": "Manager", "dept": "Project"},
    {"full_name": "Accounts Manager", "username": "accounts", "password": "acoole123", "role": "Manager", "dept": "Accounts"},
    {"full_name": "Andy Acoole", "username": "andy", "password": "andy2026", "role": "Director", "dept": "ACoole Electrical Ltd"},
    {"full_name": "System Administrator", "username": "wais", "password": "superadmin123", "role": "Super Admin", "dept": "System Administration"},
    {"full_name": "Payroll Team", "username": "payroll", "password": "payroll2026", "role": "Payroll", "dept": "Payroll Department"}
]
PERMISSION_DEFAULTS = {
    "Work Order Employee": {"can_view_all_dept": False, "can_generate_pdf": False, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False},
    "Work Order Manager": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False},
    "Staff": {"can_view_all_dept": False, "can_generate_pdf": False, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False},
    "Team Member": {"can_view_all_dept": True, "can_generate_pdf": False, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False},
    "Manager": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": True},
    "Director": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True, "can_approve_requests": True, "can_access_inspector_bonus": True},
    "Payroll": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True, "can_approve_requests": False, "can_access_inspector_bonus": True},
    "Super Admin": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True, "can_approve_requests": True, "can_access_inspector_bonus": True}
}
PERMISSION_LABELS = {
    "can_view_all_dept": "👁️ View All Department Requests",
    "can_generate_pdf": "📄 Generate & Download PDFs",
    "can_download_data": "📥 Download Data Backups",
    "can_approve_requests": "✅ Approve/Reject Requests",
    "can_access_inspector_bonus": "💰 National Grid Inspector Bonus"
}

# ============================================================
# PDF LIBRARY
# ============================================================
try:
    from fpdf import FPDF
    PDF_AVAILABLE = True
except ImportError:
    FPDF = None
    PDF_AVAILABLE = False

# ============================================================
# ✅ SAFE EXCEL INIT
# ============================================================
def safe_init_excel(path, columns):
    if not os.path.exists(path):
        pd.DataFrame(columns=columns).to_excel(path, index=False, engine="openpyxl")
        return True
    try:
        df = pd.read_excel(path, engine="openpyxl").fillna("")
        changed = False
        for col in columns:
            if col not in df.columns:
                df[col] = ""
                changed = True
        if changed:
            df.to_excel(path, index=False, engine="openpyxl")
        return True
    except Exception:
        os.remove(path)
        pd.DataFrame(columns=columns).to_excel(path, index=False, engine="openpyxl")
        return True

# ============================================================
# AUDIT LOG FUNCTIONS
# ============================================================
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

def archive_audit_log():
    os.makedirs(ARCHIVE_FOLDER, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    archive_path = os.path.join(ARCHIVE_FOLDER, f"audit_log_archive_{ts}.xlsx")
    if os.path.exists(AUDIT_LOG_FILE):
        df = pd.read_excel(AUDIT_LOG_FILE, engine="openpyxl")
        df.to_excel(archive_path, index=False, engine="openpyxl")
        return archive_path, len(df)
    return None, 0

def clear_audit_log_file():
    if os.path.exists(AUDIT_LOG_FILE):
        os.remove(AUDIT_LOG_FILE)
    pd.DataFrame(columns=AUDIT_COLUMNS).to_excel(AUDIT_LOG_FILE, index=False, engine="openpyxl")
    _invalidate_data_cache("_audit_log_cache")
    st.session_state["_audit_log_count"] = 0
    sync_saved_file_to_drive(AUDIT_LOG_FILE)

def clear_all_requests_file():
    if os.path.exists(EXCEL_PATH):
        os.remove(EXCEL_PATH)
    pd.DataFrame(columns=EXCEL_COLUMNS).to_excel(EXCEL_PATH, index=False, engine="openpyxl")
    _set_data_cache("_records_cache", [])
    sync_saved_file_to_drive(EXCEL_PATH)
    st.session_state["_live_data_reset"] = datetime.now().isoformat()

def clear_all_inspector_bonus():
    if os.path.exists(INSPECTOR_BONUS_PATH):
        os.remove(INSPECTOR_BONUS_PATH)
    pd.DataFrame(columns=INSPECTOR_BONUS_COLUMNS).to_excel(INSPECTOR_BONUS_PATH, index=False, engine="openpyxl")
    _set_data_cache("_inspector_bonus_cache", [])
    sync_saved_file_to_drive(INSPECTOR_BONUS_PATH)

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

def log_action(action, req_id="-", old_data=None, new_data=None, fields_changed=None, decision_by=None, decision_date=None):
    if not st.session_state.get("logged_in"): return
    user_info = st.session_state.user_info
    username = user_info.get("full_name", user_info.get("username", "Unknown"))
    role = user_info.get("role", "Unknown")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    SETTING_ACTIONS = ["CATEGORY_ADDED", "CATEGORY_EDITED", "CATEGORY_DELETED",
        "DEPARTMENT_ADDED", "DEPARTMENT_EDITED", "DEPARTMENT_DELETED",
        "ROLE_ADDED", "ROLE_EDITED", "ROLE_DELETED",
        "USER_CREATED", "USER_EDITED", "USER_DELETED", "PASSWORD_CHANGED", "PASSWORD_RESET",
        "INSPECTOR_BONUS_CREATED", "INSPECTOR_BONUS_APPROVED", "INSPECTOR_BONUS_REJECTED",
        "INSPECTOR_BONUS_EDITED", "INSPECTOR_BONUS_STATUS_CHANGED",
        "SUPER_ADMIN_CLEAR_INSPECTOR_BONUS"]
    if action in SETTING_ACTIONS:
        action_labels = {
            "CATEGORY_ADDED": "🏷️ Category Added", "CATEGORY_EDITED": "🏷️ Category Edited", "CATEGORY_DELETED": "🏷️ Category Deleted",
            "DEPARTMENT_ADDED": "🏢 Department Added", "DEPARTMENT_EDITED": "🏢 Department Edited", "DEPARTMENT_DELETED": "🏢 Department Deleted",
            "ROLE_ADDED": "🎖️ Role Added", "ROLE_EDITED": "🎖️ Role/Permissions Edited", "ROLE_DELETED": "🎖️ Role Deleted",
            "USER_CREATED": "👤 User Account Created", "USER_EDITED": "👤 User Account Edited",
            "USER_DELETED": "👤 User Account Deleted", "PASSWORD_CHANGED": "🔑 Password Changed", "PASSWORD_RESET": "🔑 Password Reset",
            "INSPECTOR_BONUS_CREATED": "💰 Inspector Bonus Submitted",
            "INSPECTOR_BONUS_APPROVED": "💰 Inspector Bonus Approved",
            "INSPECTOR_BONUS_REJECTED": "💰 Inspector Bonus Rejected",
            "INSPECTOR_BONUS_EDITED": "💰 Inspector Bonus Edited",
            "INSPECTOR_BONUS_STATUS_CHANGED": "💰 Inspector Bonus Status Changed",
            "SUPER_ADMIN_CLEAR_INSPECTOR_BONUS": "🧹 Super Admin — All Inspector Bonuses Cleared",
        }
        display_action = action_labels.get(action, action)
        old_val = json.dumps(old_data, ensure_ascii=False)[:300] if old_data else "-"
        new_val = json.dumps(new_data, ensure_ascii=False)[:300] if new_data else "-"
        save_audit_entry({"AuditID": _get_next_audit_id(), "Timestamp": timestamp,
            "User_Name": username, "User_Role": role, "Action": display_action,
            "Request_ID": str(req_id), "Department": "-", "Amount": "-",
            "Decision_By": decision_by or "-", "Decision_Date": decision_date or "-",
            "Field_Changed": "Inspector Bonus",
            "Old_Value": old_val, "New_Value": new_val, "IP_Address": "Auto-Logged"})
        return
    dept, amount, saved_decision_by, saved_decision_date = get_request_details(req_id)
    final_decision_by = decision_by or saved_decision_by
    final_decision_date = decision_date or saved_decision_date
    if action.startswith("WORK_ORDER_"):
        wo_labels = {
            "WORK_ORDER_CREATED": "🛠️ Work Order Created",
            "WORK_ORDER_MANAGER_CREATED": "🛠️ Work Order Submitted to Director",
            "WORK_ORDER_MANAGER_APPROVED": "🛠️ Work Order Approved by Manager → Director",
            "WORK_ORDER_MANAGER_REJECTED": "🛠️ Work Order Rejected by Manager → Returned to Employee",
            "WORK_ORDER_RETURNED": "🛠️ Work Order Returned to Employee",
            "WORK_ORDER_DIRECTOR_APPROVED": "🛠️ Work Order Approved for Payment",
            "WORK_ORDER_DIRECTOR_REJECTED": "🛠️ Work Order Rejected by Director",
            "WORK_ORDER_STATUS_CHANGED": "🛠️ Work Order Status Changed",
            "WORK_ORDER_EDITED": "🛠️ Work Order Edited",
            "WORK_ORDER_PAID": "🛠️ Work Order Paid",
        }
        old_v = json.dumps(old_data, ensure_ascii=False)[:300] if old_data else "-"
        new_v = json.dumps(new_data, ensure_ascii=False)[:300] if new_data else "-"
        save_audit_entry({
            "AuditID": _get_next_audit_id(), "Timestamp": timestamp,
            "User_Name": username, "User_Role": role, "Action": wo_labels.get(action, action),
            "Request_ID": str(req_id), "Department": dept, "Amount": amount,
            "Decision_By": final_decision_by or "-", "Decision_Date": final_decision_date or timestamp,
            "Field_Changed": "Work Order Status",
            "Old_Value": old_v if old_data else "-",
            "New_Value": new_v if new_data else action.replace("WORK_ORDER_", "").replace("_", " ").title(),
            "IP_Address": "Auto-Logged"
        })
        return
    if action in ["CREATED", "DELETED"]:
        save_audit_entry({"AuditID": _get_next_audit_id(), "Timestamp": timestamp,
            "User_Name": username, "User_Role": role, "Action": action, "Request_ID": str(req_id),
            "Department": dept, "Amount": amount, "Decision_By": "-", "Decision_Date": "-",
            "Field_Changed": "-", "Old_Value": "-",
            "New_Value": "New Request Created" if action == "CREATED" else "Request Permanently Deleted",
            "IP_Address": "Auto-Logged"})
    elif action in ["APPROVED", "REJECTED", "STATUS_CHANGED"]:
        status_text = "Approved" if action == "APPROVED" else "Rejected" if action == "REJECTED" else "Status Changed"
        save_audit_entry({"AuditID": _get_next_audit_id(), "Timestamp": timestamp,
            "User_Name": username, "User_Role": role, "Action": action, "Request_ID": str(req_id),
            "Department": dept, "Amount": amount, "Decision_By": final_decision_by,
            "Decision_Date": final_decision_date, "Field_Changed": "Status",
            "Old_Value": "Pending", "New_Value": status_text, "IP_Address": "Auto-Logged"})
    elif action == "EDITED" and old_data and new_data:
        field_labels = {"emp_name": "Employee Name", "dept": "Department", "type": "Transaction Type",
            "category": "Category", "date": "Date", "amount": "Amount (£)",
            "manager": "Line Manager", "desc": "Description", "status": "Status", "attachment_name": "Attachments"}
        for key, label in field_labels.items():
            old = str(old_data.get(key, "")).strip()
            new = str(new_data.get(key, "")).strip()
            if old != new:
                save_audit_entry({"AuditID": _get_next_audit_id(), "Timestamp": timestamp,
                    "User_Name": username, "User_Role": role, "Action": "EDITED", "Request_ID": str(req_id),
                    "Department": dept, "Amount": amount, "Decision_By": "-", "Decision_Date": "-",
                    "Field_Changed": label, "Old_Value": old, "New_Value": new, "IP_Address": "Auto-Logged"})

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
        icon = {"CREATED": "➕", "EDITED": "✏️", "APPROVED": "✅", "REJECTED": "❌", "DELETED": "🗑️", "STATUS_CHANGED": "🔄",
            "🏷️ Category Added": "🏷️", "🏷️ Category Edited": "🏷️", "🏷️ Category Deleted": "🏷️",
            "🏢 Department Added": "🏢", "🏢 Department Edited": "🏢", "🏢 Department Deleted": "🏢",
            "🎖️ Role Added": "🎖️", "🎖️ Role/Permissions Edited": "🎖️", "🎖️ Role Deleted": "🎖️",
            "👤 User Account Created": "👤", "👤 User Account Edited": "✏️", "👤 User Account Deleted": "🗑️",
            "🔑 Password Changed": "🔑", "🔑 Password Reset": "🔑"}.get(action, "ℹ️")
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

# ============================================================
# HELPER FUNCTIONS
# ============================================================
def format_date(d):
    if not d or str(d).strip().lower() in ["", "none", "nan"]:
        return "-"
    return str(d).strip()

def get_next_id(all_records):
    if not all_records: return 1
    return max(int(r.get("id", 0)) for r in all_records) + 1

def update_record_status_in_excel(req_id, new_status, comments, approved_by):
    records = load_records_from_excel()
    decision_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for r in records:
        if int(r.get("id", 0)) == int(req_id):
            r["status"] = new_status.lower()
            r["decision_date"] = decision_datetime
            r["decision_by"] = approved_by
            if comments.strip():
                r["director_comments"] = f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] {comments.strip()}"
            r["pdf_path"] = ""
            break
    save_all_records(records)

def delete_record_by_id(req_id):
    records = load_records_from_excel()
    records = [r for r in records if int(r.get("id", 0)) != int(req_id)]
    save_all_records(records)
    log_action("DELETED", req_id)

def show_old_new_comparison(old_json, new_rec):
    try: old = json.loads(old_json) if old_json and old_json != "{}" else {}
    except: old = {}
    if not old: st.info("📋 New request — no previous version."); return
    st.markdown("#### 🔄 Changes (Previous → New)")
    fields = [("emp_name", "Employee Name"), ("dept", "Department"), ("type", "Transaction Type"),
        ("category", "Category"), ("date", "Date"), ("amount", "Amount (£)"),
        ("manager", "Line Manager"), ("desc", "Description")]
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
                for path in (EXCEL_PATH, USER_DB_PATH, SETTINGS_PATH, AUDIT_LOG_PATH, INSPECTOR_BONUS_PATH, WORK_ORDERS_PATH):
                    remote = _drive_find_file(os.path.basename(path))
                    if remote:
                        _drive_download_file(remote["id"], path)
            _invalidate_data_cache("_records_cache", "_users_cache", "_settings_cache", "_audit_log_cache", "_work_orders_cache", "_inspector_bonus_cache", "_audit_log_count")
            st.session_state["_last_refresh"] = datetime.now().isoformat()
        st.rerun()

def make_request_title(req):
    status, amount, dt, dec_by, decision_dt = req["status"].upper(), f"£{req['amount']:.2f}", format_date(req.get("date", "")), req.get("decision_by", ""), format_date(req.get("decision_date", ""))
    if req["status"] == "pending": return f"🟡 ID #{req['id']} | {req['emp_name']} | PENDING | {amount} | 📅 {dt}"
    elif req["status"] == "approved": return f"🟢 ID #{req['id']} | {req['emp_name']} | APPROVED | {amount} | ✅ Approved by {dec_by} on {decision_dt}"
    elif req["status"] == "rejected": return f"🔴 ID #{req['id']} | {req['emp_name']} | REJECTED | {amount} | ❌ Rejected by {dec_by} on {decision_dt}"
    else: return f"⚪ ID #{req['id']} | {req['emp_name']} | {status} | {amount} | 📅 {dt}"

# ============================================================
# SETTINGS FUNCTIONS
# ============================================================
def init_settings():
    if not os.path.exists(SETTINGS_PATH):
        pd.DataFrame([{"setting": "categories", "value": "|".join(DEFAULT_CATEGORIES)},
            {"setting": "roles", "value": "|".join(DEFAULT_ROLES)},
            {"setting": "departments", "value": "|".join(DEFAULT_DEPARTMENTS)}]).to_excel(SETTINGS_PATH, index=False, engine="openpyxl")

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
    return _load_setting_value("departments", DEFAULT_DEPARTMENTS)

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

def load_categories():
    return _load_setting_value("categories", DEFAULT_CATEGORIES)

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

# ============================================================
# INITIALISE PERMANENT GOOGLE DRIVE STORAGE
# ============================================================
initialise_drive_storage()

# ============================================================
# USER DATABASE
# ============================================================
def init_user_db():
    safe_init_excel(USER_DB_PATH, ["full_name","username","password","role","dept",
        "can_view_all_dept","can_generate_pdf","can_download_data","can_approve_requests","can_access_inspector_bonus"])
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
        rows.append({"full_name": u.get("full_name", username), "username": username,
            "password": u.get("password", ""), "role": u.get("role", "Staff"),
            "dept": u.get("dept", ""), "can_view_all_dept": u.get("can_view_all_dept", False),
            "can_generate_pdf": u.get("can_generate_pdf", False),
            "can_download_data": u.get("can_download_data", False),
            "can_approve_requests": u.get("can_approve_requests", False),
            "can_access_inspector_bonus": u.get("can_access_inspector_bonus", False)})
    pd.DataFrame(rows).to_excel(USER_DB_PATH, index=False, engine="openpyxl")
    _invalidate_data_cache("_users_cache")
    sync_saved_file_to_drive(USER_DB_PATH)

def load_users(force=False):
    init_user_db()
    if not force and "_users_cache" in st.session_state:
        return dict(st.session_state["_users_cache"])
    try:
        df = _read_excel_records(USER_DB_PATH)
        users = {}
        for _, r in df.iterrows():
            username = str(r.get("username", "")).strip()
            if not username:
                continue
            users[username] = {
                "full_name": str(r.get("full_name", username)).strip(),
                "password": str(r.get("password", "")),
                "role": str(r.get("role", "Staff")),
                "dept": str(r.get("dept", "")),
                "can_view_all_dept": str(r.get("can_view_all_dept", "False")).lower() == "true",
                "can_generate_pdf": str(r.get("can_generate_pdf", "False")).lower() == "true",
                "can_download_data": str(r.get("can_download_data", "False")).lower() == "true",
                "can_approve_requests": str(r.get("can_approve_requests", "False")).lower() == "true",
                "can_access_inspector_bonus": str(r.get("can_access_inspector_bonus", "False")).lower() == "true",
            }
        _set_data_cache("_users_cache", users)
        return dict(users)
    except Exception as e:
        st.error(f"User DB Load Error: {e}")
        return {}

# ============================================================
# WORK ORDERS — EMPLOYEE → MANAGER → DIRECTOR → PAYROLL
# ============================================================
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
                "manual_work_order_no": str(r.get("Manual Work Order No.", "")).strip()
                    or str(r.get("Work Order ID", "")).strip(),
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
    except Exception:
        return False

def save_all_work_orders(records, sync=True):
    rows = []
    for r in records:
        rows.append({
            "Work Order ID": str(r.get("id", "")),
            "Manual Work Order No.": str(r.get("manual_work_order_no", "")).strip() or str(r.get("id", "")),
            "Employee Name": str(r.get("emp_name", "")),
            "Department": str(r.get("dept", "")), "Work Date": str(r.get("work_date", "")),
            "Hours": float(r.get("hours", 0)), "Amount (£)": float(r.get("amount", 0)),
            "Manager": str(r.get("manager", "")), "Description": str(r.get("desc", "")),
            "Attachment Name": str(r.get("attachment_name", "None")), "Status": str(r.get("status", "pending_manager")),
            "Site Address": str(r.get("site_address", "")),
            "Customer Job No.": str(r.get("customer_job_no", "")),
            "Manager Comments": str(r.get("manager_comments", "")),
            "Manager Decision Date": str(r.get("manager_decision_date", "")),
            "Manager Decision By": str(r.get("manager_decision_by", "")),
            "Director Comments": str(r.get("director_comments", "")),
            "Director Decision Date": str(r.get("director_decision_date", "")),
            "Director Decision By": str(r.get("director_decision_by", "")),
            "Submitted By": str(r.get("submitted_by", "")), "Submitted Date": str(r.get("submitted_date", "")),
            "Payroll Status": str(r.get("payroll_status", "Pending")), "Payroll Date": str(r.get("payroll_date", "")),
            "Payroll By": str(r.get("payroll_by", "")), "PDF File Path": str(r.get("pdf_path", "")),
        })
    pd.DataFrame(rows, columns=WORK_ORDER_COLUMNS).to_excel(WORK_ORDERS_PATH, index=False, engine="openpyxl")
    _set_data_cache("_work_orders_cache", list(records))
    if sync:
        sync_saved_file_to_drive(WORK_ORDERS_PATH)

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
    candidates = [
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
         "/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf"),
    ]
    for regular, bold in candidates:
        if os.path.exists(regular) and os.path.exists(bold):
            return regular, bold
    return None, None

def _pdf_text(value):
    if value is None:
        return ""
    return str(value).replace("\x00", "")

def work_order_pdf(req, upload_to_drive=True):
    if not PDF_AVAILABLE:
        return None
    try:
        pdf = FPDF()
        pdf.add_page()
        regular_font, bold_font = _pdf_font_paths()
        if regular_font and bold_font:
            pdf.add_font("DejaVu", "", regular_font)
            pdf.add_font("DejaVu", "B", bold_font)
            font_family = "DejaVu"
        else:
            font_family = "Helvetica"

        def safe(value):
            text = _pdf_text(value)
            if font_family == "Helvetica":
                return text.encode("latin-1", "replace").decode("latin-1")
            return text

        if os.path.exists(LOGO_PATH):
            try:
                pdf.image(LOGO_PATH, x=75, y=10, w=60)
                pdf.ln(25)
            except Exception:
                pdf.ln(3)
        else:
            pdf.ln(3)

        label_width = 45
        content_width = pdf.w - pdf.l_margin - pdf.r_margin
        value_width = content_width - label_width

        pdf.set_font(font_family, "B", 16)
        pdf.cell(0, 10, safe("WORK ORDER - PAYMENT AUTHORISATION"), ln=True, align="C")
        pdf.ln(5)

        pdf.set_font(font_family, "", 10)
        rows = [
            ("Work Order No.", get_work_order_number(req)),
            ("Employee",       req.get("emp_name", "")),
            ("Department",     req.get("dept", "")),
            ("Site Address",   req.get("site_address", "")),
            ("Customer Job No.", req.get("customer_job_no", "")),
            ("Work Date",      req.get("work_date", "")),
            ("Amount",         f"GBP {float(req.get('amount', 0)):.2f}"),
            ("Manager",        req.get("manager", "")),
            ("Submitted By",   req.get("submitted_by", "")),
            ("Submitted Date", req.get("submitted_date", "")),
        ]

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
            try:
                pdf.image(APPROVED_STAMP_PATH, x=pdf.get_x() + 5, y=pdf.get_y(), w=35)
            except Exception:
                pass
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
        if upload_to_drive:
            _upload_to_drive_bg(path, os.path.basename(path))
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
            st.download_button(
                "📄 Download Work Order PDF",
                f.read(),
                file_name=os.path.basename(path),
                key=f"wo_pdf_{req.get('id')}"
            )

def work_order_total_pdf(records, employee_filter, from_date, to_date, prepared_by=""):
    if not PDF_AVAILABLE:
        return None
    try:
        selected = list(records)
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
            if family == "Helvetica":
                return text.encode("latin-1", "replace").decode("latin-1")
            return text

        if os.path.exists(LOGO_PATH):
            try:
                pdf.image(LOGO_PATH, x=75, y=10, w=60)
                pdf.ln(25)
            except Exception:
                pdf.ln(3)
        else:
            pdf.ln(3)

        pdf.set_font(family, "B", 15)
        pdf.cell(0, 9, safe("APPROVED WORK ORDER TOTAL"), ln=True, align="C")
        pdf.ln(3)

        pdf.set_font(family, "", 10)
        scope = employee_filter if employee_filter and employee_filter != "All Employees" else "All Employees"
        pdf.cell(0, 6, safe(f"Employee: {scope}"), ln=True)
        pdf.cell(0, 6, safe(f"Work date range: {from_date} to {to_date}"), ln=True)
        if prepared_by:
            pdf.cell(0, 6, safe(f"Prepared by: {prepared_by}"), ln=True)
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
        filename = (
            f"Approved_Work_Order_Total_{safe_emp}_"
            f"GBP{total:.2f}_{from_date}_to_{to_date}.pdf"
        )
        path = os.path.join(WORK_ORDER_PDF_DIR, filename)
        pdf.output(path)
        _upload_to_drive_bg(path, os.path.basename(path))
        return path
    except Exception as e:
        st.error(f"Work Order Total PDF Error: {e}")
        return None

def render_work_order_total(records, scope_department=None, key_prefix="wo_total", prepared_by=""):
    st.markdown("### 💷 Approved Work Order Total")
    st.caption("Select an employee and work-date range. Only Director-approved work orders are included. Work Order No. is the manually entered number.")
    scoped = [r for r in records if r.get("status") in ("approved_payment", "approved")]
    if scope_department:
        scoped = [r for r in scoped if str(r.get("dept", "")) == str(scope_department)]
    employees = sorted({str(r.get("emp_name", "")).strip() for r in scoped if str(r.get("emp_name", "")).strip()})
    c1, c2, c3 = st.columns(3)
    with c1:
        employee = st.selectbox("👤 Employee", ["All Employees"] + employees, key=f"{key_prefix}_employee")
    with c2:
        from_date = st.date_input("📅 From Date", value=date.today().replace(day=1), key=f"{key_prefix}_from")
    with c3:
        to_date = st.date_input("📅 To Date", value=date.today(), key=f"{key_prefix}_to")
    if from_date > to_date:
        st.error("From Date cannot be after To Date.")
        return
    selected = []
    for r in scoped:
        try:
            d = pd.to_datetime(str(r.get("work_date", "")), errors="coerce").date()
        except Exception:
            d = None
        if d is None:
            continue
        if not (from_date <= d <= to_date):
            continue
        if employee != "All Employees" and str(r.get("emp_name", "")).strip() != employee:
            continue
        selected.append(r)
    total = sum(float(r.get("amount", 0) or 0) for r in selected)
    st.metric("💷 Total Approved", f"£{total:,.2f}")
    st.metric("📋 Work Orders Included", len(selected))
    if selected:
        rows = [{
            "Work Order No.": get_work_order_number(r),
            "Employee": r.get("emp_name", ""),
            "Work Date": r.get("work_date", ""),
            "Amount (£)": float(r.get("amount", 0) or 0),
            "Approved By": r.get("director_decision_by", ""),
        } for r in selected]
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
    st.caption("Select an employee and date range, then generate a ZIP containing the matching approved Work Order PDFs.")
    scoped = [r for r in records if r.get("status") in ("approved_payment", "approved")]
    if not scoped:
        st.info("No approved work orders available.")
        return
    employees = sorted({str(r.get("emp_name", "")).strip() for r in scoped if str(r.get("emp_name", "")).strip()})
    c1, c2, c3 = st.columns(3)
    with c1:
        employee = st.selectbox("👤 Employee", ["All Employees"] + employees, key=f"{key_prefix}_employee")
    with c2:
        from_date = st.date_input("📅 From Date", value=date.today().replace(day=1), key=f"{key_prefix}_from")
    with c3:
        to_date = st.date_input("📅 To Date", value=date.today(), key=f"{key_prefix}_to")
    if from_date > to_date:
        st.error("From Date cannot be after To Date.")
        return
    selected = []
    for r in scoped:
        try:
            d = pd.to_datetime(str(r.get("work_date", "")), errors="coerce").date()
        except Exception:
            d = None
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
                    if not path or not os.path.exists(path):
                        path = work_order_pdf(r, upload_to_drive=False)
                    if path and os.path.exists(path):
                        with open(path, "rb") as f:
                            zipf.writestr(os.path.basename(path), f.read())
                        pdf_count += 1
                    else:
                        failed.append(get_work_order_number(r))
        zip_buffer.seek(0)
        if pdf_count > 0:
            safe_emp = "all_employees" if employee == "All Employees" else "_".join(employee.split())
            filename = f"Work_Order_PDFs_{safe_emp}_{from_date}_to_{to_date}.zip"
            st.session_state[f"{key_prefix}_data"] = zip_buffer.getvalue()
            st.session_state[f"{key_prefix}_name"] = filename
            st.success(f"✅ Generated {pdf_count} PDF(s). Click below to download.")
            if failed:
                st.warning(f"⚠️ Could not include: {', '.join(failed)}")
        else:
            st.error("❌ No PDFs could be generated.")
    if st.session_state.get(f"{key_prefix}_data"):
        st.download_button(
            "⬇️ Download ZIP",
            data=st.session_state[f"{key_prefix}_data"],
            file_name=st.session_state.get(f"{key_prefix}_name", "Work_Order_PDFs.zip"),
            mime="application/zip",
            type="primary",
            key=f"{key_prefix}_dl",
        )

# ✅ v4.8 — Employee portal:
#   • Two tabs: "My Work Orders" AND "Department Work Orders" (own dept, incl. manager-submitted)
#   • Duplicate Work Order No. shows "Work Order No. already Exist." and KEEPS all form data
#   • Form is only cleared after a SUCCESSFUL submission
def render_work_order_employee_portal(current_user, current_dept):
    st.subheader("🛠️ Work Orders")
    orders = load_work_orders()
    managers = _manager_options_for_department(current_dept)

    _success_msg = st.session_state.pop("emp_new_wo_success", None)
    if _success_msg:
        st.success(_success_msg)

    editing_id = st.session_state.get("editing_work_order_id_emp")
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
                    try:
                        _d = datetime.strptime(str(rec.get("work_date", ""))[:10], "%Y-%m-%d").date()
                    except Exception:
                        _d = date.today()
                    e_date = st.date_input("📅 Date", value=_d)
                    e_amount = st.number_input(
                        "💷 Amount (£)", min_value=0.01, step=1.0, format="%.2f",
                        value=max(float(rec.get("amount", 0.01) or 0.01), 0.01)
                    )
                    e_desc = st.text_area("📝 Description", value=str(rec.get("desc", "")), height=150)

                if managers:
                    current_mgr = str(rec.get("manager", "") or "")
                    mgr_idx = managers.index(current_mgr) if current_mgr in managers else 0
                    e_manager = st.selectbox("👔 Send to Manager", managers, index=mgr_idx)
                else:
                    e_manager = st.text_input("👔 Manager Name", value=str(rec.get("manager", "")))

                e_files = st.file_uploader(
                    "📎 Add New Attachments (optional)",
                    type=["pdf", "png", "jpg", "jpeg"],
                    accept_multiple_files=True,
                    key=f"emp_edit_files_{editing_id}"
                )

                existing_att = str(rec.get("attachment_name", "None") or "None")
                if existing_att and existing_att.lower() not in ("none", "nan", ""):
                    st.caption(f"📎 Existing attachments: {existing_att}")

                csave, ccancel = st.columns(2)
                with csave:
                    save_edit = st.form_submit_button("💾 Save Changes & Resubmit", type="primary", use_container_width=True)
                with ccancel:
                    cancel_edit = st.form_submit_button("❌ Cancel", use_container_width=True)

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
                                with open(fp, "wb") as out_file:
                                    out_file.write(f.getbuffer())
                                _upload_to_drive_bg(fp, fn)
                                new_attachments.append(fn)

                        old_data = {
                            "manual_work_order_no": rec.get("manual_work_order_no"),
                            "emp_name": rec.get("emp_name"),
                            "site_address": rec.get("site_address"),
                            "customer_job_no": rec.get("customer_job_no"),
                            "work_date": rec.get("work_date"),
                            "amount": rec.get("amount"),
                            "desc": rec.get("desc"),
                            "manager": rec.get("manager"),
                            "status": rec.get("status"),
                        }
                        for x in orders:
                            if str(x.get("id")) == str(editing_id):
                                x["manual_work_order_no"] = e_wo_no.strip()
                                x["emp_name"] = e_emp.strip()
                                x["site_address"] = e_site.strip()
                                x["customer_job_no"] = e_cjno.strip()
                                x["work_date"] = str(e_date)
                                x["amount"] = float(e_amount)
                                x["desc"] = e_desc.strip()
                                x["manager"] = e_manager.strip()
                                x["attachment_name"] = ", ".join(new_attachments) or "None"
                                x["status"] = "pending_manager"
                                x["manager_comments"] = ""
                                x["manager_decision_date"] = ""
                                x["manager_decision_by"] = ""
                                x["director_comments"] = ""
                                x["director_decision_date"] = ""
                                x["director_decision_by"] = ""
                                x["pdf_path"] = ""
                                x["submitted_by"] = current_user
                                x["submitted_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                break
                        save_all_work_orders(orders)
                        new_data = {
                            "manual_work_order_no": e_wo_no.strip(),
                            "emp_name": e_emp.strip(),
                            "site_address": e_site.strip(),
                            "customer_job_no": e_cjno.strip(),
                            "work_date": str(e_date),
                            "amount": float(e_amount),
                            "desc": e_desc.strip(),
                            "manager": e_manager.strip(),
                            "status": "pending_manager",
                        }
                        log_action("WORK_ORDER_EDITED", editing_id, old_data=old_data, new_data=new_data)
                        st.session_state.editing_work_order_id_emp = None
                        st.success("✅ Work Order updated and re-sent to Manager for review.")
                        st.rerun()

                if cancel_edit:
                    st.session_state.editing_work_order_id_emp = None
                    st.rerun()
            st.divider()
        else:
            st.session_state.editing_work_order_id_emp = None

    st.markdown("### 📤 Submit New Work Order")
    st.caption("Complete the work-order details below. No hours/time entry is required.")
    wid = get_next_work_order_id(orders)

    K_WO    = "emp_new_wo_no"
    K_EMP   = "emp_new_contractor"
    K_SITE  = "emp_new_site"
    K_CJNO  = "emp_new_cjno"
    K_DATE  = "emp_new_date"
    K_AMT   = "emp_new_amount"
    K_DESC  = "emp_new_desc"
    K_MGR   = "emp_new_manager"
    K_FILES = "emp_new_files"

    # Only reset the form after a SUCCESSFUL submission (not on duplicate / validation error)
    if st.session_state.get("emp_new_form_reset"):
        for k in (K_WO, K_EMP, K_SITE, K_CJNO, K_DATE, K_AMT, K_DESC, K_MGR, K_FILES):
            st.session_state.pop(k, None)
        st.session_state["emp_new_form_reset"] = False

    with st.form("employee_new_work_order_form", clear_on_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            work_order_no = st.text_input(
                "🧾 Work Order No.",
                key=K_WO,
                placeholder="Enter the Work Order No. from the work-order sheet"
            )
            contractor_employee = st.text_input(
                "👤 Contractor / Employee Labour",
                key=K_EMP,
                value=current_user,
                placeholder="Enter contractor or employee name"
            )
            site_address = st.text_area(
                "📍 Site Address",
                key=K_SITE,
                placeholder="Enter the full site address",
                height=90
            )
            customer_job_no = st.text_input(
                "📘 Customer Job No.",
                key=K_CJNO,
                placeholder="Enter Customer Job No."
            )
        with c2:
            work_date = st.date_input("📅 Date", key=K_DATE, value=date.today())
            amount = st.number_input(
                "💷 Amount (£)", min_value=0.01, step=1.0, format="%.2f",
                key=K_AMT, value=0.01
            )
            description = st.text_area(
                "📝 Description",
                key=K_DESC,
                placeholder="Describe the work completed...",
                height=150
            )

        if managers:
            manager = st.selectbox("👔 Send to Manager", managers, key=K_MGR)
        else:
            manager = st.text_input("👔 Manager Name", key=K_MGR, placeholder="Enter manager name")

        files = st.file_uploader(
            "📎 Supporting Work Order Document (optional)",
            type=["pdf", "png", "jpg", "jpeg"],
            accept_multiple_files=True,
            key=K_FILES
        )
        submitted = st.form_submit_button(
            "📤 Submit Work Order to Manager",
            type="primary",
            use_container_width=True
        )

        if submitted:
            errors = []
            if not work_order_no.strip(): errors.append("Work Order No.")
            if not contractor_employee.strip(): errors.append("Contractor / Employee Labour")
            if not site_address.strip(): errors.append("Site Address")
            if not customer_job_no.strip(): errors.append("Customer Job No.")
            if not description.strip(): errors.append("Description")
            if not manager.strip(): errors.append("Manager")

            duplicate = any(
                str(r.get("manual_work_order_no", "")).strip().lower() == work_order_no.strip().lower()
                for r in orders if str(r.get("manual_work_order_no", "")).strip()
            )

            if duplicate:
                st.error(
                    "❌ **Work Order No. already Exist.** "
                    "Please change the Work Order No. and submit again — your form data has been kept."
                )
            elif errors:
                st.error("Please correct: " + ", ".join(errors) + ".")
            else:
                attachments = []
                for i, f in enumerate(files or [], 1):
                    safe_name = os.path.basename(f.name).replace("/", "_").replace("\\", "_")
                    fn = f"{wid}_F{i}_{safe_name}"
                    fp = os.path.join(UPLOAD_DIR, fn)
                    with open(fp, "wb") as out_file:
                        out_file.write(f.getbuffer())
                    _upload_to_drive_bg(fp, fn)
                    attachments.append(fn)

                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                rec = {
                    "id": wid,
                    "manual_work_order_no": work_order_no.strip(),
                    "emp_name": contractor_employee.strip(),
                    "dept": current_dept,
                    "work_date": str(work_date),
                    "hours": 0.0,
                    "customer_job_no": customer_job_no.strip(),
                    "site_address": site_address.strip(),
                    "amount": float(amount),
                    "manager": manager.strip(),
                    "desc": description.strip(),
                    "attachment_name": ", ".join(attachments) or "None",
                    "status": "pending_manager",
                    "manager_comments": "",
                    "manager_decision_date": "",
                    "manager_decision_by": "",
                    "director_comments": "",
                    "director_decision_date": "",
                    "director_decision_by": "",
                    "submitted_by": current_user,
                    "submitted_date": now,
                    "payroll_status": "Pending",
                    "payroll_date": "",
                    "payroll_by": "",
                    "pdf_path": "",
                }
                orders.append(rec)
                save_all_work_orders(orders)
                log_action("WORK_ORDER_CREATED", wid, decision_by=current_user)
                st.session_state["emp_new_wo_success"] = (
                    f"✅ Work Order No. {work_order_no.strip()} submitted to {manager} for review."
                )
                st.session_state["emp_new_form_reset"] = True
                st.rerun()

    st.divider()
    tab_mine, tab_dept = st.tabs([
        "📋 My Work Orders",
        f"🏢 Department Work Orders ({current_dept})"
    ])

    with tab_mine:
        q = st.text_input(
            "🔎 Search my work orders",
            placeholder="Search by ID, employee, manager, status, amount, date or description...",
            key="wo_employee_search"
        )
        mine = [r for r in orders if r.get("submitted_by") == current_user]
        if q.strip():
            ql = q.lower().strip()
            mine = [r for r in mine if ql in " ".join(str(v) for v in r.values()).lower()]

        if not mine:
            st.info("📋 You have not submitted any work orders yet.")
        else:
            for r in reversed(mine):
                status_raw = str(r.get("status", "")).strip().lower()
                display_status = {
                    "pending_manager": "PENDING MANAGER",
                    "pending_director": "PENDING DIRECTOR",
                    "approved_payment": "APPROVED",
                    "approved": "APPROVED",
                    "rejected_director": "REJECTED BY DIRECTOR",
                    "rejected": "REJECTED",
                    "returned_to_employee": "RETURNED — NEEDS EDITS",
                }.get(status_raw, status_raw.replace("_", " ").upper())

                icon = {
                    "pending_manager": "🟡",
                    "pending_director": "🟡",
                    "approved_payment": "🟢",
                    "approved": "🟢",
                    "rejected_director": "🔴",
                    "rejected": "🔴",
                    "returned_to_employee": "🟠",
                }.get(status_raw, "⚪")

                with st.expander(
                    f"{icon} {get_work_order_number(r)} | {r.get('emp_name')} | "
                    f"£{r.get('amount', 0):.2f} | {display_status}"
                ):
                    st.write(f"🧾 Work Order No.: **{get_work_order_number(r)}** | 👔 Manager: {r.get('manager')} | 🏢 {r.get('dept')} | 📅 {r.get('work_date')}")
                    st.write(f"💷 £{r.get('amount', 0):.2f}")
                    if r.get("site_address"):
                        st.write(f"📍 Site Address: {r.get('site_address')}")
                    if r.get("customer_job_no"):
                        st.write(f"📘 Customer Job No.: {r.get('customer_job_no')}")
                    st.info(f"📝 {r.get('desc')}")
                    display_attachments(r)

                    if r.get("manager_decision_by"):
                        st.write(f"👔 Manager reviewed by: {r.get('manager_decision_by')} on {r.get('manager_decision_date')}")
                    if r.get("manager_comments"):
                        if status_raw == "returned_to_employee":
                            st.error(f"❌ Manager Comments: {r.get('manager_comments')}")
                        else:
                            st.info(f"💬 Manager Comments: {r.get('manager_comments')}")
                    if r.get("director_decision_by"):
                        st.write(f"🎯 Approved By: {r.get('director_decision_by')} on {r.get('director_decision_date')}")
                    if r.get("director_comments"):
                        st.warning(f"💬 Director Comments: {r.get('director_comments')}")

                    if status_raw in ("pending_manager", "returned_to_employee", "rejected_director", "rejected", "pending"):
                        st.divider()
                        if st.button(
                            f"✏️ Edit Work Order {get_work_order_number(r)}",
                            key=f"wo_emp_edit_{r.get('id')}",
                            type="secondary"
                        ):
                            st.session_state.editing_work_order_id_emp = r.get("id")
                            st.rerun()

    with tab_dept:
        st.caption(
            "Work orders already recorded for your department (including those submitted by managers). "
            "👉 Check the Work Order No. here **before** submitting a new one to avoid duplicates."
        )

        dept_orders = [
            r for r in orders
            if str(r.get("dept", "")).strip() == str(current_dept).strip()
        ]

        dq = st.text_input(
            "🔎 Search department work orders",
            placeholder="Search by Work Order No., employee, manager, status, amount or date...",
            key="wo_emp_dept_search"
        )
        if dq.strip():
            dql = dq.lower().strip()
            dept_orders = [r for r in dept_orders if dql in " ".join(str(v) for v in r.values()).lower()]

        if not dept_orders:
            st.info("📋 No work orders recorded for your department yet.")
        else:
            rows = []
            for r in sorted(dept_orders, key=lambda x: str(x.get("work_date", "")), reverse=True):
                rows.append({
                    "Work Order No.": get_work_order_number(r),
                    "Employee": r.get("emp_name", ""),
                    "Manager": r.get("manager", ""),
                    "Work Date": r.get("work_date", ""),
                    "Amount (£)": float(r.get("amount", 0) or 0),
                    "Status": str(r.get("status", "")).replace("_", " ").upper(),
                    "Submitted By": r.get("submitted_by", ""),
                })
            df_view = pd.DataFrame(rows)
            st.dataframe(df_view, use_container_width=True, hide_index=True)
            st.caption(f"📊 {len(rows)} work order(s) in **{current_dept}**")

            st.divider()
            st.markdown("#### 📄 Full Details")
            for r in reversed(dept_orders):
                status_raw = str(r.get("status", "")).strip().lower()
                display_status = {
                    "pending_manager": "PENDING MANAGER",
                    "pending_director": "PENDING DIRECTOR",
                    "approved_payment": "APPROVED",
                    "approved": "APPROVED",
                    "rejected_director": "REJECTED",
                    "rejected": "REJECTED",
                    "returned_to_employee": "RETURNED",
                }.get(status_raw, status_raw.replace("_", " ").upper())

                with st.expander(
                    f"🧾 {get_work_order_number(r)} | {r.get('emp_name')} | "
                    f"£{float(r.get('amount', 0) or 0):.2f} | {display_status}"
                ):
                    st.write(f"👔 **Manager:** {r.get('manager', '-')} | 📅 **Date:** {r.get('work_date', '-')}")
                    st.write(f"📝 **Submitted by:** {r.get('submitted_by', '-')}")
                    if r.get("site_address"):
                        st.write(f"📍 **Site Address:** {r.get('site_address')}")
                    if r.get("customer_job_no"):
                        st.write(f"📘 **Customer Job No.:** {r.get('customer_job_no')}")
                    st.info(f"📝 {r.get('desc', '')}")

# ✅ v4.7 — Manager portal: added "Submitted by" display, kept v4.6 approve/reject + edit fix
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
                    try:
                        _d = datetime.strptime(str(rec.get("work_date", ""))[:10], "%Y-%m-%d").date()
                    except Exception:
                        _d = date.today()
                    e_date = st.date_input("📅 Date", value=_d)
                    e_amount = st.number_input(
                        "💷 Amount (£)", min_value=0.01, step=1.0, format="%.2f",
                        value=max(float(rec.get("amount", 0.01) or 0.01), 0.01)
                    )
                    e_desc = st.text_area("📝 Description", value=str(rec.get("desc", "")), height=150)

                csave, ccancel = st.columns(2)
                with csave:
                    save_edit = st.form_submit_button("💾 Save Changes", type="primary", use_container_width=True)
                with ccancel:
                    cancel_edit = st.form_submit_button("❌ Cancel", use_container_width=True)

                if save_edit:
                    old_data = {
                        "manual_work_order_no": rec.get("manual_work_order_no"),
                        "emp_name": rec.get("emp_name"),
                        "site_address": rec.get("site_address"),
                        "customer_job_no": rec.get("customer_job_no"),
                        "work_date": rec.get("work_date"),
                        "amount": rec.get("amount"),
                        "desc": rec.get("desc"),
                    }
                    for x in orders:
                        if str(x.get("id")) == str(editing_id):
                            x["manual_work_order_no"] = e_wo_no.strip() or str(x.get("id"))
                            x["emp_name"] = e_emp.strip()
                            x["site_address"] = e_site.strip()
                            x["customer_job_no"] = e_cjno.strip()
                            x["work_date"] = str(e_date)
                            x["amount"] = float(e_amount)
                            x["desc"] = e_desc.strip()
                            x["pdf_path"] = ""
                            break
                    save_all_work_orders(orders)
                    new_data = {
                        "manual_work_order_no": e_wo_no.strip(),
                        "emp_name": e_emp.strip(),
                        "site_address": e_site.strip(),
                        "customer_job_no": e_cjno.strip(),
                        "work_date": str(e_date),
                        "amount": float(e_amount),
                        "desc": e_desc.strip(),
                    }
                    log_action("WORK_ORDER_EDITED", editing_id, old_data=old_data, new_data=new_data)
                    st.session_state.editing_work_order_id = None
                    st.success("✅ Work Order updated.")
                    st.rerun()

                if cancel_edit:
                    st.session_state.editing_work_order_id = None
                    st.rerun()
            st.divider()
        else:
            st.session_state.editing_work_order_id = None

    manager_orders = [
        r for r in orders
        if str(r.get("manager", "")).strip() == str(manager_name).strip()
        or str(r.get("submitted_by", "")).strip() == str(manager_name).strip()
    ]
    pending = [r for r in manager_orders if r.get("status") in ("pending_director", "pending_manager", "pending")]
    approved = [r for r in manager_orders if r.get("status") in ("approved_payment", "approved")]
    rejected = [r for r in manager_orders if r.get("status") in ("rejected_director", "rejected", "returned_to_employee")]

    main_tab, total_tab = st.tabs(["🛠️ Work Orders", "💷 Approved Work Order Total"])

    with main_tab:
        st.markdown("### 📤 Submit New Work Order")
        st.caption("Complete the work-order details below. No hours/time entry is required.")
        with st.form("manager_new_work_order_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                work_order_no = st.text_input("🧾 Work Order No.", placeholder="Enter the Work Order No. from the work-order sheet")
                contractor_employee = st.text_input("👤 Contractor / Employee Labour", placeholder="Enter contractor or employee name")
                site_address = st.text_area("📍 Site Address", placeholder="Enter the full site address", height=90)
                customer_job_no = st.text_input("📘 Customer Job No.", placeholder="Enter Customer Job No.")
            with c2:
                work_date = st.date_input("📅 Date", value=date.today())
                amount = st.number_input("💷 Amount (£)", min_value=0.01, step=1.0, format="%.2f")
                description = st.text_area("📝 Description", placeholder="Describe the work completed...", height=150)
            _all_users = load_users()
            director_names = sorted({
                str(u.get("full_name", "")).strip()
                for u in _all_users.values()
                if str(u.get("role", "")).strip().lower() == "director"
                and str(u.get("full_name", "")).strip()
            })
            director = director_names[0] if director_names else "Andy Acoole"
            files = st.file_uploader("📎 Supporting Work Order Document (optional)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)
            submitted = st.form_submit_button("📤 Submit Work Order", type="primary", use_container_width=True)
        if submitted:
            errors=[]
            if not work_order_no.strip(): errors.append("Work Order No.")
            if not contractor_employee.strip(): errors.append("Contractor / Employee Labour")
            if not site_address.strip(): errors.append("Site Address")
            if not customer_job_no.strip(): errors.append("Customer Job No.")
            if not description.strip(): errors.append("Description")
            duplicate=any(
                str(r.get("manual_work_order_no","")).strip().lower()==work_order_no.strip().lower()
                for r in orders if str(r.get("manual_work_order_no","")).strip()
            )
            if duplicate: errors.append("Work Order No. already exists")
            if errors:
                st.error("Please correct: " + ", ".join(errors) + ".")
            else:
                wid = get_next_work_order_id(orders)
                attachments=[]
                for i,f in enumerate(files or [],1):
                    safe_name=os.path.basename(f.name).replace("/","_").replace("\\","_")
                    fn=f"{wid}_F{i}_{safe_name}"
                    fp=os.path.join(UPLOAD_DIR,fn)
                    with open(fp,"wb") as out_file:
                        out_file.write(f.getbuffer())
                    _upload_to_drive_bg(fp,fn)
                    attachments.append(fn)
                now=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                rec={
                    "id":wid,"manual_work_order_no":work_order_no.strip(),"emp_name":contractor_employee.strip(),
                    "dept":manager_dept,"work_date":str(work_date),"hours":0.0,
                    "customer_job_no":customer_job_no.strip(),"site_address":site_address.strip(),
                    "amount":float(amount),"manager":manager_name,"desc":description.strip(),
                    "attachment_name":", ".join(attachments) or "None","status":"pending_director",
                    "manager_comments":"","manager_decision_date":"","manager_decision_by":"",
                    "director_comments":"","director_decision_date":"","director_decision_by":"",
                    "submitted_by":manager_name,"submitted_date":now,"payroll_status":"Pending",
                    "payroll_date":"","payroll_by":"","pdf_path":"",
                }
                orders.append(rec)
                save_all_work_orders(orders)
                log_action("WORK_ORDER_MANAGER_CREATED",wid,decision_by=manager_name)
                st.success(f"✅ Work Order {work_order_no.strip()} submitted to {director} for final approval.")
                st.rerun()
        st.divider()
        wo_pending, wo_approved, wo_rejected = st.tabs([
            f"⏳ Pending ({len(pending)})", f"✅ Approved ({len(approved)})", f"❌ Returned / Rejected ({len(rejected)})"
        ])
        def filter_orders(items,key):
            q=st.text_input("🔎 Search Work Orders", placeholder="Search by Work Order No., employee, date, amount or description...", key=key).strip().lower()
            if not q: return items
            return [r for r in items if q in " ".join(str(v) for v in r.values()).lower()]
        def show_details(r):
            wo=r.get("manual_work_order_no") or r.get("id") or "-"
            st.write(f"🧾 **Work Order No.: {wo}** | 👤 **Contractor / Employee Labour:** {r.get('emp_name','-')}")
            st.write(f"📍 **Site Address:** {r.get('site_address','-')}")
            st.write(f"📘 **Customer Job No.:** {r.get('customer_job_no','-')} | 📅 **Date:** {r.get('work_date','-')}")
            st.write(f"💷 **Amount:** £{float(r.get('amount',0) or 0):.2f}")
            submitter = str(r.get("submitted_by", "") or "").strip() or "Unknown"
            sub_date = str(r.get("submitted_date", "") or "").strip()
            st.write(f"📝 **Submitted by:** {submitter}" + (f" on {sub_date}" if sub_date else ""))
            st.info(f"📝 **Description:**\n{r.get('desc','')}")
            st.write(f"👔 **Manager:** {r.get('manager') or manager_name}")
            if r.get("manager_decision_by"):
                st.write(f"👔 **Manager Review:** {r.get('manager_decision_by')} on {r.get('manager_decision_date','')}")
            if r.get("manager_comments"):
                st.info(f"💬 **Manager Comments:** {r.get('manager_comments')}")
            if r.get("director_decision_by"):
                st.write(f"🎯 **Approved By:** {r.get('director_decision_by')} on {r.get('director_decision_date','')}")
            if r.get("director_comments"):
                st.warning(f"💬 Director Comments: {r.get('director_comments')}")
            display_attachments(r)

        with wo_pending:
            items = filter_orders(pending, "wo_mgr_pending_final_search")
            if not items:
                st.info("⏳ No pending work orders.")
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
                                        x["status"] = "pending_director"
                                        x["manager_comments"] = mgr_comment.strip()
                                        x["manager_decision_by"] = manager_name
                                        x["manager_decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                        x["pdf_path"] = ""
                                        break
                                save_all_work_orders(orders)
                                log_action("WORK_ORDER_MANAGER_APPROVED", r.get("id"), decision_by=manager_name)
                                st.success(f"✅ Work Order {wo} approved and sent to Director.")
                                st.rerun()
                        with c2:
                            if st.button("❌ Reject & Return to Employee", key=f"wo_mgr_reject_{r.get('id')}"):
                                for x in orders:
                                    if str(x.get("id")) == str(r.get("id")):
                                        x["status"] = "returned_to_employee"
                                        x["manager_comments"] = mgr_comment.strip()
                                        x["manager_decision_by"] = manager_name
                                        x["manager_decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                        x["pdf_path"] = ""
                                        break
                                save_all_work_orders(orders)
                                log_action("WORK_ORDER_MANAGER_REJECTED", r.get("id"), decision_by=manager_name)
                                st.warning(f"❌ Work Order {wo} returned to employee for corrections.")
                                st.rerun()
                    else:
                        st.caption("Awaiting Director approval. You may still edit this record.")
                        if st.button(f"✏️ Edit Work Order {wo}", key=f"wo_mgr_pending_edit_{r.get('id')}"):
                            st.session_state.editing_work_order_id = r.get("id")
                            st.rerun()

        with wo_approved:
            items=filter_orders(approved,"wo_mgr_approved_final_search")
            if not items:
                st.info("✅ No approved work orders.")
            else:
                for r in reversed(items):
                    wo=r.get("manual_work_order_no") or r.get("id")
                    submitter = str(r.get("submitted_by", "") or "").strip() or "Unknown"
                    with st.expander(f"✅ {wo} | {r.get('emp_name')} | £{float(r.get('amount',0) or 0):.2f} | Submitted by: {submitter}"):
                        show_details(r)
                        st.success("Approved — read only.")
                        st.divider()
                        st.markdown("#### 📄 Work Order PDF")
                        display_work_order_pdf(r)
                st.divider()
                render_work_order_bulk_download(load_work_orders(), key_prefix="wo_mgr_bulk")

        with wo_rejected:
            items=filter_orders(rejected,"wo_mgr_rejected_final_search")
            if not items: st.info("❌ No rejected or returned work orders.")
            for r in reversed(items):
                wo=r.get("manual_work_order_no") or r.get("id")
                st_raw = str(r.get("status", "")).strip().lower()
                submitter = str(r.get("submitted_by", "") or "").strip() or "Unknown"
                lbl = "RETURNED TO EMPLOYEE" if st_raw == "returned_to_employee" else "REJECTED BY DIRECTOR"
                with st.expander(f"❌ {wo} | {r.get('emp_name')} | £{float(r.get('amount',0) or 0):.2f} | {lbl} | Submitted by: {submitter}"):
                    show_details(r)
                    st.warning("Editable — you can update and resubmit.")
                    if st.button(f"✏️ Edit & Resubmit {wo}",key=f"wo_mgr_rejected_edit_{r.get('id')}"):
                        st.session_state.editing_work_order_id=r.get("id")
                        st.rerun()

    with total_tab:
        if show_total:
            renderer = globals().get("render_work_order_total")
            if renderer:
                renderer(load_work_orders(), scope_department=manager_dept, key_prefix="wo_mgr_total_tab", prepared_by=manager_name)
            else:
                st.info("Approved Work Order Total is available in this tab.")
        else:
            st.info("Approved Work Order Total is available from the Work Orders total tab.")

def render_work_order_director_portal(director_name):
    st.subheader("🛠️ Work Orders — Director Final Approval")
    st.info("✅ Review all work orders, Approve, Reject, OR Change Status at any time. All changes are logged.")
    orders = load_work_orders()
    pending = [r for r in orders if r.get("status") == "pending_director"]
    approved = [r for r in orders if r.get("status") == "approved_payment"]
    rejected = [r for r in orders if r.get("status") == "rejected_director"]
    t1, t2, t3 = st.tabs([
        f"⏳ Manager Approved / Awaiting Director ({len(pending)})",
        f"✅ Approved for Payment ({len(approved)})",
        f"❌ Rejected ({len(rejected)})"
    ])
    def search_list(items, key):
        q = st.text_input("🔎 Search Work Orders",
                          placeholder="Search by Work Order No., employee, manager, submitter, amount, date, customer job no. or description...",
                          key=key)
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
        if r.get('manager_decision_by'):
            st.write(f"👔 **Manager review:** {r.get('manager_decision_by')} on {r.get('manager_decision_date','')}")
        st.info(f"📝 **Description:**\n{r.get('desc','')}")
        if r.get('manager_comments'):
            st.info(f"💬 **Manager Comments:** {r.get('manager_comments')}")
        if r.get('director_comments'):
            st.warning(f"💬 **Director Comments:** {r.get('director_comments')}")
        if r.get('director_decision_by'):
            st.write(f"🎯 **Approved By:** {r.get('director_decision_by')} on {r.get('director_decision_date','')}")
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
                    x["director_comments"] = (str(x.get("director_comments","")) +
                        f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M')}] ⏳ Changed to Pending by {director_name}: {comments.strip()}").strip()
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
    st.divider()
    render_work_order_total(orders, key_prefix="wo_dir_total", prepared_by=director_name)

def render_work_order_payroll_portal(payroll_name):
    st.subheader("🛠️ Work Orders — Payroll")
    st.info("View Director-approved work orders and download the authorised PDF. Payroll does not change the work order status.")
    orders = load_work_orders()
    approved = [r for r in orders if r.get("status") == "approved_payment"]
    search = st.text_input(
        "🔎 Search approved work orders",
        placeholder="Search by ID, employee, manager, director, department, amount or date...",
        key="wo_payroll_search"
    )
    if search.strip():
        q = search.lower().strip()
        approved = [r for r in approved if q in " ".join(str(v) for v in r.values()).lower()]
    st.metric("✅ Approved for Payment", len(approved))
    st.divider()
    if not approved:
        st.success("✅ No Director-approved work orders found.")
        return
    for r in reversed(approved):
        with st.expander(
            f"🟢 {r.get('id')} | {r.get('emp_name')} | "
            f"£{r.get('amount', 0):.2f} | Approved by {r.get('director_decision_by')}"
        ):
            st.write(f"🧾 Work Order No.: **{get_work_order_number(r)}**")
            st.write(
                f"📝 Submitted by: {r.get('submitted_by')} | "
                f"👔 Manager: {r.get('manager')} | "
                f"🎯 Director: {r.get('director_decision_by')}"
            )
            st.write(
                f"📅 Approval: {r.get('director_decision_date')} | "
                f"💷 £{r.get('amount', 0):.2f}"
            )
            st.info(r.get('desc', ''))
            display_attachments(r)
            display_work_order_pdf(r)
    st.divider()
    render_work_order_bulk_download(load_work_orders(), key_prefix="wo_pay_bulk")
    st.divider()
    render_work_order_total(approved, key_prefix="wo_pay_total", prepared_by=payroll_name)

# ============================================================
# NATIONAL GRID INSPECTOR BONUS
# ============================================================
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
                "total_jobs": jobs,
                "bonus_amount": amount,
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
        try: nums.append(int(raw.replace("IB-", "")))
        except: pass
    return f"IB-{max(nums) + 1 if nums else 1:04d}"

def inspector_bonus_pdf(req, force_regenerate=False, upload_to_drive=True):
    if not PDF_AVAILABLE:
        return None
    try:
        cached_path = req.get("pdf_path", "")
        if not force_regenerate and cached_path and os.path.exists(cached_path):
            return cached_path

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
            if family == "Helvetica":
                return text.encode("latin-1", "replace").decode("latin-1")
            return text

        if os.path.exists(LOGO_PATH):
            try:
                pdf.image(LOGO_PATH, x=75, y=10, w=60)
                pdf.ln(28)
            except Exception:
                pdf.ln(5)
        else:
            pdf.ln(5)

        pdf.set_font(family, "B", 16)
        pdf.cell(0, 10, safe("National Grid Inspector Bonus Approval Sheet"), ln=True, align="C")
        pdf.ln(4)

        pdf.set_font(family, "", 10)
        pdf.multi_cell(
            0, 6,
            safe("This sheet needs to be completed and passed to Andy to be signed off and given to Rachel by the 3rd of the month."),
            align="C"
        )
        pdf.ln(8)

        def field(label, value, label_w=60, value_h=10):
            pdf.set_font(family, "B", 11)
            pdf.cell(label_w, value_h, safe(label), border=1)
            pdf.set_font(family, "", 11)
            pdf.cell(0, value_h, safe("   " + str(value)), border=1, ln=True)
            pdf.ln(2)

        field("Inspector Name:", req.get("inspector_name", ""))
        field("Month & Year:", req.get("month_year", ""))
        field("Days Absent:", req.get("days_absent", "") or "-")
        field("Reasons for Absence:", req.get("reasons", "") or "-")
        field("Total Jobs Completed", f"{float(req.get('total_jobs', 0)):.2f}")
        field("Bonus Amount:", f"£{float(req.get('bonus_amount', 0)):.2f}")

        status = str(req.get("status", "")).strip().lower()
        text_content = ""
        stamp_path = None

        if status == "approved":
            approved_by = req.get("director_decision_by", "") or "Andy Acoole"
            decision_date = req.get("director_decision_date", "")
            text_content = f"   {approved_by} on {decision_date}"
            stamp_path = APPROVED_STAMP_PATH
        elif status == "rejected":
            text_content = "   Rejected"
            stamp_path = REJECTED_STAMP_PATH
        else:
            text_content = "   (Pending Director signature)"

        pdf.set_font(family, "B", 11)
        pdf.cell(60, 10, safe("Approved By:"), border=1)

        pdf.set_font(family, "", 11)
        pdf.cell(80, 10, safe(text_content), border=1)

        current_x = pdf.get_x()
        current_y = pdf.get_y()
        pdf.cell(50, 10, "", border=1, ln=True)

        if stamp_path and os.path.exists(stamp_path):
            try:
                pdf.image(stamp_path, x=current_x + 5, y=current_y - 1, w=40)
            except Exception:
                pass

        pdf.ln(4)

        if req.get("director_comments"):
            pdf.set_font(family, "B", 10)
            pdf.cell(0, 6, safe("Director Comments:"), ln=True)
            pdf.set_font(family, "", 10)
            pdf.set_x(pdf.l_margin)
            pdf.multi_cell(0, 6, safe(req.get("director_comments", "")))
            pdf.ln(4)

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
        if upload_to_drive:
            _upload_to_drive_bg(path, os.path.basename(path))
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
        st.info("📄 PDF download is available once the Director approves this bonus.")
        return
    req_id = str(req.get("id", "unknown"))
    dl_key = f"{key_prefix}_dl_{req_id}"
    gen_key = f"{key_prefix}_gen_{req_id}"

    cached_path = req.get("pdf_path", "")
    if cached_path and os.path.exists(cached_path):
        with open(cached_path, "rb") as f:
            st.download_button(
                "⬇️ Download Bonus Approval PDF",
                data=f.read(),
                file_name=os.path.basename(cached_path),
                mime="application/pdf",
                type="primary",
                key=dl_key
            )
        return

    if st.button(f"📄 Generate PDF for {req_id}", key=gen_key, type="primary"):
        with st.spinner("Generating PDF..."):
            path = inspector_bonus_pdf(req, force_regenerate=True, upload_to_drive=True)
        if path and os.path.exists(path):
            records = load_inspector_bonus()
            for r in records:
                if str(r.get("id")) == req_id:
                    r["pdf_path"] = path
            save_all_inspector_bonus(records, sync=False)
            st.success("✅ PDF generated. Click below to download.")
            st.rerun()
        else:
            st.error("❌ Could not generate PDF.")

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
                with c_save:
                    save_edit = st.form_submit_button("💾 Save Changes & Resubmit", type="primary", use_container_width=True)
                with c_cancel:
                    cancel_edit = st.form_submit_button("❌ Cancel", use_container_width=True)

                if save_edit:
                    if not e_inspector.strip() or not e_month.strip():
                        st.error("❌ Inspector Name and Month & Year are required.")
                    else:
                        old_data = {
                            "inspector_name": rec.get("inspector_name"),
                            "month_year": rec.get("month_year"),
                            "days_absent": rec.get("days_absent"),
                            "reasons": rec.get("reasons"),
                            "total_jobs": rec.get("total_jobs"),
                            "bonus_amount": rec.get("bonus_amount"),
                        }
                        for x in bonus_records:
                            if str(x.get("id")) == str(editing_id):
                                x["inspector_name"] = e_inspector.strip()
                                x["month_year"] = e_month.strip()
                                x["days_absent"] = e_days.strip()
                                x["reasons"] = e_reasons.strip()
                                x["total_jobs"] = float(e_jobs)
                                x["bonus_amount"] = float(e_amount)
                                x["status"] = "pending_director"
                                x["director_comments"] = ""
                                x["director_decision_date"] = ""
                                x["director_decision_by"] = ""
                                x["pdf_path"] = ""
                                x["submitted_by"] = user_name
                                x["submitted_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                break
                        save_all_inspector_bonus(bonus_records)
                        new_data = {
                            "inspector_name": e_inspector.strip(),
                            "month_year": e_month.strip(),
                            "days_absent": e_days.strip(),
                            "reasons": e_reasons.strip(),
                            "total_jobs": float(e_jobs),
                            "bonus_amount": float(e_amount),
                        }
                        log_action("INSPECTOR_BONUS_EDITED", editing_id, old_data=old_data, new_data=new_data)
                        st.session_state.editing_inspector_bonus_id = None
                        st.success(f"✅ Bonus record {editing_id} updated and re-sent to Director.")
                        st.rerun()

                if cancel_edit:
                    st.session_state.editing_inspector_bonus_id = None
                    st.rerun()
            st.divider()
        else:
            st.session_state.editing_inspector_bonus_id = None

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
            if not inspector_name.strip():
                st.error("❌ Inspector Name is required.")
            elif not month_year.strip():
                st.error("❌ Month & Year is required.")
            else:
                rec = {
                    "id": new_id,
                    "inspector_name": inspector_name.strip(),
                    "month_year": month_year.strip(),
                    "days_absent": days_absent.strip(),
                    "reasons": reasons.strip(),
                    "total_jobs": float(total_jobs),
                    "bonus_amount": float(bonus_amount),
                    "status": "pending_director",
                    "director_comments": "",
                    "director_decision_date": "",
                    "director_decision_by": "",
                    "submitted_by": user_name,
                    "submitted_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "pdf_path": ""
                }
                bonus_records.append(rec)
                save_all_inspector_bonus(bonus_records)
                log_action("INSPECTOR_BONUS_CREATED", new_id, new_data=rec)
                st.success(f"✅ Bonus Approval {new_id} for {inspector_name} submitted to Director for approval.")
                st.rerun()

    st.divider()
    st.subheader("📋 Submitted Bonus Approvals")
    search = st.text_input("🔎 Search records", placeholder="Search by ID, inspector, month, amount...", key="ib_search")
    filtered = list(bonus_records)
    if search.strip():
        q = search.lower().strip()
        filtered = [r for r in filtered if q in " ".join(str(v) for v in r.values()).lower()]

    if not filtered:
        st.info("📋 No bonus approval records found.")
    else:
        for r in reversed(filtered):
            status_raw = str(r.get("status", "")).strip().lower()
            if status_raw == "approved":
                icon = "🟢"; status_label = "APPROVED"
            elif status_raw == "rejected":
                icon = "🔴"; status_label = "REJECTED"
            else:
                icon = "🟡"; status_label = "PENDING DIRECTOR"

            with st.expander(f"{icon} {r.get('id')} | {r.get('inspector_name')} | {r.get('month_year')} | £{r.get('bonus_amount',0):.2f} | {status_label}"):
                st.write(f"**Inspector:** {r.get('inspector_name')} | **Month:** {r.get('month_year')}")
                st.write(f"**Days Absent:** {r.get('days_absent') or '-'}")
                st.write(f"**Reasons:** {r.get('reasons') or '-'}")
                st.write(f"**Total Jobs Completed:** {r.get('total_jobs')}")
                st.write(f"**Bonus Amount:** £{r.get('bonus_amount',0):.2f}")
                st.write(f"**Status:** {status_label}")
                if r.get("director_decision_by"):
                    st.write(f"**Director:** {r.get('director_decision_by')} on {r.get('director_decision_date','')}")
                if r.get("director_comments"):
                    st.info(f"💬 **Director Comments:** {r.get('director_comments')}")
                st.caption(f"Submitted by {r.get('submitted_by')} on {r.get('submitted_date')}")
                display_inspector_bonus_pdf_button(r, key_prefix="ib_mgr")

                if status_raw == "pending_director":
                    st.divider()
                    if st.button(f"✏️ Edit Bonus Record {r.get('id')}", key=f"ib_edit_{r.get('id')}", type="secondary"):
                        st.session_state.editing_inspector_bonus_id = r.get("id")
                        st.rerun()

def render_inspector_bonus_director_portal(director_name):
    st.subheader("💰 National Grid Inspector Bonus — Director Approval")
    st.info("Review Inspector Bonus submissions. Approve, Reject, OR Change Status at any time. All changes are logged.")
    st.divider()

    records = load_inspector_bonus()
    pending = [r for r in records if str(r.get("status", "")).strip().lower() == "pending_director"]
    approved = [r for r in records if str(r.get("status", "")).strip().lower() == "approved"]
    rejected = [r for r in records if str(r.get("status", "")).strip().lower() == "rejected"]

    t1, t2, t3 = st.tabs([
        f"⏳ Pending ({len(pending)})",
        f"✅ Approved ({len(approved)})",
        f"❌ Rejected ({len(rejected)})"
    ])

    def apply_decision(rec_id, new_status, comments, old_status):
        for x in records:
            if str(x.get("id")) == str(rec_id):
                x["status"] = new_status
                if new_status == "pending_director":
                    x["director_comments"] = (str(x.get("director_comments","")) +
                        f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M')}] ⏳ Changed to Pending by {director_name}: {comments.strip()}").strip()
                    x["director_decision_by"] = ""
                    x["director_decision_date"] = ""
                else:
                    x["director_comments"] = comments.strip() if comments.strip() else x.get("director_comments", "")
                    x["director_decision_by"] = director_name
                    x["director_decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                x["pdf_path"] = ""
                break
        save_all_inspector_bonus(records)
        if old_status == "pending_director":
            action = "INSPECTOR_BONUS_APPROVED" if new_status == "approved" else ("INSPECTOR_BONUS_REJECTED" if new_status == "rejected" else "INSPECTOR_BONUS_STATUS_CHANGED")
        else:
            action = "INSPECTOR_BONUS_STATUS_CHANGED"
        log_action(action, rec_id,
                   old_data={"status": old_status}, new_data={"status": new_status, "comments": comments},
                   decision_by=director_name)
        st.success(f"✅ Record {rec_id} → **{new_status.replace('_',' ').title()}**.")

    def show_details(r):
        st.write(f"**Inspector:** {r.get('inspector_name')} | **Month:** {r.get('month_year')}")
        st.write(f"**Days Absent:** {r.get('days_absent') or '-'}")
        st.write(f"**Reasons:** {r.get('reasons') or '-'}")
        st.write(f"**Total Jobs Completed:** {r.get('total_jobs')}")
        st.write(f"**Bonus Amount:** £{r.get('bonus_amount',0):.2f}")
        st.caption(f"Submitted by {r.get('submitted_by')} on {r.get('submitted_date')}")
        if r.get("director_decision_by"):
            st.write(f"**Decision By:** {r.get('director_decision_by')} on {r.get('director_decision_date','')}")
        if r.get("director_comments"):
            st.info(f"💬 **Director Comments:** {r.get('director_comments')}")

    with t1:
        if not pending:
            st.success("✅ No pending Inspector Bonus submissions.")
        for r in reversed(pending):
            rec_id = r.get("id")
            with st.expander(f"🟡 {rec_id} | {r.get('inspector_name')} | {r.get('month_year')} | £{r.get('bonus_amount',0):.2f}"):
                show_details(r)
                st.divider()
                comments = st.text_area("Director Comments", key=f"ib_dir_comm_{rec_id}")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("✅ Approve", key=f"ib_dir_app_{rec_id}", type="primary"):
                        apply_decision(rec_id, "approved", comments, "pending_director")
                        st.rerun()
                with c2:
                    if st.button("❌ Reject", key=f"ib_dir_rej_{rec_id}"):
                        apply_decision(rec_id, "rejected", comments, "pending_director")
                        st.rerun()

    with t2:
        if not approved:
            st.info("✅ No approved Inspector Bonus records yet.")
        if approved:
            st.info("🔄 **Change Status:** Move back to Pending ❘ Change to Rejected")
        for r in reversed(approved):
            rec_id = r.get("id")
            with st.expander(f"🟢 {rec_id} | {r.get('inspector_name')} | {r.get('month_year')} | £{r.get('bonus_amount',0):.2f}"):
                show_details(r)
                st.divider()
                display_inspector_bonus_pdf_button(r, key_prefix="ib_dir")
                st.divider()
                st.markdown("### 🔄 Change Status")
                new_comments = st.text_area(
                    "Add comment (optional)", key=f"ib_dir_chg_app_{rec_id}",
                    placeholder="Reason for status change..."
                )
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("⏳ Move to Pending", key=f"ib_dir_app_to_pend_{rec_id}"):
                        apply_decision(rec_id, "pending_director", new_comments, "approved")
                        st.rerun()
                with c2:
                    if st.button("❌ Change to Rejected", key=f"ib_dir_app_to_rej_{rec_id}"):
                        apply_decision(rec_id, "rejected", new_comments, "approved")
                        st.rerun()

    with t3:
        if not rejected:
            st.info("❌ No rejected Inspector Bonus records.")
        if rejected:
            st.info("🔄 **Change Status:** Move back to Pending ❘ Change to Approved")
        for r in reversed(rejected):
            rec_id = r.get("id")
            with st.expander(f"🔴 {rec_id} | {r.get('inspector_name')} | {r.get('month_year')} | £{r.get('bonus_amount',0):.2f}"):
                show_details(r)
                st.divider()
                st.markdown("### 🔄 Change Status")
                new_comments = st.text_area(
                    "Add comment (optional)", key=f"ib_dir_chg_rej_{rec_id}",
                    placeholder="Reason for status change..."
                )
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("⏳ Move to Pending", key=f"ib_dir_rej_to_pend_{rec_id}"):
                        apply_decision(rec_id, "pending_director", new_comments, "rejected")
                        st.rerun()
                with c2:
                    if st.button("✅ Change to Approved", key=f"ib_dir_rej_to_app_{rec_id}", type="primary"):
                        apply_decision(rec_id, "approved", new_comments, "rejected")
                        st.rerun()

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
            if r.get("director_comments"):
                st.info(f"💬 **Director Comments:** {r.get('director_comments')}")
            st.caption(f"Submitted by {r.get('submitted_by')} on {r.get('submitted_date')}")
            st.divider()
            display_inspector_bonus_pdf_button(r, key_prefix="ib_payroll")

# ============================================================
# REQUESTS EXCEL
# ============================================================
def initialise_excel():
    safe_init_excel(EXCEL_PATH, EXCEL_COLUMNS)

initialise_excel()
initialise_work_orders()
initialise_inspector_bonus()

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
            try: record_id = int(r.get("ID", 0))
            except: record_id = 0
            try: amount = float(r.get("Amount (£)", 0))
            except: amount = 0.0
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
                "old_data": str(r.get("Old Data", "")).strip(),
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
            "Old Data": str(r.get("old_data", "")),
        })
    pd.DataFrame(export, columns=EXCEL_COLUMNS).to_excel(EXCEL_PATH, index=False, engine="openpyxl")
    _set_data_cache("_records_cache", list(records))
    sync_saved_file_to_drive(EXCEL_PATH)

def save_record_to_excel(new_record):
    current = load_records_from_excel()
    current.append(new_record)
    save_all_records(current)

# ============================================================
# REQUESTER / COMPLETION DISPLAY HELPERS
# ============================================================
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
# PDF GENERATION
# ============================================================
def generate_approval_pdf(request_data):
    if not PDF_AVAILABLE:
        return False, None, "Install fpdf2: pip install fpdf2"
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
        pdf.line(10, line_y, 200, line_y)
        pdf.line(10, line_y + 1.5, 200, line_y + 1.5)
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
                    else:
                        pdf.cell(0, 5, "     Non-image file - see original upload", ln=True); pdf.ln(3)
                else:
                    pdf.cell(0, 5, "     Warning: File not found on server", ln=True); pdf.ln(3)
        else:
            pdf.cell(0, 6, "- No files were attached to this request", ln=True)
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
    if not filename or str(filename).strip().lower() in ("none", "nan", ""):
        return None
    filename = os.path.basename(str(filename).strip())
    local_path = os.path.join(UPLOAD_DIR, filename)
    if os.path.exists(local_path):
        return local_path
    if drive_service is not None:
        remote = _drive_find_file(filename)
        if remote and _drive_download_file(remote["id"], local_path):
            return local_path
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
                with open(path, "rb") as f:
                    file_data = f.read()
                st.download_button(label=f"⬇️ Download {name}", data=file_data, file_name=name, key=f"attachment_{req.get('id', idx)}_{idx}")
        if not found_any:
            st.info("📎 Attachments referenced but files are not available.")
    except Exception as e:
        st.error(f"❌ Could not display attachments: {e}")

# ============================================================
# 📊 DASHBOARD COMPONENT
# ============================================================
def show_dashboard(user, all_requests):
    role = user.get("role", "")
    dept = user.get("dept", "")
    full_name = user.get("full_name", user.get("username", "User"))
    if not role: return
    if role in ["Director", "Payroll", "Super Admin"]:
        visible = all_requests
    else:
        visible = [r for r in all_requests if r.get("dept", "") == dept]
    pending = [r for r in visible if r.get("status", "") == "pending"]
    approved = [r for r in visible if r.get("status", "") == "approved"]
    rejected = [r for r in visible if r.get("status", "") == "rejected"]
    total_approved_value = sum(r.get("amount", 0) for r in approved)
    st.subheader(f"👋 Welcome, {full_name}")
    st.markdown(f"**Role:** {role} | **Department:** {dept} | 📅 {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    st.divider()
    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("⏳ Pending", len(pending))
    with col2: st.metric("✅ Approved", len(approved))
    with col3: st.metric("❌ Rejected", len(rejected))
    with col4: st.metric("💰 Approved Total", f"£{total_approved_value:.2f}")
    st.divider()

# ============================================================
# SESSION STATE & LOGIN
# ============================================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_info" not in st.session_state:
    st.session_state.user_info = {}
if "editing_request_id" not in st.session_state:
    st.session_state.editing_request_id = None
if "editing_inspector_bonus_id" not in st.session_state:
    st.session_state.editing_inspector_bonus_id = None
if "editing_work_order_id" not in st.session_state:
    st.session_state.editing_work_order_id = None
if "editing_work_order_id_emp" not in st.session_state:
    st.session_state.editing_work_order_id_emp = None

def display_company_header():
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        if os.path.exists(LOGO_PATH): st.image(LOGO_PATH, width=300)
        else: st.title("⚡ ACOOLE ELECTRICAL LTD")
        st.caption("Acoole Operations & Authorisation Portal")
        st.divider()

def change_my_password_form():
    if not st.session_state.get("logged_in") or not st.session_state.get("user_info"):
        return
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
                    st.error("❌ Current password is NOT correct!"); return
                if new_pass1 != new_pass2:
                    st.error("❌ New passwords do NOT match!"); return
                if len(new_pass1) < 4:
                    st.error("❌ New password must be at least 4 characters!"); return
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
                elif new_cat.strip() in current_cats:
                    st.warning("⚠️ Category already exists!")
        st.divider()
        for i, cat in enumerate(current_cats):
            c1, c2, c3 = st.columns([4, 1, 1])
            with c1: st.markdown(f"• **{cat}**")
            with c2:
                if st.button(f"✏️ Edit", key=f"edit_cat_{i}"):
                    st.session_state[f"editing_cat_{i}"] = True
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
                        if st.form_submit_button("❌ Cancel"):
                            st.session_state[f"editing_cat_{i}"] = False; st.rerun()
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
                elif new_dept.strip() in current_depts:
                    st.warning("⚠️ Department already exists!")
        st.divider()
        for i, dept_name in enumerate(current_depts):
            c1, c2, c3 = st.columns([4, 1, 1])
            with c1: st.markdown(f"• **{dept_name}**")
            with c2:
                if st.button(f"✏️ Edit", key=f"edit_dept_{i}"):
                    st.session_state[f"editing_dept_{i}"] = True
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
                        if st.form_submit_button("❌ Cancel"):
                            st.session_state[f"editing_dept_{i}"] = False; st.rerun()
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
                elif new_role.strip() in current_roles:
                    st.warning("⚠️ Role already exists!")
        st.divider()
        for i, role in enumerate(current_roles):
            c1, c2, c3 = st.columns([4, 1, 1])
            with c1: st.markdown(f"• **{role}**")
            with c2:
                if st.button(f"✏️ Edit", key=f"edit_role_{i}"):
                    st.session_state[f"editing_role_{i}"] = True
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
                        if st.form_submit_button("❌ Cancel"):
                            st.session_state[f"editing_role_{i}"] = False; st.rerun()

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
            perm_inspector = st.checkbox(PERMISSION_LABELS["can_access_inspector_bonus"], value=defaults.get("can_access_inspector_bonus", False))
            new_dept = st.selectbox("🏢 Department", load_departments())
            if st.form_submit_button("✅ Create User Account", type="primary"):
                if not new_full_name.strip() or not new_username or not new_password:
                    st.error("❌ All fields required!")
                elif new_username in USERS:
                    st.error(f"❌ Username '{new_username}' already exists!")
                else:
                    USERS[new_username] = {
                        "full_name": new_full_name.strip(), "password": new_password,
                        "role": new_role, "dept": new_dept,
                        "can_view_all_dept": perm_view_all, "can_generate_pdf": perm_pdf,
                        "can_download_data": perm_download, "can_approve_requests": perm_approve,
                        "can_access_inspector_bonus": perm_inspector
                    }
                    save_users(USERS)
                    log_action("USER_CREATED", new_data={
                        "username": new_username, "full_name": new_full_name.strip(),
                        "role": new_role, "department": new_dept,
                        "permissions": {"can_view_all_dept": perm_view_all, "can_generate_pdf": perm_pdf,
                        "can_download_data": perm_download, "can_approve_requests": perm_approve,
                        "can_access_inspector_bonus": perm_inspector}
                    })
                    st.success(f"✅ User **'{new_full_name}'** created!"); st.balloons()
    with tab2:
        st.markdown("### ✏️ Edit User")
        edit_user_sel = st.selectbox("Select User to Edit", list(USERS.keys()), key="edit_user_selector")
        if edit_user_sel:
            curr = USERS[edit_user_sel]
            st.info(f"Current: **{curr.get('full_name', edit_user_sel)}** | {curr['role']} | {curr['dept']}")
            with st.form(f"edit_user_form_{edit_user_sel}", border=True, clear_on_submit=True):
                upd_full_name = st.text_input("👤 Full Name", value=curr.get("full_name", edit_user_sel))
                upd_username_new = st.text_input("🔐 Change Username", value=edit_user_sel).lower().strip()
                upd_password = st.text_input("🔑 New Password (leave blank to keep)", type="password")
                upd_role = st.selectbox("🎖️ Role", ROLES, index=ROLES.index(curr["role"]) if curr["role"] in ROLES else 0)
                dept_list = load_departments()
                upd_dept = st.selectbox("🏢 Department", dept_list, index=dept_list.index(curr["dept"]) if curr["dept"] in dept_list else 0)
                st.markdown("### ✅ Update Permissions")
                curr_perm_view = str(curr.get("can_view_all_dept", "False")).lower() == "true"
                curr_perm_pdf = str(curr.get("can_generate_pdf", "False")).lower() == "true"
                curr_perm_dl = str(curr.get("can_download_data", "False")).lower() == "true"
                curr_perm_app = str(curr.get("can_approve_requests", "False")).lower() == "true"
                curr_perm_ib = str(curr.get("can_access_inspector_bonus", "False")).lower() == "true"
                ecol1, ecol2 = st.columns(2)
                edit_view = ecol1.checkbox(PERMISSION_LABELS["can_view_all_dept"], value=curr_perm_view)
                edit_pdf = ecol1.checkbox(PERMISSION_LABELS["can_generate_pdf"], value=curr_perm_pdf)
                edit_dl = ecol2.checkbox(PERMISSION_LABELS["can_download_data"], value=curr_perm_dl)
                edit_app = ecol2.checkbox(PERMISSION_LABELS["can_approve_requests"], value=curr_perm_app)
                edit_ib = st.checkbox(PERMISSION_LABELS["can_access_inspector_bonus"], value=curr_perm_ib)
                if st.form_submit_button("🔄 Update User", type="primary"):
                    USERS = load_users()
                    if upd_username_new != edit_user_sel:
                        if upd_username_new in USERS:
                            st.error(f"❌ Username '{upd_username_new}' already exists!"); return
                        USERS[upd_username_new] = {
                            "full_name": upd_full_name.strip(),
                            "password": upd_password if upd_password else curr["password"],
                            "role": upd_role, "dept": upd_dept,
                            "can_view_all_dept": edit_view, "can_generate_pdf": edit_pdf,
                            "can_download_data": edit_dl, "can_approve_requests": edit_app,
                            "can_access_inspector_bonus": edit_ib
                        }
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
                    save_users(USERS)
                    log_action("USER_EDITED", old_data=curr, new_data={
                        "full_name": upd_full_name.strip(), "username": upd_username_new,
                        "role": upd_role, "department": upd_dept,
                        "permissions": {"can_view_all_dept": edit_view, "can_generate_pdf": edit_pdf,
                        "can_download_data": edit_dl, "can_approve_requests": edit_app,
                        "can_access_inspector_bonus": edit_ib}
                    })
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
# LOGIN PAGE
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
                st.session_state.logged_in = True
                st.session_state.user_info = {**USERS[username], "username": username}
                st.rerun()
            else:
                st.error("❌ Invalid Username or Password. Please try again.")
    st.stop()

# ============================================================
# ✅ REFRESH LEFT | LOGOUT RIGHT — SAME ROW ✅
# ============================================================
col_left, col_right = st.columns([4, 1])
with col_left:
    refresh_data_button()
with col_right:
    if st.button("🔒 Secure Logout", type="secondary", key="top_right_logout"):
        st.session_state.clear()
        st.rerun()

# ============================================================
# ✅ CENTERED LOGO — BELOW BUTTONS ✅
# ============================================================
display_company_header()

# ============================================================
# ✅ WELCOME BANNER — CORRECTLY POSITIONED ✅
# ============================================================
user_info = st.session_state.get("user_info", {})
full_name = user_info.get("full_name", user_info.get("username", "User"))
dept = user_info.get("dept", "")
role = user_info.get("role", "")

st.info(f"👤 Welcome: {full_name} | {dept} | {role}")

change_my_password_form()

# ============================================================
# ✅ LOAD DATA — AFTER WELCOME ✅
# ============================================================
all_live_requests = load_records_from_excel()
CATEGORIES = load_categories()

with st.sidebar:
    st.divider()
    SHOW_PDF_SECTION = False
    gen_range = False
    gen_all = False
    if SHOW_PDF_SECTION:
        st.subheader("📄 Generate Approved PDFs")
        st.caption("Filter by Date Range")
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            from_date = st.date_input("From Date", value=None, label_visibility="collapsed")
        with col_s2:
            to_date = st.date_input("To Date", value=None, label_visibility="collapsed")
        st.divider()
        gen_range = st.button("📥 Generate & Download (Range)", type="primary", key="sb_range")
        gen_all = st.button("📄 Generate ALL Approved PDFs", key="sb_all")
        if st.session_state.get("zip_range_data"):
            st.download_button("📦 Download Range PDFs (ZIP)", data=st.session_state.zip_range_data,
                file_name=st.session_state.zip_range_name, mime="application/zip",
                type="primary", use_container_width=True, key="dl_sb_range")
        if st.session_state.get("zip_all_data"):
            st.download_button("📦 Download ALL PDFs (ZIP)", data=st.session_state.zip_all_data,
                file_name=st.session_state.zip_all_name, mime="application/zip",
                type="primary", use_container_width=True, key="dl_sb_all")

if "zip_range_data" not in st.session_state:
    st.session_state.zip_range_data = None
if "zip_range_name" not in st.session_state:
    st.session_state.zip_range_name = None
if "zip_all_data" not in st.session_state:
    st.session_state.zip_all_data = None
if "zip_all_name" not in st.session_state:
    st.session_state.zip_all_name = None

def create_pdf_from_request(req):
    try:
        from fpdf import FPDF
        req_id = str(req.get("id", "unknown"))
        emp_name = str(req.get("emp_name", "Request")).replace(" ", "_")
        date_str = str(req.get("date", "unknown"))[:10].replace("-", "")
        filename = f"{emp_name}_{date_str}.pdf"
        filepath = os.path.join(PDF_DIR, filename)
        os.makedirs(PDF_DIR, exist_ok=True)
        pdf = FPDF()
        pdf.add_page()
        if os.path.exists(LOGO_PATH):
            pdf.image(LOGO_PATH, x=60, y=10, w=90)
        pdf.ln(35)
        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(0, 10, "Addition & Deduction Approval Form", ln=True, align="C")
        pdf.ln(5)
        pdf.set_draw_color(0, 0, 0)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(8)
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 8, "REQUEST DETAILS", ln=True)
        pdf.set_font("Helvetica", "", 11)
        def row(label, value):
            pdf.cell(55, 7, f"{label}:", 0, 0)
            pdf.cell(0, 7, str(value), 0, 1)
        row("Request ID", req_id)
        row("Employee Name", req.get("emp_name", ""))
        row("Department", req.get("dept", ""))
        row("Transaction Type", req.get("type", ""))
        row("Category / Reason", req.get("category", ""))
        row("Request Date", str(req.get("date", ""))[:10])
        amount = req.get("amount", "")
        try: amount = f"£{float(amount):,.2f}"
        except: pass
        row("Amount Approved", amount)
        row("Line Manager", req.get("manager", ""))
        submitted_by = get_submitted_by(req)
        if submitted_by:
            row("Submitted By", submitted_by)
        pdf.ln(5)
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 8, "DESCRIPTION / JUSTIFICATION", ln=True)
        pdf.set_font("Helvetica", "", 11)
        justification = str(req.get("justification", req.get("description", "")))
        pdf.multi_cell(0, 7, justification)
        pdf.ln(5)
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 8, "DIRECTOR APPROVAL", ln=True)
        pdf.set_font("Helvetica", "", 11)
        row("Decision", "APPROVED")
        row("Approved By", req.get("approved_by", "Andy Acoole"))
        row("Approval Date / Time", str(req.get("approved_date", ""))[:16])
        pdf.ln(10)
        pdf.set_draw_color(100, 100, 100)
        pdf.dashed_line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(8)
        pdf.cell(0, 7, "Authorised Signature / Director", ln=True, align="C")
        pdf.ln(15)
        pdf.set_font("Helvetica", "B", 22)
        pdf.set_text_color(0, 120, 0)
        pdf.cell(0, 15, "[ APPROVED ]", ln=True, align="C")
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 7, "Acoole Electrical Ltd", ln=True, align="C")
        pdf.set_text_color(0, 0, 0)
        pdf.output(filepath)
        return filepath
    except Exception as e:
        st.warning(f"⚠️ PDF Error for #{req.get('id', '?')}: {str(e)}")
        return None

if gen_range:
    st.success("✅ Generating PDFs for selected date range...")
    records = load_records_from_excel()
    st.info(f"📋 Total records loaded: {len(records)}")
    approved_records = [r for r in records if str(r.get("status", "")).strip().lower() == "approved"]
    st.info(f"✅ Approved records (before date filter): {len(approved_records)}")
    matched_records = []
    for req in approved_records:
        req_date = None
        date_str = str(req.get("date", "")).strip()[:10]
        for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%m-%d-%Y"):
            try:
                req_date = datetime.strptime(date_str, fmt).date()
                break
            except: continue
        include = True
        if 'from_date' in locals() and from_date and req_date and req_date < from_date: include = False
        if 'to_date' in locals() and to_date and req_date and req_date > to_date: include = False
        if include: matched_records.append(req)
    st.info(f"📅 Records matched by date range: {len(matched_records)}")
    if matched_records:
        import zipfile, io
        zip_buffer = io.BytesIO()
        pdf_count = 0
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for req in matched_records:
                pdf_path = create_pdf_from_request(req)
                if pdf_path and os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as f:
                        zipf.writestr(os.path.basename(pdf_path), f.read())
                        pdf_count += 1
        zip_buffer.seek(0)
        if pdf_count > 0:
            st.success(f"✅ Generated {pdf_count} PDFs! ⬅️ Go to sidebar to download")
            st.session_state.zip_range_data = zip_buffer.getvalue()
            st.session_state.zip_range_name = f"Approved_PDFs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        else:
            st.error("❌ PDF generation failed.")
            st.session_state.zip_range_data = None
    else:
        st.info("ℹ️ No records matched.")
        st.session_state.zip_range_data = None

if gen_all:
    st.success("✅ Generating ALL approved PDFs...")
    records = load_records_from_excel()
    approved_records = [r for r in records if str(r.get("status", "")).strip().lower() == "approved"]
    if approved_records:
        import zipfile, io
        zip_buffer = io.BytesIO()
        pdf_count = 0
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for req in approved_records:
                pdf_path = create_pdf_from_request(req)
                if pdf_path and os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as f:
                        zipf.writestr(os.path.basename(pdf_path), f.read())
                        pdf_count += 1
        zip_buffer.seek(0)
        if pdf_count > 0:
            st.success(f"✅ Generated {pdf_count} PDFs! ⬅️ Go to sidebar to download")
            st.session_state.zip_all_data = zip_buffer.getvalue()
            st.session_state.zip_all_name = f"All_Approved_PDFs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        else:
            st.error("❌ PDF generation failed.")
            st.session_state.zip_all_data = None
    else:
        st.info("ℹ️ No approved requests found.")
        st.session_state.zip_all_data = None

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

    tab_add_ded, tab_work_orders, tab_inspector_bonus = st.tabs([
        "➕ Addition & Deduction",
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
                        old_data_raw = req.get("old_data", "")
                        if old_data_raw and old_data_raw not in ["", "{}", "None"]:
                            st.divider(); st.markdown("### 🔄 What Changed / Edits"); show_old_new_comparison(old_data_raw, req)
                        else:
                            st.divider(); st.success("✅ **New Request — No previous version**")
                        st.divider(); display_pdf_button(req, can_generate=True)
        with tab_approved:
            approved = [r for r in all_live_requests if r.get("status") == "approved"]
            if not approved: st.info("📋 No approved requests.")
            else:
                st.metric("✅ Approved", len(approved))
                approved_search = st.text_input("🔎 Search approved requests", placeholder="Search by ID, employee, department, submitted by, approved by, amount, date or comments...", key="payroll_approved_search")
                if approved_search.strip():
                    q = approved_search.strip().lower()
                    approved = [r for r in approved if q in " ".join([
                        str(r.get("id", "")), str(r.get("emp_name", "")), str(r.get("dept", "")),
                        str(r.get("decision_by", "")), str(r.get("approved_by", "")),
                        str(r.get("amount", "")), str(r.get("decision_date", "")),
                        str(r.get("date", "")), str(r.get("director_comments", "")),
                    ]).lower()]
                    st.caption(f"🔎 Showing {len(approved)} matching approved request(s).")
                st.divider()
                if not approved: st.warning("No approved requests match your search.")
                for req in reversed(approved):
                    dec_by = req.get('decision_by', 'Director')
                    dec_date = req.get('decision_date', '')
                    if dec_date and " " in dec_date:
                        date_part, time_part = dec_date.split(" ", 1)
                        display_date = f"{date_part} ⏰ {time_part}"
                    else: display_date = dec_date if dec_date else ""
                    extra_text = f" | ✅ Approved by {dec_by} on {display_date}" if display_date else f" | ✅ Approved by {dec_by}"
                    title = f"🟢 ID #{req.get('id')} | {req.get('emp_name')} | £{float(req.get('amount',0)):.2f} | {req.get('dept','')}{extra_text}"
                    with st.expander(title):
                        st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept', '')}")
                        submitted_by = get_submitted_by(req)
                        if submitted_by: st.write(f"📝 **Submitted by:** {submitted_by}")
                        st.write(f"💷 Amount: £{float(req.get('amount',0)):.2f}")
                        st.write(f"🎯 Approved By: {dec_by}")
                        if display_date: st.write(f"📅 Approval Date: {display_date}")
                        st.info(f"💬 Comments: {req.get('director_comments', 'None')}")
                        display_attachments(req)
                        st.divider(); display_pdf_button(req, can_generate=True)
        with tab_rejected:
            rejected = [r for r in all_live_requests if r.get("status") == "rejected"]
            if not rejected: st.success("✅ No rejected requests!")
            else:
                st.metric("❌ Rejected", len(rejected)); st.divider()
                for req in reversed(rejected):
                    with st.expander(f"🔴 ID #{req.get('id')} | {req.get('emp_name')} | £{float(req.get('amount',0)):.2f} | {req.get('dept')}"):
                        st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept')}")
                        submitted_by = get_submitted_by(req)
                        if submitted_by: st.write(f"📝 **Submitted by:** {submitted_by}")
                        st.write(f"💷 Amount: £{float(req.get('amount',0)):.2f}")
                        st.error(f"❌ Rejected By: {req.get('decision_by', '—')} on {format_date(req.get('decision_date', ''))}")
                        st.error(f"💬 Reason: {req.get('director_comments', 'None')}")
                        display_attachments(req)
                        st.divider(); display_pdf_button(req, can_generate=True)

    with tab_work_orders:
        render_work_order_payroll_portal(full_name)

    with tab_inspector_bonus:
        render_inspector_bonus_payroll_portal(full_name)

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
                            with open(file_path, "rb") as f:
                                col_dl.download_button("⬇️", f.read(), file_name=fname, key=f"dl_{safe_key}")
                        else: col_dl.caption("⚠️ Missing")
                        if keep: files_to_keep.append(fname)
                        else: files_to_remove.append(fname)
                    if files_to_remove: st.warning(f"🗑️ Will remove: {', '.join(files_to_remove)}")
                else: st.info("📋 No attachments currently attached.")
                st.markdown("#### ➕ Attach New Files")
                new_files_upload = st.file_uploader("Upload additional files", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key=f"new_upload_{eid}")
                st.info(f"✅ Result: **{len(files_to_keep)} kept** + **{len(new_files_upload or [])} new** = {len(files_to_keep)+len(new_files_upload or [])} total files")
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
                                with open(edit_path, "wb") as outfile:
                                    outfile.write(f.getbuffer())
                                _upload_to_drive_bg(edit_path, fn)
                                final_attachments.append(fn)
                        records = load_records_from_excel()
                        old_data_dict = {"emp_name": rec.get("emp_name"), "dept": rec.get("dept"), "type": rec.get("type"), "category": rec.get("category"), "date": rec.get("date"), "amount": rec.get("amount"), "manager": rec.get("manager"), "desc": rec.get("desc")}
                        new_data_dict = {"emp_name": en.strip(), "dept": dept_name, "type": rt, "category": ct, "date": str(dt_val), "amount": amt, "manager": mgr.strip(), "desc": desc.strip()}
                        for r in records:
                            if int(r.get("id", 0)) == int(eid):
                                r["emp_name"] = en.strip(); r["type"] = rt; r["category"] = ct; r["amount"] = amt
                                r["date"] = str(dt_val); r["manager"] = mgr.strip(); r["desc"] = desc.strip()
                                r["status"] = "pending"; r["attachment_name"] = ", ".join(final_attachments) or "None"
                                r["old_data"] = json.dumps(old_data_dict); break
                        log_action("EDITED", eid, old_data=old_data_dict, new_data=new_data_dict)
                        save_all_records(records)
                        st.success(f"✅ Updated! Removed {len(files_to_remove)} | Kept {len(files_to_keep)} | Added {len(new_files_upload or [])}")
                        st.session_state.editing_request_id = None; st.rerun()
                if st.button("❌ Cancel", key=f"cancel_edit_{eid}"):
                    st.session_state.editing_request_id = None; st.rerun()
        else:
            st.subheader(f"➕ New Request — {dept_name}")
            nid = get_next_id(all_live_requests)
            st.markdown(f"**🆔 Request ID:** `#{nid}`")
            with st.form("new_req", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1:
                    en = st.text_input("👤 Employee Name")
                    rt = st.selectbox("🔄 Transaction Type", ["Addition", "Deduction"])
                    ct = st.selectbox("🏷️ Category / Reason", CATEGORIES)
                    amt = st.number_input("💷 Amount (£)", 0.01, step=10.0)
                with c2:
                    from datetime import datetime as dt
                    dt_val = st.date_input("📅 Date", value=dt.today())
                    mgr = st.text_input("👔 Line Manager")
                    files = st.file_uploader("📎 Attachments", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)
                    desc = st.text_area("📝 Description / Justification")
                if st.form_submit_button("📤 Send to Director", type="primary"):
                    if en.strip() and mgr.strip() and desc.strip():
                        att_list = []
                        if files:
                            for i, f in enumerate(files, 1):
                                fn = f"ID_{nid}_F{i}_{f.name}"
                                file_path = os.path.join(UPLOAD_DIR, fn)
                                with open(file_path, "wb") as out:
                                    out.write(f.getbuffer())
                                att_list.append(fn)
                                _upload_to_drive_bg(file_path, fn)
                        payload = {"id": nid, "emp_name": en.strip(), "dept": dept_name, "type": rt, "category": ct, "date": str(dt_val), "amount": amt, "manager": mgr.strip(), "desc": desc.strip(), "attachment_name": ", ".join(att_list) or "None", "status": "pending", "director_comments": "", "decision_date": "", "decision_by": "", "submitted_by": full_name, "pdf_path": "", "edited_from_id": "", "old_data": ""}
                        save_record_to_excel(payload)
                        log_action("CREATED", nid)
                        st.success(f"✅ Request #{nid} sent for approval!"); st.rerun()
                    else: st.error("⚠️ Please fill in: Employee Name, Line Manager, and Description")
            st.divider()
            st.subheader("📋 My Department Requests")
            my_reqs = [r for r in all_live_requests if r.get("dept") == dept_name]
            my_dept_search = st.text_input("🔎 Search my department requests", placeholder="Search by ID, employee, status, amount, manager, category, date, description, approver or comments...", key="my_department_requests_search")
            if my_dept_search.strip():
                q = my_dept_search.strip().lower()
                my_reqs = [r for r in my_reqs if q in " ".join([str(r.get("id", "")), str(r.get("emp_name", "")), str(r.get("dept", "")), str(r.get("status", "")), str(r.get("amount", "")), str(r.get("manager", "")), str(r.get("type", "")), str(r.get("category", "")), str(r.get("date", "")), str(r.get("desc", "")), str(r.get("decision_by", "")), str(r.get("approved_by", "")), str(r.get("submitted_by", "")), str(r.get("decision_date", "")), str(r.get("director_comments", "")), str(r.get("attachment_name", ""))]).lower()]
                st.caption(f"🔎 Showing {len(my_reqs)} matching department request(s).")
            if not my_reqs:
                if my_dept_search.strip(): st.warning("No department requests match your search.")
                else: st.info("📋 No requests yet.")
            else:
                for req in reversed(my_reqs):
                    status = req.get("status", "pending").lower()
                    icon = "🟡" if status == "pending" else ("🟢" if status == "approved" else "🔴")
                    dec_by = req.get("decision_by", ""); dec_date = format_date(req.get("decision_date", ""))
                    if status in ["approved", "rejected"] and dec_by:
                        title = f"{icon} ID #{req.get('id')} | {req.get('emp_name')} | {status.upper()} | £{float(req.get('amount',0)):.2f} | ✅ {dec_by} — {dec_date.replace(' ', ' 🕓 ')}"
                    else:
                        title = f"{icon} ID #{req.get('id')} | {req.get('emp_name')} | {status.upper()} | £{float(req.get('amount',0)):.2f} | 📅 {format_date(req.get('date', ''))}"
                    with st.expander(title):
                        st.write(f"👤 Employee: {req.get('emp_name')} | 👔 Manager: {req.get('manager')}")
                        submitted_by = get_submitted_by(req)
                        if submitted_by: st.write(f"📝 **Submitted by:** {submitted_by}")
                        st.write(f"🔄 Type: {req.get('type')} | 🏷️ Category: {req.get('category')}")
                        st.info(f"📝 Description: {req.get('desc')}")
                        display_attachments(req)
                        if req.get("director_comments"): st.info(f"💬 Director Comments: {req.get('director_comments')}")
                        if status == "approved": display_pdf_button(req, can_generate=True)
                        if status in ["pending", "rejected"]:
                            if st.button(f"✏️ Edit Request #{req.get('id')}", key=f"edit_{req.get('id')}"):
                                st.session_state.editing_request_id = req.get("id"); st.rerun()
    with work_order_tab:
        st.subheader("🛠️ Work Orders — Manager")
        st.info("Only users assigned the Work Order Manager role can access work orders.")
        render_work_order_manager_portal(full_name, dept_name, show_total=True)

elif role in ["Manager", "Staff", "Team Member"]:
    dept_name = dept

    has_inspector_bonus = user_info.get("can_access_inspector_bonus", False)

    if has_inspector_bonus:
        tab_main, tab_bonus = st.tabs(["➕ Addition & Deduction", "💰 National Grid Inspector Bonus"])
    else:
        tab_main = st.container()
        tab_bonus = None

    with tab_main:
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
                            with open(file_path, "rb") as f:
                                col_dl.download_button("⬇️", f.read(), file_name=fname, key=f"dl_{safe_key}")
                        else: col_dl.caption("⚠️ Missing")
                        if keep: files_to_keep.append(fname)
                        else: files_to_remove.append(fname)
                    if files_to_remove: st.warning(f"🗑️ Will remove: {', '.join(files_to_remove)}")
                else: st.info("📋 No attachments currently attached.")
                st.markdown("#### ➕ Attach New Files")
                new_files_upload = st.file_uploader("Upload additional files", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key=f"new_upload_{eid}")
                st.info(f"✅ Result: **{len(files_to_keep)} kept** + **{len(new_files_upload or [])} new** = {len(files_to_keep)+len(new_files_upload or [])} total files")
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
                                with open(edit_path, "wb") as outfile:
                                    outfile.write(f.getbuffer())
                                _upload_to_drive_bg(edit_path, fn)
                                final_attachments.append(fn)
                        records = load_records_from_excel()
                        old_data_dict = {"emp_name": rec.get("emp_name"), "dept": rec.get("dept"), "type": rec.get("type"), "category": rec.get("category"), "date": rec.get("date"), "amount": rec.get("amount"), "manager": rec.get("manager"), "desc": rec.get("desc")}
                        new_data_dict = {"emp_name": en.strip(), "dept": dept_name, "type": rt, "category": ct, "date": str(dt_val), "amount": amt, "manager": mgr.strip(), "desc": desc.strip()}
                        for r in records:
                            if int(r.get("id", 0)) == int(eid):
                                r["emp_name"] = en.strip(); r["type"] = rt; r["category"] = ct; r["amount"] = amt
                                r["date"] = str(dt_val); r["manager"] = mgr.strip(); r["desc"] = desc.strip()
                                r["status"] = "pending"; r["attachment_name"] = ", ".join(final_attachments) or "None"
                                r["old_data"] = json.dumps(old_data_dict); break
                        log_action("EDITED", eid, old_data=old_data_dict, new_data=new_data_dict)
                        save_all_records(records)
                        st.success(f"✅ Updated! Removed {len(files_to_remove)} | Kept {len(files_to_keep)} | Added {len(new_files_upload or [])}")
                        st.session_state.editing_request_id = None; st.rerun()
                if st.button("❌ Cancel", key=f"cancel_edit_{eid}"):
                    st.session_state.editing_request_id = None; st.rerun()
        else:
            st.subheader(f"➕ New Request — {dept_name}")
            nid = get_next_id(all_live_requests)
            st.markdown(f"**🆔 Request ID:** `#{nid}`")
            with st.form("new_req", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1:
                    en = st.text_input("👤 Employee Name")
                    rt = st.selectbox("🔄 Transaction Type", ["Addition", "Deduction"])
                    ct = st.selectbox("🏷️ Category / Reason", CATEGORIES)
                    amt = st.number_input("💷 Amount (£)", 0.01, step=10.0)
                with c2:
                    from datetime import datetime as dt
                    dt_val = st.date_input("📅 Date", value=dt.today())
                    mgr = st.text_input("👔 Line Manager")
                    files = st.file_uploader("📎 Attachments", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)
                    desc = st.text_area("📝 Description / Justification")
                if st.form_submit_button("📤 Send to Director", type="primary"):
                    if en.strip() and mgr.strip() and desc.strip():
                        att_list = []
                        if files:
                            for i, f in enumerate(files, 1):
                                fn = f"ID_{nid}_F{i}_{f.name}"
                                file_path = os.path.join(UPLOAD_DIR, fn)
                                with open(file_path, "wb") as out:
                                    out.write(f.getbuffer())
                                att_list.append(fn)
                                _upload_to_drive_bg(file_path, fn)
                        payload = {"id": nid, "emp_name": en.strip(), "dept": dept_name, "type": rt, "category": ct, "date": str(dt_val), "amount": amt, "manager": mgr.strip(), "desc": desc.strip(), "attachment_name": ", ".join(att_list) or "None", "status": "pending", "director_comments": "", "decision_date": "", "decision_by": "", "submitted_by": full_name, "pdf_path": "", "edited_from_id": "", "old_data": ""}
                        save_record_to_excel(payload)
                        log_action("CREATED", nid)
                        st.success(f"✅ Request #{nid} sent for approval!"); st.rerun()
                    else: st.error("⚠️ Please fill in: Employee Name, Line Manager, and Description")
            st.divider()
            st.subheader("📋 My Department Requests")
            my_reqs = [r for r in all_live_requests if r.get("dept") == dept_name]
            my_dept_search = st.text_input("🔎 Search my department requests", placeholder="Search by ID, employee, status, amount, manager, category, date, description, approver or comments...", key="my_department_requests_search")
            if my_dept_search.strip():
                q = my_dept_search.strip().lower()
                my_reqs = [r for r in my_reqs if q in " ".join([str(r.get("id", "")), str(r.get("emp_name", "")), str(r.get("dept", "")), str(r.get("status", "")), str(r.get("amount", "")), str(r.get("manager", "")), str(r.get("type", "")), str(r.get("category", "")), str(r.get("date", "")), str(r.get("desc", "")), str(r.get("decision_by", "")), str(r.get("approved_by", "")), str(r.get("submitted_by", "")), str(r.get("decision_date", "")), str(r.get("director_comments", "")), str(r.get("attachment_name", ""))]).lower()]
                st.caption(f"🔎 Showing {len(my_reqs)} matching department request(s).")
            if not my_reqs:
                if my_dept_search.strip(): st.warning("No department requests match your search.")
                else: st.info("📋 No requests yet.")
            else:
                for req in reversed(my_reqs):
                    status = req.get("status", "pending").lower()
                    icon = "🟡" if status == "pending" else ("🟢" if status == "approved" else "🔴")
                    dec_by = req.get("decision_by", ""); dec_date = format_date(req.get("decision_date", ""))
                    if status in ["approved", "rejected"] and dec_by:
                        title = f"{icon} ID #{req.get('id')} | {req.get('emp_name')} | {status.upper()} | £{float(req.get('amount',0)):.2f} | ✅ {dec_by} — {dec_date.replace(' ', ' 🕓 ')}"
                    else:
                        title = f"{icon} ID #{req.get('id')} | {req.get('emp_name')} | {status.upper()} | £{float(req.get('amount',0)):.2f} | 📅 {format_date(req.get('date', ''))}"
                    with st.expander(title):
                        st.write(f"👤 Employee: {req.get('emp_name')} | 👔 Manager: {req.get('manager')}")
                        submitted_by = get_submitted_by(req)
                        if submitted_by: st.write(f"📝 **Submitted by:** {submitted_by}")
                        st.write(f"🔄 Type: {req.get('type')} | 🏷️ Category: {req.get('category')}")
                        st.info(f"📝 Description: {req.get('desc')}")
                        display_attachments(req)
                        if req.get("director_comments"): st.info(f"💬 Director Comments: {req.get('director_comments')}")
                        if status == "approved": display_pdf_button(req, can_generate=True)
                        if status in ["pending", "rejected"]:
                            if st.button(f"✏️ Edit Request #{req.get('id')}", key=f"edit_{req.get('id')}"):
                                st.session_state.editing_request_id = req.get("id"); st.rerun()

    if tab_bonus:
        with tab_bonus:
            render_inspector_bonus_portal(full_name, dept_name)

elif role == "Director":
    director_addition_tab, director_work_order_tab, director_inspector_tab = st.tabs([
        "➕ Addition & Deduction",
        "🛠️ Work Orders",
        "💰 National Grid Inspector Bonus"
    ])
    with director_addition_tab:
        st.subheader("🎛️ Director Approval Portal — Andy Acoole")
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
                            st.write(f"📅 **Date:** {format_date(req.get('date',''))}")
                            st.info(f"📝 **Description / Justification:**\n{req.get('desc','')}")
                            display_attachments(req)
                            old_data_raw = req.get("old_data", "")
                            if old_data_raw and old_data_raw not in ["", "{}", "None"]:
                                st.divider(); st.markdown("### 🔄 What Changed / Edits"); show_old_new_comparison(old_data_raw, req)
                            else:
                                st.divider(); st.success("✅ **New Request — No previous version**")
                        with col_right:
                            st.markdown("### ✍️ Decision")
                            comments = st.text_area("Director Comments", key=f"comm_{req_id}")
                            approve_btn = st.button("✅ APPROVE", type="primary", key=f"appr_{req_id}")
                            reject_btn = st.button("❌ REJECT", type="secondary", key=f"rejt_{req_id}")
                            if approve_btn:
                                records = load_records_from_excel()
                                for r in records:
                                    if int(r.get("id",0)) == int(req_id):
                                        r["status"] = "approved"; r["approved_by"] = full_name; r["decision_by"] = full_name
                                        r["director_comments"] = comments; r["decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                        r["approved_date"] = r["decision_date"]; break
                                save_all_records(records); log_action("APPROVED", req_id)
                                st.success(f"✅ Request #{req_id} APPROVED."); st.rerun()
                            if reject_btn:
                                records = load_records_from_excel()
                                for r in records:
                                    if int(r.get("id",0)) == int(req_id):
                                        r["status"] = "rejected"; r["decision_by"] = full_name
                                        r["director_comments"] = comments; r["decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S"); break
                                save_all_records(records); log_action("REJECTED", req_id)
                                st.error(f"❌ Request #{req_id} REJECTED."); st.rerun()
        with tab_approved:
            approved = [r for r in all_live_requests if r.get("status") == "approved"]
            if not approved: st.info("📋 No approved requests yet.")
            else:
                st.metric("✅ Approved Requests", len(approved))
                approved_search = st.text_input("🔎 Search approved requests", placeholder="Search by ID, employee, department, submitted by, approved by, amount, date or comments...", key="director_approved_search")
                if approved_search.strip():
                    q = approved_search.strip().lower()
                    approved = [r for r in approved if q in " ".join([str(r.get("id", "")), str(r.get("emp_name", "")), str(r.get("dept", "")), str(r.get("decision_by", "")), str(r.get("approved_by", "")), str(r.get("submitted_by", "")), str(r.get("amount", "")), str(r.get("decision_date", "")), str(r.get("date", "")), str(r.get("director_comments", ""))]).lower()]
                    st.caption(f"🔎 Showing {len(approved)} matching approved request(s).")
                st.info("🔄 **Change Status:** Move back to Pending ❘ Change to Rejected")
                st.divider()
                if not approved: st.warning("No approved requests match your search.")
                for req in reversed(approved):
                    req_id = req.get("id")
                    dec_by = req.get('decision_by', 'Director'); dec_date = format_date(req.get('decision_date',''))
                    with st.expander(f"🟢 ID #{req_id} | {req.get('emp_name')} | £{float(req.get('amount',0)):.2f} | {req.get('dept','')} | ✅ {dec_by} — {dec_date[:10]} ⏰{dec_date[11:]}"):
                        st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept','')}")
                        submitted_by = get_submitted_by(req)
                        if submitted_by: st.write(f"📝 **Submitted by:** {submitted_by}")
                        st.write(f"💷 Amount: £{float(req.get('amount',0)):.2f}")
                        st.write(f"🎯 Approved By: {dec_by}")
                        st.write(f"📅 Approval Date: {dec_date}")
                        st.info(f"💬 Comments: {req.get('director_comments', 'None')}")
                        display_attachments(req)
                        st.divider(); display_pdf_button(req, can_generate=True)
                        st.markdown("### 🔄 Change Status")
                        new_comments = st.text_area("Add comment (optional)", key=f"chg_comm_app_{req_id}", placeholder="Reason for status change...")
                        col1, col2 = st.columns(2)
                        with col1:
                            if st.button("⏳ Move to Pending", key=f"app_to_pend_{req_id}"):
                                records = load_records_from_excel()
                                for r in records:
                                    if int(r.get("id",0)) == int(req_id):
                                        r["status"] = "pending"
                                        r["director_comments"] = (r.get("director_comments","") + f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M')}] ⏳ Changed to Pending by {full_name}: {new_comments}").strip()
                                        r["decision_date"] = ""; r["decision_by"] = ""; break
                                save_all_records(records)
                                log_action("STATUS_CHANGED", req_id, old_data={"status":"approved"}, new_data={"status":"pending"})
                                st.success(f"✅ Request #{req_id} moved to Pending. Audit Log updated."); st.rerun()
                        with col2:
                            if st.button("❌ Change to REJECTED", key=f"app_to_rej_{req_id}"):
                                records = load_records_from_excel()
                                for r in records:
                                    if int(r.get("id",0)) == int(req_id):
                                        r["status"] = "rejected"
                                        r["director_comments"] = (r.get("director_comments","") + f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M')}] ❌ Changed to Rejected by {full_name}: {new_comments}").strip()
                                        r["decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S"); r["decision_by"] = full_name; break
                                save_all_records(records)
                                log_action("STATUS_CHANGED", req_id, old_data={"status":"approved"}, new_data={"status":"rejected"})
                                st.success(f"✅ Request #{req_id} changed to Rejected. Audit Log updated."); st.rerun()
        with tab_rejected:
            rejected = [r for r in all_live_requests if r.get("status") == "rejected"]
            if not rejected: st.success("✅ No rejected requests!")
            else:
                st.metric("❌ Rejected Requests", len(rejected))
                st.info("🔄 **Change Status:** Move back to Pending ❘ Change to Approved")
                st.divider()
                for req in reversed(rejected):
                    req_id = req.get("id")
                    dec_by = req.get('decision_by', 'Director'); dec_date = format_date(req.get('decision_date',''))
                    with st.expander(f"🔴 ID #{req_id} | {req.get('emp_name')} | £{float(req.get('amount',0)):.2f} | {req.get('dept','')} | ❌ {dec_by} — {dec_date[:10]} ⏰{dec_date[11:]}"):
                        st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept')}")
                        submitted_by = get_submitted_by(req)
                        if submitted_by: st.write(f"📝 **Submitted by:** {submitted_by}")
                        st.write(f"💷 Amount: £{float(req.get('amount',0)):.2f}")
                        st.error(f"❌ Rejected By: {dec_by} on {dec_date}")
                        st.error(f"💬 Reason: {req.get('director_comments', 'None')}")
                        display_attachments(req)
                        st.markdown("### 🔄 Change Status")
                        new_comments = st.text_area("Add comment (optional)", key=f"chg_comm_rej_{req_id}", placeholder="Reason for status change...")
                        col1, col2 = st.columns(2)
                        with col1:
                            if st.button("⏳ Move to Pending", key=f"rej_to_pend_{req_id}"):
                                records = load_records_from_excel()
                                for r in records:
                                    if int(r.get("id",0)) == int(req_id):
                                        r["status"] = "pending"
                                        r["director_comments"] = (r.get("director_comments","") + f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M')}] ⏳ Changed to Pending by {full_name}: {new_comments}").strip()
                                        r["decision_date"] = ""; r["decision_by"] = ""; break
                                save_all_records(records)
                                log_action("STATUS_CHANGED", req_id, old_data={"status":"rejected"}, new_data={"status":"pending"})
                                st.success(f"✅ Request #{req_id} moved to Pending. Audit Log updated."); st.rerun()
                        with col2:
                            if st.button("✅ Change to APPROVED", key=f"rej_to_app_{req_id}"):
                                records = load_records_from_excel()
                                for r in records:
                                    if int(r.get("id",0)) == int(req_id):
                                        r["status"] = "approved"
                                        r["director_comments"] = (r.get("director_comments","") + f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M')}] ✅ Changed to Approved by {full_name}: {new_comments}").strip()
                                        r["decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S"); r["decision_by"] = full_name; break
                                save_all_records(records)
                                log_action("STATUS_CHANGED", req_id, old_data={"status":"rejected"}, new_data={"status":"approved"})
                                st.success(f"✅ Request #{req_id} changed to Approved. Audit Log updated."); st.rerun()
    with director_work_order_tab:
        render_work_order_director_portal(full_name)
    with director_inspector_tab:
        render_inspector_bonus_director_portal(full_name)

elif role == "Super Admin":
    st.subheader("🛡️ Super Admin — All Requests")
    st.info("✅ View ALL requests across ALL departments. Download PDFs. **Approval → Director only.**")
    st.divider()
    tab_pending, tab_approved, tab_rejected, tab_manage = st.tabs(["⏳ All Pending", "✅ All Approved", "❌ All Rejected", "🔧 System Management"])
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
                    st.write(f"👔 Line Manager: {req.get('manager')} | 📅 Date: {format_date(req.get('date',''))}")
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
                dec_by = req.get('decision_by', 'Director'); dec_date = req.get('decision_date', '')
                display_date = dec_date[:10] if dec_date and len(dec_date) >= 10 else ""
                extra_text = f" | ✅ Approved by {dec_by} on {display_date}" if display_date else f" | ✅ Approved by {dec_by}"
                title = f"🟢 ID #{req_id} | {req.get('emp_name')} | {req.get('dept')} | £{amount:.2f}{extra_text}"
                with st.expander(title):
                    st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept')}")
                    st.write(f"💷 Amount: £{amount:.2f}")
                    submitted_by = get_submitted_by(req)
                    if submitted_by: st.write(f"📝 **Submitted by:** {submitted_by}")
                    st.write(f"🎯 **Approved By:** {dec_by}")
                    if display_date: st.write(f"📅 **Approval Date:** {display_date}")
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
                dec_by = req.get('decision_by', 'Director'); dec_date = format_date(req.get('decision_date',''))
                title = f"🔴 ID #{req_id} | {req.get('emp_name')} | {req.get('dept')} | £{amount:.2f}"
                with st.expander(title):
                    st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept')}")
                    st.write(f"💷 Amount: £{amount:.2f}")
                    submitted_by = get_submitted_by(req)
                    if submitted_by: st.write(f"📝 **Submitted by:** {submitted_by}")
                    st.error(f"❌ Rejected By: {dec_by} on {dec_date}")
                    st.error(f"💬 Reason: {req.get('director_comments', 'None')}")
                    display_attachments(req)
                    st.divider(); display_pdf_button(req, can_generate=True)
    with tab_manage:
        tab_settings, tab_users, tab_audit = st.tabs(["⚙️ System Settings", "👤 User Management", "📖 Audit History"])
        with tab_settings: settings_management_panel()
        with tab_users: user_management_panel()
        with tab_audit:
            if "display_audit_log_panel" in globals(): display_audit_log_panel()
            else: st.info("📖 Audit log panel not defined — skipping")
            st.divider()
            st.subheader("⚠️ Super Admin — Data Reset / Live Launch")
            st.warning("These controls are permanent. They are intended for preparing the portal for live use. Clearing requests removes all request records from requests.xlsx. Clearing the audit history removes all audit entries. User accounts, system settings and uploaded attachment files are NOT deleted by these controls.")
            danger_col1, danger_col2 = st.columns(2)
            with danger_col1:
                if not st.session_state.get("confirm_clear_requests", False):
                    if st.button("🧹 Clear All Submitted Requests", type="secondary", key="clear_all_requests_btn"):
                        st.session_state["confirm_clear_requests"] = True; st.rerun()
                else:
                    st.error("⚠️ This will permanently remove ALL request records from the application.")
                    c1, c2 = st.columns(2)
                    with c1:
                        if st.button("✅ Yes, Clear Requests", type="primary", key="confirm_clear_all_requests_btn"):
                            clear_all_requests_file(); st.session_state["confirm_clear_requests"] = False
                            st.success("✅ All submitted request records have been cleared."); st.rerun()
                    with c2:
                        if st.button("↩️ Cancel", key="cancel_clear_all_requests_btn"):
                            st.session_state["confirm_clear_requests"] = False; st.rerun()
            with danger_col2:
                reset_col1, reset_col2, reset_col3 = st.columns(3)
                with reset_col1:
                    if not st.session_state.get("confirm_clear_inspector_bonus", False):
                        if st.button("💰 Clear All Inspector Bonuses", key="super_admin_clear_all_inspector_bonus", type="secondary", use_container_width=True):
                            st.session_state["confirm_clear_inspector_bonus"] = True
                    else:
                        st.warning("⚠️ This permanently removes ALL Inspector Bonus records.")
                        cc1, cc2 = st.columns(2)
                        with cc1:
                            if st.button("✅ Confirm Clear Bonuses", key="super_admin_confirm_clear_inspector_bonus", use_container_width=True):
                                clear_all_inspector_bonus()
                                log_action("SUPER_ADMIN_CLEAR_INSPECTOR_BONUS", "ALL", decision_by=full_name)
                                st.session_state["confirm_clear_inspector_bonus"] = False
                                st.success("✅ All Inspector Bonus records have been cleared."); st.rerun()
                        with cc2:
                            if st.button("Cancel", key="super_admin_cancel_clear_inspector_bonus", use_container_width=True):
                                st.session_state["confirm_clear_inspector_bonus"] = False; st.rerun()
                with reset_col2:
                    if not st.session_state.get("confirm_clear_all_work_orders", False):
                        if st.button("🛠️ Clear All Work Orders", key="super_admin_clear_all_work_orders", type="secondary", use_container_width=True):
                            st.session_state["confirm_clear_all_work_orders"] = True
                    else:
                        st.warning("⚠️ This permanently removes ALL Work Order records.")
                        confirm_col, cancel_col = st.columns(2)
                        with confirm_col:
                            if st.button("✅ Confirm Clear Work Orders", key="super_admin_confirm_clear_all_work_orders", use_container_width=True):
                                save_all_work_orders([])
                                log_action("SUPER_ADMIN_CLEAR_WORK_ORDERS", "ALL", decision_by=full_name)
                                st.session_state["confirm_clear_all_work_orders"] = False
                                st.success("✅ All Work Order records have been cleared."); st.rerun()
                        with cancel_col:
                            if st.button("Cancel", key="super_admin_cancel_clear_all_work_orders", use_container_width=True):
                                st.session_state["confirm_clear_all_work_orders"] = False; st.rerun()
                with reset_col3:
                    if not st.session_state.get("confirm_clear_audit", False):
                        if st.button("🗑️ Clear Audit History", type="secondary", key="clear_audit_history_btn"):
                            st.session_state["confirm_clear_audit"] = True; st.rerun()
                    else:
                        st.error("⚠️ This will permanently remove the entire audit history.")
                        c1, c2 = st.columns(2)
                        with c1:
                            if st.button("✅ Yes, Clear Audit", type="primary", key="confirm_clear_audit_btn"):
                                clear_audit_log_file(); st.session_state["confirm_clear_audit"] = False
                                st.success("✅ Audit history has been cleared."); st.rerun()
                        with c2:
                            if st.button("↩️ Cancel", key="cancel_clear_audit_btn"):
                                st.session_state["confirm_clear_audit"] = False; st.rerun()
            st.divider()
            st.markdown("**🚀 Fresh Live Start**")
            st.caption("Use this when you are ready to go live and want both the request history and audit history to start empty. This does not delete users, departments, categories, permissions or uploaded files.")
            if not st.session_state.get("confirm_live_reset", False):
                if st.button("🚀 Prepare System for Live Use", type="primary", key="prepare_live_use_btn"):
                    st.session_state["confirm_live_reset"] = True; st.rerun()
            else:
                st.error("🚨 FINAL CONFIRMATION: All submitted requests AND the entire audit history will be permanently cleared.")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("🚀 Yes, Prepare for Live Use", type="primary", key="confirm_live_reset_btn"):
                        clear_live_request_and_audit_data()
                        st.session_state["confirm_live_reset"] = False
                        st.session_state["confirm_clear_requests"] = False
                        st.session_state["confirm_clear_audit"] = False
                        st.success("✅ Live launch reset complete. Request and audit history are now empty."); st.rerun()
                with c2:
                    if st.button("↩️ Cancel", key="cancel_live_reset_btn"):
                        st.session_state["confirm_live_reset"] = False; st.rerun()
        st.divider()
        st.subheader("📥 Download Data Backups")
        backup_col1, backup_col2, backup_col3 = st.columns(3)
        with backup_col1:
            if "EXCEL_PATH" in globals() and os.path.exists(EXCEL_PATH):
                with open(EXCEL_PATH, "rb") as f:
                    st.download_button("📥 Download Requests", f.read(), file_name=f"BACKUP_requests_{datetime.now().strftime('%Y-%m-%d')}.xlsx", type="primary", key="backup_requests")
        with backup_col2:
            if "USER_DB_PATH" in globals() and os.path.exists(USER_DB_PATH):
                with open(USER_DB_PATH, "rb") as f:
                    st.download_button("📥 Download Users", f.read(), file_name=f"BACKUP_users_{datetime.now().strftime('%Y-%m-%d')}.xlsx", type="primary", key="backup_users")
        with backup_col3:
            if "SETTINGS_PATH" in globals() and os.path.exists(SETTINGS_PATH):
                with open(SETTINGS_PATH, "rb") as f:
                    st.download_button("📥 Download Settings", f.read(), file_name=f"BACKUP_settings_{datetime.now().strftime('%Y-%m-%d')}.xlsx", type="primary", key="backup_settings")
        if os.path.exists(WORK_ORDERS_PATH):
            st.download_button("📥 Download Work Orders", open(WORK_ORDERS_PATH, "rb").read(), file_name=f"BACKUP_work_orders_{datetime.now().strftime('%Y-%m-%d')}.xlsx", type="primary", key="backup_work_orders")
        if os.path.exists(INSPECTOR_BONUS_PATH):
            st.download_button("📥 Download Inspector Bonus", open(INSPECTOR_BONUS_PATH, "rb").read(), file_name=f"BACKUP_inspector_bonus_{datetime.now().strftime('%Y-%m-%d')}.xlsx", type="primary", key="backup_inspector_bonus")
        st.caption("💾 Save these files to your computer for backup")

else:
    st.subheader("🔐 Access Restricted")
    st.error("❌ Your role does not have a defined portal. Please contact Super Admin.")
# ========================================================
# ✅ END OF ROLE-BASED PORTALS
# ============================================================
# ✅ END OF FILE — NOTHING AFTER THIS!
# ============================================================
