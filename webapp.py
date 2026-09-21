# ============================================================
# 🔄 ACOOLE PORTAL — PROFESSIONAL VERSION v4.18
# ============================================================
# ✅ v4.18 (ATTACHMENT PRESERVATION FIX):
#    • Fixed: Attachments were being cleared before submission
#      due to `clear_on_submit=True`. Now using versioned forms
#      to preserve uploads and reset forms cleanly after submit.
# ✅ v4.17 (BASE64 WHITESPACE & PADDING FIX):
#    • Strips whitespace/newlines from base64 secrets before decoding.
#    • Adds automatic padding to fix 'Incorrect padding' errors.
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
import re
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
# ============================================================
# ✅ YOUR PROVIDED FOLDER ID — MUST BE INSIDE A SHARED DRIVE
# ============================================================
GOOGLE_DRIVE_FOLDER_ID = "1U7gfbt38TsJArEfdANsYFKaJr3HrCcuw"
GOOGLE_DRIVE_ATTACHMENTS_FOLDER_NAME = "uploaded_attachments"

USER_DB_COLUMNS = [
    "full_name", "username", "password", "role", "dept",
    "can_view_all_dept", "can_generate_pdf", "can_download_data",
    "can_approve_requests", "can_access_inspector_bonus",
    "can_access_addition_deduction", "can_access_work_orders",
    "can_access_wo_total",
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
_LAST_DRIVE_UPLOAD_ERROR = ""

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
        media = MediaFileUpload(local_file_path, resumable=True)
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
        print(f"Google Drive upload failed: {e}")
        return None

def _drive_find_file(filename, parent_id=GOOGLE_DRIVE_FOLDER_ID):
    if drive_service is None: return None
    try:
        safe_name = str(filename).replace("'", "\\'")
        q = (f"name = '{safe_name}' and '{parent_id}' in parents and trashed = false")
        result = drive_service.files().list(q=q, spaces="drive", fields="files(id,name,modifiedTime,parents,mimeType)", orderBy="modifiedTime desc", pageSize=10, includeItemsFromAllDrives=True, supportsAllDrives=True).execute()
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
        media = MediaFileUpload(local_path, resumable=True)
        if file_id:
            return drive_service.files().update(fileId=file_id, media_body=media, fields="id,name", supportsAllDrives=True).execute()
        metadata = {"name": filename, "parents": [parent_id]}
        return drive_service.files().create(body=metadata, media_body=media, fields="id,name", supportsAllDrives=True).execute()
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
        raise e  # Re-raise to show real error

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
    """Synchronise one persistent workbook with Drive without background threads."""
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
    """Synchronously upload a changed workbook/file to Google Drive."""
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

def _drive_get_or_create_folder(folder_name, parent_id=GOOGLE_DRIVE_FOLDER_ID):
    """Return a Drive folder id, creating it under parent_id when necessary."""
    if drive_service is None:
        return None
    cache_key = f"folder::{parent_id}::{folder_name}"
    cached = _DRIVE_ID_CACHE.get(cache_key)
    if cached:
        return cached
    try:
        safe_name = str(folder_name).replace("'", "\\'")
        q = (
            f"name = '{safe_name}' and '{parent_id}' in parents "
            f"and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        )
        result = drive_service.files().list(
            q=q, spaces="drive", fields="files(id,name,parents,mimeType)", pageSize=10, includeItemsFromAllDrives=True, supportsAllDrives=True
        ).execute()
        folders = result.get("files", [])
        if folders:
            folder_id = folders[0]["id"]
        else:
            metadata = {
                "name": folder_name,
                "mimeType": "application/vnd.google-apps.folder",
                "parents": [parent_id],
            }
            created = drive_service.files().create(
                body=metadata, fields="id,name,parents", supportsAllDrives=True
            ).execute()
            folder_id = created.get("id")
        if folder_id:
            _DRIVE_ID_CACHE[cache_key] = folder_id
        return folder_id
    except Exception as e:
        print(f"Google Drive folder setup failed for {folder_name}: {e}")
        return None

def _upload_to_drive_bg(local_path, filename):
    """Upload a local file to Google Drive and return True on success."""
    global _LAST_DRIVE_UPLOAD_ERROR
    _LAST_DRIVE_UPLOAD_ERROR = ""
    if drive_service is None or not os.path.exists(local_path):
        _LAST_DRIVE_UPLOAD_ERROR = "Google Drive is not initialised or the local file does not exist."
        return False
    with _DRIVE_SYNC_LOCK:
        errors = []
        try:
            is_attachment = os.path.abspath(local_path).startswith(os.path.abspath(UPLOAD_DIR) + os.sep)
            candidate_parents = []
            if is_attachment:
                attachment_parent = _drive_get_or_create_folder(GOOGLE_DRIVE_ATTACHMENTS_FOLDER_NAME)
                if attachment_parent:
                    candidate_parents.append((attachment_parent, "uploaded_attachments"))
            candidate_parents.append((GOOGLE_DRIVE_FOLDER_ID, "Acoole_App_Uploads"))

            for parent_id, label in candidate_parents:
                try:
                    uploaded_id = _drive_upload_path(local_path, filename, parent_id=parent_id)
                    if uploaded_id:
                        return True
                    errors.append(f"{label}: Drive returned no file id")
                except Exception as e:
                    errors.append(f"{label}: {type(e).__name__}: {e}")

            _LAST_DRIVE_UPLOAD_ERROR = " | ".join(errors) or "No writable Drive destination was available."
            print(f"Google Drive attachment upload failed for {filename}: {_LAST_DRIVE_UPLOAD_ERROR}")
            return False
        except Exception as e:
            _LAST_DRIVE_UPLOAD_ERROR = f"{type(e).__name__}: {e}"
            print(f"Google Drive attachment upload failed for {filename}: {_LAST_DRIVE_UPLOAD_ERROR}")
            return False

def save_uploaded_attachment(uploaded_file, filename):
    """Save an uploaded attachment locally and require successful Drive backup."""
    # Sanitize filename to remove any problematic characters
    safe_name = "".join(c for c in str(filename) if c.isalnum() or c in "._- ").strip()
    if not safe_name:
        safe_name = "attachment"
    
    local_path = os.path.join(UPLOAD_DIR, safe_name)
    try:
        with open(local_path, "wb") as out_file:
            out_file.write(uploaded_file.getbuffer())
    except Exception as e:
        st.error(f"❌ Could not save attachment '{safe_name}': {e}")
        st.stop()
    if not _upload_to_drive_bg(local_path, safe_name):
        try:
            if os.path.exists(local_path):
                os.remove(local_path)
        except Exception:
            pass
        detail = str(_LAST_DRIVE_UPLOAD_ERROR or "No additional Drive error was returned.")
        st.error(
            f"❌ Attachment '{safe_name}' could not be uploaded to Google Drive. "
            "The record was NOT submitted."
        )
        st.caption(f"Google Drive upload detail: {detail}")
        st.stop()
    return safe_name

def initialise_drive_storage():
    if drive_service is None or st.session_state.get("drive_storage_initialised"): return
    os.makedirs(APP_FOLDER, exist_ok=True)
    with _DRIVE_SYNC_LOCK:
        _drive_get_or_create_folder(GOOGLE_DRIVE_ATTACHMENTS_FOLDER_NAME)
        targets = [
        (EXCEL_PATH, EXCEL_COLUMNS),
        (USER_DB_PATH, USER_DB_COLUMNS),
        (SETTINGS_PATH, ["setting", "value"]),
        (AUDIT_LOG_PATH, AUDIT_COLUMNS),
            (WORK_ORDERS_PATH, WORK_ORDER_COLUMNS),
            (INSPECTOR_BONUS_PATH, INSPECTOR_BONUS_COLUMNS),
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
    {"full_name": "National Grid Manager", "username": "national_grid", "password": "acoole123", "role": "Manager", "dept": "National Grid", "can_access_inspector_bonus": True, "can_access_addition_deduction": True, "can_access_work_orders": False, "can_access_wo_total": False, "is_active": True},
    {"full_name": "Isolator Manager", "username": "isolator", "password": "acoole123", "role": "Manager", "dept": "Isolator", "is_active": True},
    {"full_name": "Project Manager", "username": "project", "password": "acoole123", "role": "Manager", "dept": "Project", "is_active": True},
    {"full_name": "Accounts Manager", "username": "accounts", "password": "acoole123", "role": "Manager", "dept": "Accounts", "is_active": True},
    {"full_name": "Andy Acoole", "username": "andy", "password": "andy2026", "role": "Director", "dept": "ACoole Electrical Ltd", "is_active": True},
    {"full_name": "System Administrator", "username": "wais", "password": "superadmin123", "role": "Super Admin", "dept": "System Administration", "is_active": True},
    {"full_name": "Payroll Team", "username": "payroll", "password": "payroll2026", "role": "Payroll", "dept": "Payroll Department", "is_active": True}
]
PERMISSION_DEFAULTS = {
    "Work Order Employee": {"can_view_all_dept": False, "can_generate_pdf": False, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False, "can_access_addition_deduction": False, "can_access_work_orders": True, "can_access_wo_total": False},
    "Work Order Manager": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False, "can_access_addition_deduction": False, "can_access_work_orders": True, "can_access_wo_total": True},
    "Staff": {"can_view_all_dept": False, "can_generate_pdf": False, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False, "can_access_addition_deduction": True, "can_access_work_orders": False, "can_access_wo_total": False},
    "Team Member": {"can_view_all_dept": True, "can_generate_pdf": False, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": False, "can_access_addition_deduction": True, "can_access_work_orders": False, "can_access_wo_total": False},
    "Manager": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": False, "can_approve_requests": False, "can_access_inspector_bonus": True, "can_access_addition_deduction": True, "can_access_work_orders": False, "can_access_wo_total": False},
    "Director": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True, "can_approve_requests": True, "can_access_inspector_bonus": True, "can_access_addition_deduction": True, "can_access_work_orders": True, "can_access_wo_total": True},
    "Payroll": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True, "can_approve_requests": False, "can_access_inspector_bonus": True, "can_access_addition_deduction": True, "can_access_work_orders": True, "can_access_wo_total": True},
    "Super Admin": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True, "can_approve_requests": True, "can_access_inspector_bonus": True, "can_access_addition_deduction": True, "can_access_work_orders": True, "can_access_wo_total": True}
}
PERMISSION_LABELS = {
    "can_view_all_dept": "👁️ View All Department Requests",
    "can_generate_pdf": "📄 Generate & Download PDFs",
    "can_download_data": "📥 Download Data Backups",
    "can_approve_requests": "✅ Approve/Reject Requests",
    "can_access_inspector_bonus": "💰 National Grid Inspector Bonus",
    "can_access_addition_deduction": "➕ Addition & Deduction",
    "can_access_work_orders": "🛠️ Work Orders",
    "can_access_wo_total": "💷 Approved Work Order Total"
}

try:
    from fpdf import FPDF
    PDF_AVAILABLE = True
except ImportError:
    FPDF = None
    PDF_AVAILABLE = False

_EXCEL_INIT_LOCK = threading.Lock()

def _write_empty_excel(path, columns):
    """Create/replace an Excel file safely."""
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
    """Create/repair an Excel workbook safely without deleting a live file."""
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
    SETTING_ACTIONS = ["CATEGORY_ADDED", "CATEGORY_EDITED", "CATEGORY_DELETED", "DEPARTMENT_ADDED", "DEPARTMENT_EDITED", "DEPARTMENT_DELETED", "ROLE_ADDED", "ROLE_EDITED", "ROLE_DELETED", "USER_CREATED", "USER_EDITED", "USER_DELETED", "PASSWORD_CHANGED", "PASSWORD_RESET", "INSPECTOR_BONUS_CREATED", "INSPECTOR_BONUS_APPROVED", "INSPECTOR_BONUS_REJECTED", "INSPECTOR_BONUS_EDITED", "INSPECTOR_BONUS_STATUS_CHANGED", "SUPER_ADMIN_CLEAR_INSPECTOR_BONUS"]
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
        }
        display_action = action_labels.get(action, action)
        old_val = json.dumps(old_data, ensure_ascii=False)[:300] if old_data else "-"
        new_val = json.dumps(new_data, ensure_ascii=False)[:300] if new_data else "-"
        save_audit_entry({"AuditID": _get_next_audit_id(), "Timestamp": timestamp, "User_Name": username, "User_Role": role, "Action": display_action, "Request_ID": str(req_id), "Department": "-", "Amount": "-", "Decision_By": decision_by or "-", "Decision_Date": decision_date or "-", "Field_Changed": "Inspector Bonus", "Old_Value": old_val, "New_Value": new_val, "IP_Address": "Auto-Logged"})
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
        with st.spinner("Refreshing data..."):
            _invalidate_data_cache("_records_cache", "_users_cache", "_settings_cache", "_audit_log_cache", "_work_orders_cache", "_inspector_bonus_cache", "_audit_log_count")
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
        pd.DataFrame([{"setting": "categories", "value": "|".join(DEFAULT_CATEGORIES)}, {"setting": "roles", "value": "|".join(DEFAULT_ROLES)}, {"setting": "departments", "value": "|".join(DEFAULT_DEPARTMENTS)}]).to_excel(SETTINGS_PATH, index=False, engine="openpyxl")

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
                "is_active": _active_or_default(r.get("is_active", ""))
            }
            if user_role == "Super Admin":
                users[username].update({
                    "can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True,
                    "can_approve_requests": True, "can_access_inspector_bonus": True,
                    "can_access_addition_deduction": True, "can_access_work_orders": True,
                    "can_access_wo_total": True, "is_active": True
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
                                    saved_fn = save_uploaded_attachment(f, fn)
                                    new_attachments.append(saved_fn)
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
                        saved_fn = save_uploaded_attachment(f, fn)
                        attachments.append(saved_fn)
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
                    saved_fn = save_uploaded_attachment(f, fn)
                    attachments.append(saved_fn)
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
                        st
