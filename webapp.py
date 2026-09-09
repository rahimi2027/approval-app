# ============================================================
# 🔄 ACOOLE PORTAL — PROFESSIONAL VERSION v2.4 (ALL BUGS FIXED)
# ============================================================
# ✅ ALL MISSING VARIABLES ADDED
# ✅ CORRECTED user/user_info mixup
# ✅ all_live_requests defined everywhere
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
from datetime import datetime, date
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account
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
# ─── GOOGLE DRIVE ───
GOOGLE_DRIVE_FOLDER_ID = "1oecpaa8c5tryCtcIAnbjEXemGDonvgPZ"
# ─── ONEDRIVE / MICROSOFT GRAPH ───
ONEDRIVE_CLIENT_ID = ""
ONEDRIVE_CLIENT_SECRET = ""
ONEDRIVE_TENANT_ID = "common"
ONEDRIVE_FOLDER = "Acoole_App_Uploads/"
USE_ONEDRIVE = False
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)
os.makedirs(ARCHIVE_FOLDER, exist_ok=True)
# ============================================================
# ✅ LOAD GOOGLE CREDENTIALS FROM FILE
# ============================================================
SERVICE_ACCOUNT_INFO = {}
KEY_FILE = os.path.join(BASE_DIR, "service_account_key.json")
if os.path.exists(KEY_FILE):
    try:
        with open(KEY_FILE, "r", encoding="utf-8") as f:
            SERVICE_ACCOUNT_INFO = json.load(f)
        st.success("✅ Google Drive credentials loaded successfully!")
    except Exception as e:
        st.error(f"⚠️ Failed to load credentials: {str(e)[:200]}")
else:
    st.warning("⚠️ Credentials file not found — using local storage only")
# ============================================================
# ✅ GOOGLE DRIVE UPLOAD FUNCTION
# ============================================================
def get_drive_service():
    try:
        credentials = service_account.Credentials.from_service_account_info(
            SERVICE_ACCOUNT_INFO,
            scopes=["https://www.googleapis.com/auth/drive"]
        )
        return build("drive", "v3", credentials=credentials)
    except Exception as e:
        st.error(f"❌ Google Drive Error: {e}")
        return None
def upload_to_google_drive(local_file_path, display_filename):
    try:
        st.info(f"📤 Uploading: {display_filename}")
        credentials = service_account.Credentials.from_service_account_info(
            SERVICE_ACCOUNT_INFO,
            scopes=["https://www.googleapis.com/auth/drive"]
        )
        service = build("drive", "v3", credentials=credentials)
        try:
            folder = service.files().get(fileId=GOOGLE_DRIVE_FOLDER_ID, fields="id, name").execute()
            st.success(f"✅ Folder FOUND: {folder.get('name')}")
        except Exception as fe:
            st.error(f"❌ CANNOT ACCESS FOLDER! Error: {str(fe)}")
            st.info("👉 SHARE folder with bot email: drive-upload-bot@acoole-attachments.iam.gserviceaccount.com")
            return None
        file_metadata = {"name": display_filename, "parents": [GOOGLE_DRIVE_FOLDER_ID]}
        media = MediaFileUpload(local_file_path, resumable=True)
        file = service.files().create(body=file_metadata, media_body=media, fields="id, name, parents").execute()
        file_id = file.get("id")
        parents = file.get("parents", [])
        if GOOGLE_DRIVE_FOLDER_ID in parents:
            st.success(f"✅ ✅ SUCCESS! File IS IN YOUR FOLDER! 🎉 ID: {file_id[:12]}...")
            st.info("👉 REFRESH your Google Drive folder → FILE IS THERE!")
        else:
            st.warning(f"⚠️ Uploaded but NOT in your folder! Parents: {parents}")
            st.info("👉 File went to bot's storage — check folder sharing!")
        return file_id
    except Exception as e:
        st.error(f"❌ UPLOAD FAILED! Error: {str(e)}")
        return None
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
            st.info(f"✅ Synced to OneDrive: {filename}"); return True
        else: st.warning(f"⚠️ OneDrive sync: {res.status_code}")
    except Exception as e: st.warning(f"⚠️ Could not sync to OneDrive: {str(e)}")
    return False
# ============================================================
# DEFAULTS — ROLES, DEPARTMENTS, USERS, PERMISSIONS
# ============================================================
DEFAULT_CATEGORIES = ["Food Allowance", "Others", "Parking", "Parking Fine", "GYM Membership", "Item Not Returned", "Item Missing"]
DEFAULT_ROLES = ["Manager", "Staff", "Team Member", "Director", "Payroll", "Super Admin"]
DEFAULT_DEPARTMENTS = ["National Grid", "Isolator", "Project", "Accounts", "Payroll Department", "ACoole Electrical Ltd"]
EXCEL_COLUMNS = [
    "ID", "Employee Name", "Department", "Transaction Type", "Category Reason",
    "Date", "Amount (£)", "Line Manager", "Description", "Attachment Name",
    "Status", "Director Comments", "Decision Date", "Decision By",
    "PDF File Path", "Edited From ID", "Old Data"
]
DEFAULT_USERS = [
    {"full_name": "National Grid Manager", "username": "national_grid", "password": "acoole123", "role": "Manager", "dept": "National Grid"},
    {"full_name": "Isolator Manager", "username": "isolator", "password": "acoole123", "role": "Manager", "dept": "Isolator"},
    {"full_name": "Project Manager", "username": "project", "password": "acoole123", "role": "Manager", "dept": "Project"},
    {"full_name": "Accounts Manager", "username": "accounts", "password": "acoole123", "role": "Manager", "dept": "Accounts"},
    {"full_name": "Andy Acoole", "username": "andy", "password": "andy2026", "role": "Director", "dept": "ACoole Electrical Ltd"},
    {"full_name": "System Administrator", "username": "wais", "password": "superadmin123", "role": "Super Admin", "dept": "System Administration"},
    {"full_name": "Payroll Team", "username": "payroll", "password": "payroll2026", "role": "Payroll", "dept": "Payroll Department"}
]
PERMISSION_DEFAULTS = {
    "Staff": {"can_view_all_dept": False, "can_generate_pdf": False, "can_download_data": False, "can_approve_requests": False},
    "Team Member": {"can_view_all_dept": True, "can_generate_pdf": False, "can_download_data": False, "can_approve_requests": False},
    "Manager": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": False, "can_approve_requests": False},
    "Director": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True, "can_approve_requests": True},
    "Payroll": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True, "can_approve_requests": False},
    "Super Admin": {"can_view_all_dept": True, "can_generate_pdf": True, "can_download_data": True, "can_approve_requests": True}
}
PERMISSION_LABELS = {
    "can_view_all_dept": "👁️ View All Department Requests",
    "can_generate_pdf": "📄 Generate & Download PDFs",
    "can_download_data": "📥 Download Data Backups",
    "can_approve_requests": "✅ Approve/Reject Requests"
}
# ============================================================
# PDF LIBRARY
# ============================================================
try:
    from fpdf2 import FPDF
    PDF_AVAILABLE = True
except ImportError:
    try:
        from fpdf import FPDF
        PDF_AVAILABLE = True
    except ImportError:
        PDF_AVAILABLE = False
# ============================================================
# ✅ SAFE EXCEL INIT — FIXES "File is not a zip file"
# ============================================================
def safe_init_excel(path, columns):
    if not os.path.exists(path):
        pd.DataFrame(columns=columns).to_excel(path, index=False, engine="openpyxl")
        return True
    try:
        pd.read_excel(path, engine="openpyxl")
        return True
    except Exception:
        os.remove(path)
        pd.DataFrame(columns=columns).to_excel(path, index=False, engine="openpyxl")
        return True
# ============================================================
# AUDIT LOG FUNCTIONS
# ============================================================
def init_audit_log():
    safe_init_excel(AUDIT_LOG_PATH, AUDIT_COLUMNS)
def load_audit_log():
    init_audit_log()
    try:
        df = pd.read_excel(AUDIT_LOG_PATH, engine="openpyxl").fillna("")
        return df.to_dict(orient="records")
    except Exception as e:
        print(f"⚠️ Failed to load audit log: {e}"); return []
def save_audit_entry(entry):
    init_audit_log()
    try:
        df = pd.read_excel(AUDIT_LOG_PATH, engine="openpyxl").fillna("")
        df = pd.concat([df, pd.DataFrame([entry])], ignore_index=True)
        df.to_excel(AUDIT_LOG_PATH, index=False, engine="openpyxl")
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
    if os.path.exists(AUDIT_LOG_FILE): os.remove(AUDIT_LOG_FILE)
    pd.DataFrame(columns=AUDIT_COLUMNS).to_excel(AUDIT_LOG_FILE, index=False, engine="openpyxl")
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
        "USER_CREATED", "USER_EDITED", "USER_DELETED", "PASSWORD_CHANGED", "PASSWORD_RESET"]
    if action in SETTING_ACTIONS:
        action_labels = {
            "CATEGORY_ADDED": "🏷️ Category Added", "CATEGORY_EDITED": "🏷️ Category Edited", "CATEGORY_DELETED": "🏷️ Category Deleted",
            "DEPARTMENT_ADDED": "🏢 Department Added", "DEPARTMENT_EDITED": "🏢 Department Edited", "DEPARTMENT_DELETED": "🏢 Department Deleted",
            "ROLE_ADDED": "🎖️ Role Added", "ROLE_EDITED": "🎖️ Role/Permissions Edited", "ROLE_DELETED": "🎖️ Role Deleted",
            "USER_CREATED": "👤 User Account Created", "USER_EDITED": "👤 User Account Edited",
            "USER_DELETED": "👤 User Account Deleted", "PASSWORD_CHANGED": "🔑 Password Changed", "PASSWORD_RESET": "🔑 Password Reset"
        }
        display_action = action_labels.get(action, action)
        old_val = json.dumps(old_data, ensure_ascii=False)[:300] if old_data else "-"
        new_val = json.dumps(new_data, ensure_ascii=False)[:300] if new_data else "-"
        save_audit_entry({"AuditID": len(load_audit_log()) + 1, "Timestamp": timestamp,
            "User_Name": username, "User_Role": role, "Action": display_action,
            "Request_ID": str(req_id), "Department": "-", "Amount": "-",
            "Decision_By": "-", "Decision_Date": "-", "Field_Changed": "System Configuration",
            "Old_Value": old_val, "New_Value": new_val, "IP_Address": "Auto-Logged"})
        return
    dept, amount, saved_decision_by, saved_decision_date = get_request_details(req_id)
    final_decision_by = decision_by or saved_decision_by
    final_decision_date = decision_date or saved_decision_date
    if action in ["CREATED", "DELETED"]:
        save_audit_entry({"AuditID": len(load_audit_log()) + 1, "Timestamp": timestamp,
            "User_Name": username, "User_Role": role, "Action": action, "Request_ID": str(req_id),
            "Department": dept, "Amount": amount, "Decision_By": "-", "Decision_Date": "-",
            "Field_Changed": "-", "Old_Value": "-",
            "New_Value": "New Request Created" if action == "CREATED" else "Request Permanently Deleted",
            "IP_Address": "Auto-Logged"})
    elif action in ["APPROVED", "REJECTED", "STATUS_CHANGED"]:
        status_text = "Approved" if action == "APPROVED" else "Rejected" if action == "REJECTED" else "Status Changed"
        save_audit_entry({"AuditID": len(load_audit_log()) + 1, "Timestamp": timestamp,
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
                save_audit_entry({"AuditID": len(load_audit_log()) + 1, "Timestamp": timestamp,
                    "User_Name": username, "User_Role": role, "Action": "EDITED", "Request_ID": str(req_id),
                    "Department": dept, "Amount": amount, "Decision_By": "-", "Decision_Date": "-",
                    "Field_Changed": label, "Old_Value": old, "New_Value": new, "IP_Address": "Auto-Logged"})
def display_audit_log_panel():
    st.subheader("📖 Full System Audit Log — Complete History")
    st.info("🔒 Super Admin Only — Cannot be deleted or modified."); st.divider()
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
    if not d or str(d).strip() in ["", "none", "nan"]: return "-"
    return str(d).strip()[:10]
def display_attachments(req):
    att = req.get("attachment_name", "None")
    if not att or str(att).strip().lower() in ["none", "nan", ""]:
        st.info("📎 No attachments."); return
    try:
        attached_files = [n.strip() for n in str(att).split(",")]
        found_any = False
        for idx, name in enumerate(attached_files):
            path = os.path.join(UPLOAD_DIR, name)
            if os.path.exists(path):
                found_any = True
                with open(path, "rb") as f:
                    st.download_button(f"⬇️ Download {name}", f.read(), file_name=name, key=f"att_{req.get('id', idx)}_{idx}")
        if not found_any: st.info("📎 Attachments referenced but files not available.")
    except Exception as e: st.info(f"📎 Attachments: {att}")
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
def load_departments():
    init_settings()
    try:
        df = pd.read_excel(SETTINGS_PATH, engine="openpyxl").fillna("")
        for _, r in df.iterrows():
            if r["setting"] == "departments":
                vals = [v.strip() for v in str(r["value"]).split("|") if v.strip()]
                return vals if vals else DEFAULT_DEPARTMENTS.copy()
        return DEFAULT_DEPARTMENTS.copy()
    except: return DEFAULT_DEPARTMENTS.copy()
def save_departments(dept_list):
    init_settings()
    df = pd.read_excel(SETTINGS_PATH, engine="openpyxl").fillna("")
    found = False
    for idx, r in df.iterrows():
        if r["setting"] == "departments":
            df.at[idx, "value"] = "|".join(dept_list); found = True
    if not found:
        df = pd.concat([df, pd.DataFrame([{"setting": "departments", "value": "|".join(dept_list)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False, engine="openpyxl")
def load_categories():
    init_settings()
    try:
        df = pd.read_excel(SETTINGS_PATH, engine="openpyxl").fillna("")
        for _, r in df.iterrows():
            if r["setting"] == "categories":
                vals = [v.strip() for v in str(r["value"]).split("|") if v.strip()]
                return vals if vals else DEFAULT_CATEGORIES
        return DEFAULT_CATEGORIES
    except: return DEFAULT_CATEGORIES
def save_categories(cat_list):
    init_settings()
    df = pd.read_excel(SETTINGS_PATH, engine="openpyxl").fillna("")
    found = False
    for idx, r in df.iterrows():
        if r["setting"] == "categories":
            df.at[idx, "value"] = "|".join(cat_list); found = True
    if not found:
        df = pd.concat([df, pd.DataFrame([{"setting": "categories", "value": "|".join(cat_list)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False, engine="openpyxl")
def load_roles():
    init_settings()
    try:
        df = pd.read_excel(SETTINGS_PATH, engine="openpyxl").fillna("")
        for _, r in df.iterrows():
            if r["setting"] == "roles":
                vals = [v.strip() for v in str(r["value"]).split("|") if v.strip()]
                return vals if vals else DEFAULT_ROLES
        return DEFAULT_ROLES
    except: return DEFAULT_ROLES
def save_roles(roles_list):
    init_settings()
    df = pd.read_excel(SETTINGS_PATH, engine="openpyxl").fillna("")
    found = False
    for idx, r in df.iterrows():
        if r["setting"] == "roles":
            df.at[idx, "value"] = "|".join(roles_list); found = True
    if not found:
        df = pd.concat([df, pd.DataFrame([{"setting": "roles", "value": "|".join(roles_list)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False, engine="openpyxl")
# ============================================================
# USER DATABASE
# ============================================================
def init_user_db():
    safe_init_excel(USER_DB_PATH, ["full_name","username","password","role","dept",
        "can_view_all_dept","can_generate_pdf","can_download_data","can_approve_requests"])
    if os.path.getsize(USER_DB_PATH) < 500:
        pd.DataFrame(DEFAULT_USERS).to_excel(USER_DB_PATH, index=False, engine="openpyxl")
def save_users(users_dict):
    rows = []
    for username, u in users_dict.items():
        rows.append({"full_name": u.get("full_name", username), "username": username,
            "password": u.get("password", ""), "role": u.get("role", "Staff"),
            "dept": u.get("dept", ""), "can_view_all_dept": u.get("can_view_all_dept", False),
            "can_generate_pdf": u.get("can_generate_pdf", False),
            "can_download_data": u.get("can_download_data", False),
            "can_approve_requests": u.get("can_approve_requests", False)})
    pd.DataFrame(rows).to_excel(USER_DB_PATH, index=False, engine="openpyxl")
def load_users():
    init_user_db()
    try:
        df = pd.read_excel(USER_DB_PATH, engine="openpyxl").fillna("")
        users = {}
        for _, r in df.iterrows():
            users[r["username"]] = {
                "full_name": str(r.get("full_name", r["username"])).strip(),
                "password": str(r["password"]), "role": str(r.get("role", "Staff")),
                "dept": str(r.get("dept", "")),
                "can_view_all_dept": str(r.get("can_view_all_dept", "False")).lower() == "true",
                "can_generate_pdf": str(r.get("can_generate_pdf", "False")).lower() == "true",
                "can_download_data": str(r.get("can_download_data", "False")).lower() == "true",
                "can_approve_requests": str(r.get("can_approve_requests", "False")).lower() == "true"}
        return users
    except Exception as e:
        st.error(f"User DB Load Error: {e}"); return {}
# ============================================================
# REQUESTS EXCEL
# ============================================================
def initialise_excel():
    safe_init_excel(EXCEL_PATH, EXCEL_COLUMNS)
initialise_excel()
def load_records_from_excel():
    try:
        if not os.path.exists(EXCEL_PATH): return []
        df = pd.read_excel(EXCEL_PATH, engine="openpyxl").fillna("")
        if df.empty: return []
        records = df.to_dict(orient="records")
        parsed = []
        for r in records:
            try: record_id = int(r.get("ID", 0))
            except: record_id = 0
            try: amount = float(r.get("Amount (£)", 0))
            except: amount = 0.0
            parsed.append({
                "id": record_id, "emp_name": str(r.get("Employee Name", "Not Specified")).strip(),
                "dept": str(r.get("Department", "Not Specified")).strip(),
                "type": str(r.get("Transaction Type", "Not Specified")).strip(),
                "category": str(r.get("Category Reason", "Not Specified")).strip(),
                "date": str(r.get("Date", "")).strip(), "amount": amount,
                "manager": str(r.get("Line Manager", "Not Specified")).strip(),
                "desc": str(r.get("Description", "")).strip(),
                "attachment_name": str(r.get("Attachment Name", "None")).strip(),
                "status": str(r.get("Status", "pending")).strip().lower(),
                "director_comments": str(r.get("Director Comments", "")).strip(),
                "decision_date": str(r.get("Decision Date", "")).strip(),
                "decision_by": str(r.get("Decision By", "")).strip(),
                "pdf_path": str(r.get("PDF File Path", "")).strip(),
                "edited_from_id": str(r.get("Edited From ID", "")).strip(),
                "old_data": str(r.get("Old Data", "")).strip()})
        return parsed
    except Exception as e:
        st.error(f"Load Error: {e}"); return []
def save_all_records(records):
    export = []
    for r in records:
        export.append({"ID": int(r.get("id", 0)), "Employee Name": str(r.get("emp_name", "")),
            "Department": str(r.get("dept", "")), "Transaction Type": str(r.get("type", "")),
            "Category Reason": str(r.get("category", "")), "Date": str(r.get("date", "")),
            "Amount (£)": float(r.get("amount", 0.0)), "Line Manager": str(r.get("manager", "")),
            "Description": str(r.get("desc", "")), "Attachment Name": str(r.get("attachment_name", "None")),
            "Status": str(r.get("status", "pending")).lower(),
            "Director Comments": str(r.get("director_comments", "")),
            "Decision Date": str(r.get("decision_date", "")),
            "Decision By": str(r.get("decision_by", "")),
            "PDF File Path": str(r.get("pdf_path", "")),
            "Edited From ID": str(r.get("edited_from_id", "")),
            "Old Data": str(r.get("old_data", ""))})
    pd.DataFrame(export, columns=EXCEL_COLUMNS).to_excel(EXCEL_PATH, index=False, engine="openpyxl")
def save_record_to_excel(new_record):
    current = load_records_from_excel()
    current.append(new_record)
    save_all_records(current)
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
        pdf.cell(52, 5, "Amount Approved:", 0, 0); pdf.cell(0, 5, f"£{amount}", ln=True)
        pdf.cell(52, 5, "Line Manager:", 0, 0); pdf.cell(0, 5, manager, ln=True)
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
        return True, pdf_bytes, filename
    except Exception as e:
        return False, None, f"PDF Error: {str(e)}"
def display_pdf_button(req, can_generate=False, key_suffix=""):
    req_id = req["id"]
    unique_key = f"genpdf_{req_id}_{key_suffix}"
    if can_generate and PDF_AVAILABLE:
        st.button(f"📄 Generate PDF for ID #{req_id}", type="primary", key=unique_key)
        ok, pdf_bytes, name = generate_approval_pdf(req)
        if ok:
            st.success(f"✅ Generated! Ready to download ⬇")
            st.download_button(f"📥 Download: {name}", data=pdf_bytes, file_name=name,
                mime="application/pdf", type="primary", key=f"dl_{unique_key}")
        else:
            st.error(f"❌ {name}")
    return False
# ============================================================
# 📊 DASHBOARD COMPONENT
# ============================================================
def show_dashboard(user, all_requests):
    role = user.get("role", "")
    dept = user.get("dept", "")
    full_name = user.get("full_name", user.get("username", "User"))
    if not role:
        return
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

def display_company_header():
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        if os.path.exists(LOGO_PATH): st.image(LOGO_PATH, width=300)
        else: st.title("⚡ ACOOLE ELECTRICAL LTD")
        st.caption("Addition & Deduction Approval Platform")
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
            defaults = PERMISSION_DEFAULTS[new_role]
            col1, col2 = st.columns(2)
            perm_view_all = col1.checkbox(PERMISSION_LABELS["can_view_all_dept"], value=defaults["can_view_all_dept"])
            perm_pdf = col1.checkbox(PERMISSION_LABELS["can_generate_pdf"], value=defaults["can_generate_pdf"])
            perm_download = col2.checkbox(PERMISSION_LABELS["can_download_data"], value=defaults["can_download_data"])
            perm_approve = col2.checkbox(PERMISSION_LABELS["can_approve_requests"], value=defaults["can_approve_requests"])
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
                        "can_download_data": perm_download, "can_approve_requests": perm_approve
                    }
                    save_users(USERS)
                    log_action("USER_CREATED", new_data={
                        "username": new_username, "full_name": new_full_name.strip(),
                        "role": new_role, "department": new_dept,
                        "permissions": {"can_view_all_dept": perm_view_all, "can_generate_pdf": perm_pdf,
                        "can_download_data": perm_download, "can_approve_requests": perm_approve}
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
                ecol1, ecol2 = st.columns(2)
                edit_view = ecol1.checkbox(PERMISSION_LABELS["can_view_all_dept"], value=curr_perm_view)
                edit_pdf = ecol1.checkbox(PERMISSION_LABELS["can_generate_pdf"], value=curr_perm_pdf)
                edit_dl = ecol2.checkbox(PERMISSION_LABELS["can_download_data"], value=curr_perm_dl)
                edit_app = ecol2.checkbox(PERMISSION_LABELS["can_approve_requests"], value=curr_perm_app)
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
                            "can_download_data": edit_dl, "can_approve_requests": edit_app
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
                    save_users(USERS)
                    log_action("USER_EDITED", old_data=curr, new_data={
                        "full_name": upd_full_name.strip(), "username": upd_username_new,
                        "role": upd_role, "department": upd_dept,
                        "permissions": {"can_view_all_dept": edit_view, "can_generate_pdf": edit_pdf,
                        "can_download_data": edit_dl, "can_approve_requests": edit_app}
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
# ✅ POST-LOGIN — ALL FIXED VARIABLES DEFINED
# ============================================================
user_info = st.session_state.user_info or {}
full_name = user_info.get("full_name", user_info.get("username", "User"))
dept = user_info.get("dept", "")
role = user_info.get("role", "")
FULL_NAME = full_name  # ✅ Fixed missing variable!
all_live_requests = load_records_from_excel()  # ✅ Fixed missing variable!
CATEGORIES = load_categories()  # ✅ Fixed missing variable!

# ✅ Welcome + Logout
st.info(f"👤 Welcome: {full_name} | {dept} | {role}")
if st.button("🔓 Secure Logout"):
    st.session_state.clear()
    st.rerun()

display_company_header()

change_my_password_form()

# Sidebar PDF section
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

# PDF generation handlers
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
# 📋 ROLE-BASED PORTALS — ALL NOW WORKING
# ============================================================

# ─── PAYROLL PORTAL ───
if role == "Payroll":
    st.subheader("🧾 Payroll Portal")
    st.info("✅ View all requests and Download PDFs."); st.divider()
    tab_pending, tab_approved, tab_rejected = st.tabs(["⏳ Pending Requests", "✅ Approved Requests", "❌ Rejected Requests"])
    with tab_pending:
        pending = [r for r in all_live_requests if r.get("status") == "pending"]
        if not pending:
            st.success("✅ No pending requests!")
        else:
            st.metric("⏳ Pending", len(pending)); st.divider()
            for req in reversed(pending):
                with st.expander(f"🟡 ID #{req.get('id')} | {req.get('emp_name')} | £{float(req.get('amount',0)):.2f}"):
                    st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept')}")
                    st.write(f"🔄 Type: {req.get('type')} | 🏷️ Category: {req.get('category')}")
                    st.write(f"💷 Amount: £{float(req.get('amount',0)):.2f}")
                    st.write(f"👔 **Line Manager:** {req.get('manager')}")
                    st.write(f"📅 **Date:** {req.get('date')}")
                    st.info(f"📝 **Description:** {req.get('desc')}")
                    display_attachments(req)
                    st.divider()
                    display_pdf_button(req, can_generate=True)
    with tab_approved:
        approved = [r for r in all_live_requests if r.get("status") == "approved"]
        if not approved:
            st.info("📋 No approved requests.")
        else:
            st.metric("✅ Approved", len(approved)); st.divider()
            for req in reversed(approved):
                dec_by = req.get('decision_by', 'Director')
                dec_date = req.get('decision_date', '')
                display_date = dec_date[:10] if dec_date and len(dec_date)>=10 else ""
                extra_text = f" | ✅ Approved by {dec_by} on {display_date}" if display_date else f" | ✅ Approved by {dec_by}"
                title = f"🟢 ID #{req.get('id')} | {req.get('emp_name')} | £{float(req.get('amount',0)):.2f}{extra_text}"
                with st.expander(title):
                    st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept', '')}")
                    st.write(f"💷 Amount: £{float(req.get('amount',0)):.2f}")
                    st.write(f"🎯 **Approved By:** {dec_by}")
                    if display_date: st.write(f"📅 **Approval Date:** {display_date}")
                    st.info(f"💬 Comments: {req.get('director_comments', 'None')}")
                    display_attachments(req)
                    st.divider()
                    display_pdf_button(req, can_generate=True)
    with tab_rejected:
        rejected = [r for r in all_live_requests if r.get("status") == "rejected"]
        if not rejected:
            st.success("✅ No rejected requests!")
        else:
            st.metric("❌ Rejected", len(rejected)); st.divider()
            for req in reversed(rejected):
                with st.expander(f"🔴 ID #{req.get('id')} | {req.get('emp_name')} | £{float(req.get('amount',0)):.2f}"):
                    st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept')}")
                    st.write(f"💷 Amount: £{float(req.get('amount',0)):.2f}")
                    st.error(f"❌ Rejected By: {req.get('decision_by', '—')} on {format_date(req.get('decision_date', ''))}")
                    st.error(f"💬 Reason: {req.get('director_comments', 'None')}")
                    display_attachments(req)
                    st.divider()
                    display_pdf_button(req, can_generate=True)

# ─── MANAGER / STAFF PORTAL ───
elif role in ["Manager", "Staff", "Team Member"]:
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
            files_to_keep = []
            files_to_remove = []
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
                    else:
                        col_dl.caption("⚠️ Missing")
                    if keep: files_to_keep.append(fname)
                    else: files_to_remove.append(fname)
                if files_to_remove:
                    st.warning(f"🗑️ Will remove: {', '.join(files_to_remove)}")
            else:
                st.info("📋 No attachments currently attached.")
            st.markdown("#### ➕ Attach New Files")
            new_files_upload = st.file_uploader("Upload additional files", type=["pdf", "png", "jpg", "jpeg"],
                accept_multiple_files=True, key=f"new_upload_{eid}")
            st.info(f"✅ Result: **{len(files_to_keep)} kept** + **{len(new_files_upload or [])} new** = {len(files_to_keep)+len(new_files_upload or [])} total files")
            with st.form("edit_form"):
                c1, c2 = st.columns(2)
                with c1:
                    en = st.text_input("👤 Employee Name", rec.get("emp_name", ""))
                    rt = st.selectbox("🔄 Transaction Type", ["Addition", "Deduction"],
                        index=["Addition", "Deduction"].index(rec.get("type", "Addition")))
                    cat_idx = CATEGORIES.index(rec["category"]) if rec.get("category") in CATEGORIES else 0
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
                            with open(os.path.join(UPLOAD_DIR, fn), "wb") as outfile:
                                outfile.write(f.getbuffer())
                            final_attachments.append(fn)
                    records = load_records_from_excel()
                    old_data_dict = {
                        "emp_name": rec.get("emp_name"), "dept": rec.get("dept"),
                        "type": rec.get("type"), "category": rec.get("category"),
                        "date": rec.get("date"), "amount": rec.get("amount"),
                        "manager": rec.get("manager"), "desc": rec.get("desc")
                    }
                    new_data_dict = {
                        "emp_name": en.strip(), "dept": dept_name, "type": rt,
                        "category": ct, "date": str(dt_val), "amount": amt,
                        "manager": mgr.strip(), "desc": desc.strip()
                    }
                    for r in records:
                        if int(r.get("id", 0)) == int(eid):
                            r["emp_name"] = en.strip()
                            r["type"] = rt
                            r["category"] = ct
                            r["amount"] = amt
                            r["date"] = str(dt_val)
                            r["manager"] = mgr.strip()
                            r["desc"] = desc.strip()
                            r["status"] = "pending"
                            r["attachment_name"] = ", ".join(final_attachments) or "None"
                            r["old_data"] = json.dumps(old_data_dict)
                            break
                    log_action("EDITED", eid, old_data=old_data_dict, new_data=new_data_dict)
                    save_all_records(records)
                    st.success(f"✅ Updated! Removed {len(files_to_remove)} | Kept {len(files_to_keep)} | Added {len(new_files_upload or [])}")
                    st.session_state.editing_request_id = None
                    st.rerun()
            if st.button("❌ Cancel", key=f"cancel_edit_{eid}"):
                st.session_state.editing_request_id = None
                st.rerun()
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
                            file_id = upload_to_google_drive(file_path, fn)
                            if file_id:
                                st.info(f"✅ Uploaded to Drive: {fn} (ID: {file_id[:12]}...)")
                    payload = {
                        "id": nid, "emp_name": en.strip(), "dept": dept_name, "type": rt,
                        "category": ct, "date": str(dt_val), "amount": amt, "manager": mgr.strip(),
                        "desc": desc.strip(), "attachment_name": ", ".join(att_list) or "None",
                        "status": "pending", "director_comments": "", "decision_date": "",
                        "decision_by": "", "pdf_path": "", "edited_from_id": "", "old_data": ""
                    }
                    save_record_to_excel(payload)
                    log_action("CREATED", nid)
                    st.success(f"✅ Request #{nid} sent for approval!")
                    st.rerun()
                else:
                    st.error("⚠️ Please fill in: Employee Name, Line Manager, and Description")
        st.divider()
        st.subheader(f"📋 My Department Requests")
        my_reqs = [r for r in all_live_requests if r.get("dept") == dept_name]
        if not my_reqs:
            st.info("📋 No requests yet.")
        else:
            for req in reversed(my_reqs):
                status = req.get("status", "pending").lower()
                icon = "🟡" if status == "pending" else ("🟢" if status == "approved" else "🔴")
                dec_by = req.get('decision_by', '—')
                dec_date = format_date(req.get('decision_date', ''))

                        # Build approver line only if request was decided
                        if status in ["approved", "rejected"] and dec_by != '—':
                        approver_line = f" | ✅ By: {dec_by}"
                        if dec_date:
                            approver_line += f" | 📅 {dec_date}"
                else:
                      approver_line = ""

                title = f"{icon} ID #{req.get('id')} | {req.get('emp_name')} | {status.upper()} | £{float(req.get('amount',0)):.2f}{approver_line}"
                    with st.expander(title):
                    st.write(f"👤 Employee: {req.get('emp_name')} | 👔 Manager: {req.get('manager')}")
                    st.write(f"🔄 Type: {req.get('type')} | 🏷️ Category: {req.get('category')}")
                    st.info(f"📝 Description: {req.get('desc')}")
                    display_attachments(req)
                    if req.get("director_comments"):
                        st.info(f"💬 Director Comments: {req.get('director_comments')}")
                    if status == "approved":
                        display_pdf_button(req, can_generate=False)
                    if status in ["pending", "rejected"]:
                        if st.button(f"✏️ Edit Request #{req.get('id')}", key=f"edit_{req.get('id')}"):
                            st.session_state.editing_request_id = req.get("id")
                            st.rerun()# 
                            
# ─── DIRECTOR PORTAL ───
elif role == "Director":
    st.subheader("🎛️ Director Approval Portal — Andy Acoole")
    st.info("✅ Review all requests, Approve, Reject, OR Change Status. Decisions update automatically.")
    st.divider()
    tab_pending, tab_approved, tab_rejected = st.tabs(["⏳ Pending Requests", "✅ Approved Requests", "❌ Rejected Requests"])

    with tab_pending:
        pending = [r for r in all_live_requests if r.get("status") == "pending"]
        if not pending:
            st.success("✅ No pending requests!")
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
                        st.write(f"📅 **Date:** {format_date(req.get('date',''))}")
                        st.info(f"📝 **Description / Justification:**\n{req.get('desc','')}")
                        display_attachments(req)
                    with col_right:
                        st.markdown("### ✍️ Decision")
                        comments = st.text_area("Director Comments", key=f"comm_{req_id}")
                        approve_btn = st.button("✅ APPROVE", type="primary", key=f"appr_{req_id}")
                        reject_btn = st.button("❌ REJECT", type="secondary", key=f"rejt_{req_id}")

                        if approve_btn:
                            records = load_records_from_excel()
                            for r in records:
                                if int(r.get("id",0)) == int(req_id):
                                    r["status"] = "approved"
                                    r["approved_by"] = full_name
                                    r["decision_by"] = full_name
                                    r["director_comments"] = comments
                                    r["decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    r["approved_date"] = r["decision_date"]
                                    break
                            save_all_records(records)
                            log_action("APPROVED", req_id)
                            st.success(f"✅ Request #{req_id} APPROVED.")
                            st.rerun()

                        if reject_btn:
                            records = load_records_from_excel()
                            for r in records:
                                if int(r.get("id",0)) == int(req_id):
                                    r["status"] = "rejected"
                                    r["decision_by"] = full_name
                                    r["director_comments"] = comments
                                    r["decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    break
                            save_all_records(records)
                            log_action("REJECTED", req_id)
                            st.error(f"❌ Request #{req_id} REJECTED.")
                            st.rerun()

    with tab_approved:
        approved = [r for r in all_live_requests if r.get("status") == "approved"]
        if not approved:
            st.info("📋 No approved requests yet.")
        else:
            st.metric("✅ Approved Requests", len(approved)); st.divider()
            for req in reversed(approved):
                req_id = req.get("id")
                dec_by = req.get('decision_by', 'Director')
                dec_date = format_date(req.get('decision_date',''))
                with st.expander(f"🟢 ID #{req_id} | {req.get('emp_name')} | £{float(req.get('amount',0)):.2f} | ✅ {dec_by} — {dec_date}"):
                    st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept','')}")
                    st.write(f"💷 Amount: £{float(req.get('amount',0)):.2f}")
                    st.write(f"🎯 Approved By: {dec_by}")
                    st.write(f"📅 Approval Date: {dec_date}")
                    st.info(f"💬 Comments: {req.get('director_comments', 'None')}")
                    display_attachments(req)
                    st.divider()
                    display_pdf_button(req, can_generate=True)

    with tab_rejected:
        rejected = [r for r in all_live_requests if r.get("status") == "rejected"]
        if not rejected:
            st.success("✅ No rejected requests!")
        else:
            st.metric("❌ Rejected Requests", len(rejected)); st.divider()
            for req in reversed(rejected):
                req_id = req.get("id")
                dec_by = req.get('decision_by', 'Director')
                dec_date = format_date(req.get('decision_date',''))
                with st.expander(f"🔴 ID #{req_id} | {req.get('emp_name')} | £{float(req.get('amount',0)):.2f} | ❌ {dec_by} — {dec_date}"):
                    st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept','')}")
                    st.write(f"💷 Amount: £{float(req.get('amount',0)):.2f}")
                    st.error(f"❌ Rejected By: {dec_by} on {dec_date}")
                    st.error(f"💬 Reason: {req.get('director_comments', 'None')}")
                    display_attachments(req)
                    

# ─── 4️⃣ SUPER ADMIN PORTAL ───
elif role == "Super Admin":
    st.subheader("🛡️ Super Admin — All Requests")
    st.info("✅ View ALL requests across ALL departments. Download PDFs. **Approval → Director only.**")
    st.divider()
    tab_pending, tab_approved, tab_rejected, tab_manage = st.tabs([
        "⏳ All Pending", "✅ All Approved", "❌ All Rejected", "🔧 System Management"
    ])

    with tab_pending:
        pending = [r for r in all_live_requests if str(r.get("status", "")).strip().lower() == "pending"]
        if not pending:
            st.success("✅ No pending requests.")
        else:
            st.metric("⏳ All Pending", len(pending)); st.divider()
            for req in reversed(pending):
                req_id = req.get("id")
                amount = float(req.get("amount", 0))
                title = f"🟡 ID #{req_id} | {req.get('emp_name')} | {req.get('dept')} | £{amount:.2f}"
                with st.expander(title):
                    st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept')}")
                    st.write(f"🔄 Type: {req.get('type')} | 🏷️ Category: {req.get('category')}")
                    st.write(f"💷 Amount: £{amount:.2f}")
                    st.write(f"👔 Line Manager: {req.get('manager')} | 📅 Date: {format_date(req.get('date',''))}")
                    st.info(f"📝 Description: {req.get('desc')}")
                    display_attachments(req)
                    if req.get("director_comments"):
                        st.info(f"💬 Director Comments: {req.get('director_comments')}")
                    st.divider()
                    display_pdf_button(req, can_generate=True)

    with tab_approved:
        approved = [r for r in all_live_requests if str(r.get("status", "")).strip().lower() == "approved"]
        if not approved:
            st.info("📋 No approved requests.")
        else:
            st.metric("✅ All Approved", len(approved)); st.divider()
            for req in reversed(approved):
                req_id = req.get("id")
                amount = float(req.get("amount", 0))
                dec_by = req.get('decision_by', 'Director')
                dec_date = req.get('decision_date', '')
                display_date = dec_date[:10] if dec_date and len(dec_date) >= 10 else ""
                extra_text = f" | ✅ Approved by {dec_by} on {display_date}" if display_date else f" | ✅ Approved by {dec_by}"
                title = f"🟢 ID #{req_id} | {req.get('emp_name')} | {req.get('dept')} | £{amount:.2f}{extra_text}"
                with st.expander(title):
                    st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept')}")
                    st.write(f"💷 Amount: £{amount:.2f}")
                    st.write(f"🎯 **Approved By:** {dec_by}")
                    if display_date:
                        st.write(f"📅 **Approval Date:** {display_date}")
                    st.success(f"💬 Director Comments: {req.get('director_comments', 'None')}")
                    display_attachments(req)
                    st.divider()
                    display_pdf_button(req, can_generate=True)

    with tab_rejected:
        rejected = [r for r in all_live_requests if str(r.get("status", "")).strip().lower() == "rejected"]
        if not rejected:
            st.info("📋 No rejected requests.")
        else:
            st.metric("❌ All Rejected", len(rejected)); st.divider()
            for req in reversed(rejected):
                req_id = req.get("id")
                amount = float(req.get("amount", 0))
                dec_by = req.get('decision_by', 'Director')
                dec_date = format_date(req.get('decision_date',''))
                title = f"🔴 ID #{req_id} | {req.get('emp_name')} | {req.get('dept')} | £{amount:.2f}"
                with st.expander(title):
                    st.write(f"👤 Employee: {req.get('emp_name')} | 🏢 Department: {req.get('dept')}")
                    st.write(f"💷 Amount: £{amount:.2f}")
                    st.error(f"❌ Rejected By: {dec_by} on {dec_date}")
                    st.error(f"💬 Reason: {req.get('director_comments', 'None')}")
                    display_attachments(req)
                    st.divider()
                    display_pdf_button(req, can_generate=True)

    with tab_manage:
        tab_settings, tab_users, tab_audit = st.tabs([
            "⚙️ System Settings", "👤 User Management", "📖 Audit History"
        ])
        with tab_settings:
            settings_management_panel()
        with tab_users:
            user_management_panel()
        with tab_audit:
            if "display_audit_log_panel" in globals():
                display_audit_log_panel()
            else:
                st.info("📖 Audit log panel not defined — skipping")
        st.divider()
        st.subheader("📥 Download Data Backups")
        backup_col1, backup_col2, backup_col3 = st.columns(3)
        with backup_col1:
            if "EXCEL_PATH" in globals() and os.path.exists(EXCEL_PATH):
                with open(EXCEL_PATH, "rb") as f:
                    st.download_button(
                        "📥 Download Requests",
                        f.read(),
                        file_name=f"BACKUP_requests_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                        type="primary"
                    )
        with backup_col2:
            if "USER_DB_PATH" in globals() and os.path.exists(USER_DB_PATH):
                with open(USER_DB_PATH, "rb") as f:
                    st.download_button(
                        "📥 Download Users",
                        f.read(),
                        file_name=f"BACKUP_users_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                        type="primary"
                    )
        with backup_col3:
            if "SETTINGS_PATH" in globals() and os.path.exists(SETTINGS_PATH):
                with open(SETTINGS_PATH, "rb") as f:
                    st.download_button(
                        "📥 Download Settings",
                        f.read(),
                        file_name=f"BACKUP_settings_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                        type="primary"
                    )
        st.caption("💾 Save these files to your computer for backup")

# ─── ✅ DEFAULT / FALLBACK (MUST BE LAST!) ───
else:
    st.subheader("🔐 Access Restricted")
    st.error("❌ Your role does not have a defined portal. Please contact Super Admin.")

# ========================================================
# ✅ END OF ROLE-BASED PORTALS
# ========================================================
# Auto-save to GitHub after every page load
#github_auto_save()
# ============================================================
# ✅ END OF FILE — NOTHING AFTER THIS!
# ============================================================
