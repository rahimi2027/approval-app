# ============================================================
# 🔄 ACOOLE PORTAL — PROFESSIONAL VERSION v2.4 (ALL BUGS FIXED)
# ============================================================
# ✅ SyntaxError FIXED: functions defined BEFORE role chain
# ✅ NameError FIXED: all functions exist before being called
# ✅ Role chain intact
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
# ✅ ALL CONSTANTS DEFINED FIRST
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
APP_FOLDER = os.path.join(BASE_DIR, "Acoole_App_Uploads")
UPLOAD_DIR = os.path.join(APP_FOLDER, "uploaded_attachments")
PDF_DIR = os.path.join(APP_FOLDER, "approved_pdfs")
ARCHIVE_FOLDER = os.path.join(APP_FOLDER, "audit_archives")
EXCEL_PATH = os.path.join(APP_FOLDER, "requests.xlsx")
USER_DB_PATH = os.path.join(APP_FOLDER, "user_database.xlsx")
SETTINGS_PATH = os.path.join(APP_FOLDER, "settings.xlsx")
AUDIT_LOG_PATH = os.path.join(APP_FOLDER, "audit_log.xlsx")
AUDIT_LOG_FILE = AUDIT_LOG_PATH

GOOGLE_DRIVE_FOLDER_ID = ""
USE_ONEDRIVE = False
ONEDRIVE_CLIENT_ID = ""
ONEDRIVE_CLIENT_SECRET = ""
ONEDRIVE_TENANT_ID = ""
ONEDRIVE_FOLDER = "Approved Requests/"

LOGO_PATH = os.path.join(BASE_DIR, "logo.png")
APPROVED_STAMP_PATH = os.path.join(BASE_DIR, "approved.png")
REJECTED_STAMP_PATH = os.path.join(BASE_DIR, "rejected.png")

AUDIT_COLUMNS = [
    "AuditID", "Timestamp", "User_Name", "User_Role",
    "Action", "Request_ID", "Department", "Amount",
    "Decision_By", "Decision_Date", "Field_Changed",
    "Old_Value", "New_Value", "IP_Address"
]
# ============================================================
# ✅ AUTO-CREATE FOLDERS & FILES
# ============================================================
os.makedirs(APP_FOLDER, exist_ok=True)
for folder in [UPLOAD_DIR, PDF_DIR, ARCHIVE_FOLDER]:
    os.makedirs(folder, exist_ok=True)

def create_empty_if_missing(path, columns):
    if not os.path.exists(path):
        pd.DataFrame(columns=columns).to_excel(path, index=False, engine="openpyxl")

create_empty_if_missing(EXCEL_PATH, [
    "ID", "Employee Name", "Department", "Transaction Type", "Category Reason",
    "Date", "Amount (£)", "Line Manager", "Description", "Attachment Name",
    "Status", "Director Comments", "Decision Date", "Decision By",
    "PDF File Path", "Edited From ID", "Old Data"
])
create_empty_if_missing(USER_DB_PATH, [
    "full_name","username","password","role","dept",
    "can_view_all_dept","can_generate_pdf","can_download_data","can_approve_requests"
])
create_empty_if_missing(SETTINGS_PATH, ["setting", "value"])
create_empty_if_missing(AUDIT_LOG_PATH, AUDIT_COLUMNS)

# ============================================================
# ✅ PAGE CONFIG
# ============================================================
st.markdown("""
    <style>
    .block-container { padding-top: 2rem !important; padding-left: 0.3rem !important; padding-right: 2rem !important; max-width: 1400px !important; width: 90% !important; }
    section[data-testid="stSidebar"] { width: 320px !important; }
    section[data-testid="stSidebar"] > div:first-child > div { padding-left: 0.2rem !important; padding-right: 0.2rem !important; }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# ✅ GOOGLE / ONEDRIVE FUNCTIONS
# ============================================================
SERVICE_ACCOUNT_INFO = {}
KEY_FILE = os.path.join(BASE_DIR, "service_account_key.json")
if os.path.exists(KEY_FILE):
    with open(KEY_FILE, "r", encoding="utf-8") as f:
        SERVICE_ACCOUNT_INFO = json.load(f)

def get_drive_service():
    if not SERVICE_ACCOUNT_INFO or not GOOGLE_DRIVE_FOLDER_ID: return None
    try:
        credentials = service_account.Credentials.from_service_account_info(
            SERVICE_ACCOUNT_INFO, scopes=["https://www.googleapis.com/auth/drive"])
        return build("drive", "v3", credentials=credentials)
    except: return None

def upload_to_google_drive(local_file_path, display_filename):
    if not SERVICE_ACCOUNT_INFO or not GOOGLE_DRIVE_FOLDER_ID: return None
    try:
        credentials = service_account.Credentials.from_service_account_info(
            SERVICE_ACCOUNT_INFO, scopes=["https://www.googleapis.com/auth/drive"])
        service = build("drive", "v3", credentials=credentials)
        file_metadata = {"name": display_filename, "parents": [GOOGLE_DRIVE_FOLDER_ID]}
        media = MediaFileUpload(local_file_path, resumable=True)
        file = service.files().create(body=file_metadata, media_body=media, fields="id").execute()
        return file.get("id")
    except: return None

def get_onedrive_token():
    if not USE_ONEDRIVE: return None
    try:
        url = f"https://login.microsoftonline.com/{ONEDRIVE_TENANT_ID}/oauth2/v2.0/token"
        data = {"grant_type": "client_credentials", "client_id": ONEDRIVE_CLIENT_ID,
                "client_secret": ONEDRIVE_CLIENT_SECRET, "scope": "https://graph.microsoft.com/.default"}
        res = requests.post(url, data=data, timeout=30)
        return res.json().get("access_token") if res.status_code == 200 else None
    except: return None

def upload_to_onedrive(local_file_path, remote_filename=None):
    if not USE_ONEDRIVE: return False
    token = get_onedrive_token()
    if not token: return False
    filename = remote_filename or os.path.basename(local_file_path)
    remote_path = f"{ONEDRIVE_FOLDER}{filename}"
    try:
        with open(local_file_path, 'rb') as f: content = f.read()
        url = f"https://graph.microsoft.com/v1.0/drives/me/items/root:/{remote_path}:/content"
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/octet-stream"}
        res = requests.put(url, data=content, headers=headers, timeout=60)
        return res.status_code in (200, 201)
    except: return False

# ============================================================
# ✅ DEFAULTS
# ============================================================
DEFAULT_CATEGORIES = ["Food Allowance", "Others", "Parking", "Parking Fine", "GYM Membership", "Item Not Returned", "Item Missing"]
DEFAULT_ROLES = ["Manager", "Staff", "Team Member", "Director", "Payroll", "Super Admin"]
DEFAULT_DEPARTMENTS = ["National Grid", "Isolator", "Project", "Accounts", "Payroll Department", "ACoole Electrical Ltd"]
EXCEL_COLUMNS = ["ID", "Employee Name", "Department", "Transaction Type", "Category Reason",
    "Date", "Amount (£)", "Line Manager", "Description", "Attachment Name",
    "Status", "Director Comments", "Decision Date", "Decision By",
    "PDF File Path", "Edited From ID", "Old Data"]
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
# ✅ PDF LIBRARY
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
# ✅ AUDIT, SETTINGS, USER, EXCEL HELPERS
# ============================================================
def safe_init_excel(path, columns):
    if not os.path.exists(path):
        pd.DataFrame(columns=columns).to_excel(path, index=False, engine="openpyxl"); return True
    try: pd.read_excel(path, engine="openpyxl"); return True
    except Exception: os.remove(path); pd.DataFrame(columns=columns).to_excel(path, index=False, engine="openpyxl"); return True

def init_audit_log(): safe_init_excel(AUDIT_LOG_PATH, AUDIT_COLUMNS)
def load_audit_log():
    init_audit_log()
    try: return pd.read_excel(AUDIT_LOG_PATH, engine="openpyxl").fillna("").to_dict(orient="records")
    except: return []
def save_audit_entry(entry):
    init_audit_log()
    try:
        df = pd.read_excel(AUDIT_LOG_PATH, engine="openpyxl").fillna("")
        df = pd.concat([df, pd.DataFrame([entry])], ignore_index=True)
        df.to_excel(AUDIT_LOG_PATH, index=False, engine="openpyxl")
    except: pass

def init_settings():
    if not os.path.exists(SETTINGS_PATH):
        pd.DataFrame([{"setting": "categories", "value": "|".join(DEFAULT_CATEGORIES)},
            {"setting": "roles", "value": "|".join(DEFAULT_ROLES)},
            {"setting": "departments", "value": "|".join(DEFAULT_DEPARTMENTS)}]).to_excel(SETTINGS_PATH, index=False, engine="openpyxl")

def load_departments():
    init_settings()
    try:
        df = pd.read_excel(SETTINGS_PATH, engine="openpyxl").fillna("")
        row = df[df["setting"] == "departments"]
        if len(row): return [v.strip() for v in row.iloc[0]["value"].split("|") if v.strip()]
    except: pass
    return DEFAULT_DEPARTMENTS.copy()

def save_departments(dept_list):
    init_settings()
    df = pd.read_excel(SETTINGS_PATH, engine="openpyxl").fillna("")
    if "departments" in list(df["setting"]):
        df.loc[df["setting"]=="departments", "value"] = "|".join(dept_list)
    else:
        df = pd.concat([df, pd.DataFrame([{"setting": "departments", "value": "|".join(dept_list)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False, engine="openpyxl")

def load_categories():
    init_settings()
    try:
        df = pd.read_excel(SETTINGS_PATH, engine="openpyxl").fillna("")
        row = df[df["setting"] == "categories"]
        if len(row): return [v.strip() for v in row.iloc[0]["value"].split("|") if v.strip()]
    except: pass
    return DEFAULT_CATEGORIES

def save_categories(cat_list):
    init_settings()
    df = pd.read_excel(SETTINGS_PATH, engine="openpyxl").fillna("")
    if "categories" in list(df["setting"]):
        df.loc[df["setting"]=="categories", "value"] = "|".join(cat_list)
    else:
        df = pd.concat([df, pd.DataFrame([{"setting": "categories", "value": "|".join(cat_list)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False, engine="openpyxl")

def load_roles():
    init_settings()
    try:
        df = pd.read_excel(SETTINGS_PATH, engine="openpyxl").fillna("")
        row = df[df["setting"] == "roles"]
        if len(row): return [v.strip() for v in row.iloc[0]["value"].split("|") if v.strip()]
    except: pass
    return DEFAULT_ROLES

def save_roles(roles_list):
    init_settings()
    df = pd.read_excel(SETTINGS_PATH, engine="openpyxl").fillna("")
    if "roles" in list(df["setting"]):
        df.loc[df["setting"]=="roles", "value"] = "|".join(roles_list)
    else:
        df = pd.concat([df, pd.DataFrame([{"setting": "roles", "value": "|".join(roles_list)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False, engine="openpyxl")

def init_user_db():
    safe_init_excel(USER_DB_PATH, ["full_name","username","password","role","dept",
        "can_view_all_dept","can_generate_pdf","can_download_data","can_approve_requests"])
    if os.path.getsize(USER_DB_PATH) < 500:
        pd.DataFrame(DEFAULT_USERS).to_excel(USER_DB_PATH, index=False, engine="openpyxl")

def save_users(users_dict):
    rows = []
    for u, d in users_dict.items():
        rows.append({"full_name":d.get("full_name",u),"username":u,"password":d.get("password",""),
            "role":d.get("role","Staff"),"dept":d.get("dept",""),
            "can_view_all_dept":d.get("can_view_all_dept",False),"can_generate_pdf":d.get("can_generate_pdf",False),
            "can_download_data":d.get("can_download_data",False),"can_approve_requests":d.get("can_approve_requests",False)})
    pd.DataFrame(rows).to_excel(USER_DB_PATH, index=False, engine="openpyxl")

def load_users():
    init_user_db()
    try:
        df = pd.read_excel(USER_DB_PATH, engine="openpyxl").fillna("")
        u = {}
        for _,r in df.iterrows():
            u[r["username"]] = {"full_name":str(r.get("full_name",r["username"])).strip(),"password":str(r["password"]),
                "role":str(r.get("role","Staff")),"dept":str(r.get("dept","")),
                "can_view_all_dept":str(r.get("can_view_all_dept","False")).lower()=="true",
                "can_generate_pdf":str(r.get("can_generate_pdf","False")).lower()=="true",
                "can_download_data":str(r.get("can_download_data","False")).lower()=="true",
                "can_approve_requests":str(r.get("can_approve_requests","False")).lower()=="true"}
        return u
    except: return {}

def load_records_from_excel():
    try:
        if not os.path.exists(EXCEL_PATH): return []
        df = pd.read_excel(EXCEL_PATH, engine="openpyxl").fillna("")
        if df.empty: return []
        recs = df.to_dict(orient="records")
        out = []
        for r in recs:
            try: rid = int(r.get("ID",0))
            except: rid = 0
            try: amt = float(r.get("Amount (£)",0))
            except: amt = 0.0
            out.append({"id":rid,"emp_name":str(r.get("Employee Name","")).strip(),"dept":str(r.get("Department","")).strip(),
                "type":str(r.get("Transaction Type","")).strip(),"category":str(r.get("Category Reason","")).strip(),
                "date":str(r.get("Date","")).strip(),"amount":amt,"manager":str(r.get("Line Manager","")).strip(),
                "desc":str(r.get("Description","")).strip(),"attachment_name":str(r.get("Attachment Name","None")).strip(),
                "status":str(r.get("Status","pending")).strip().lower(),"director_comments":str(r.get("Director Comments","")).strip(),
                "decision_date":str(r.get("Decision Date","")).strip(),"decision_by":str(r.get("Decision By","")).strip(),
                "pdf_path":str(r.get("PDF File Path","")).strip(),"edited_from_id":str(r.get("Edited From ID","")).strip(),
                "old_data":str(r.get("Old Data","")).strip()})
        return out
    except: return []

def save_all_records(records):
    out = []
    for r in records:
        out.append({"ID":int(r.get("id",0)),"Employee Name":str(r.get("emp_name","")),
            "Department":str(r.get("dept","")),"Transaction Type":str(r.get("type","")),
            "Category Reason":str(r.get("category","")),"Date":str(r.get("date","")),
            "Amount (£)":float(r.get("amount",0.0)),"Line Manager":str(r.get("manager","")),
            "Description":str(r.get("desc","")),"Attachment Name":str(r.get("attachment_name","None")),
            "Status":str(r.get("status","pending")).lower(),"Director Comments":str(r.get("director_comments","")),
            "Decision Date":str(r.get("decision_date","")),"Decision By":str(r.get("decision_by","")),
            "PDF File Path":str(r.get("pdf_path","")),"Edited From ID":str(r.get("edited_from_id","")),
            "Old Data":str(r.get("old_data",""))})
    pd.DataFrame(out, columns=EXCEL_COLUMNS).to_excel(EXCEL_PATH, index=False, engine="openpyxl")

def save_record_to_excel(new_rec):
    recs = load_records_from_excel()
    recs.append(new_rec)
    save_all_records(recs)

# ============================================================
# ✅ PDF GENERATION FUNCTION
# ============================================================
def generate_approval_pdf(request_data):
    if not PDF_AVAILABLE: return False, None, "Install fpdf2"
    try:
        def clean(t): return str(t).replace("\u2013","-").encode("latin-1","ignore").decode("latin-1").strip() if t else ""
        rd = request_data
        pdf = FPDF(); pdf.add_page()
        if os.path.exists(LOGO_PATH): pdf.image(LOGO_PATH, x=75, y=10, w=60)
        pdf.ln(25); pdf.set_font("Courier","B",11); pdf.cell(0,6,"Approval Form",ln=True,align="C"); pdf.ln(5)
        pdf.set_font("Courier","",9)
        pdf.cell(50,6,"Request ID:"); pdf.cell(0,6,str(rd.get("id","")),ln=True)
        pdf.cell(50,6,"Employee:"); pdf.cell(0,6,clean(rd.get("emp_name","")),ln=True)
        pdf.cell(50,6,"Dept:"); pdf.cell(0,6,clean(rd.get("dept","")),ln=True)
        pdf.cell(50,6,"Amount:"); pdf.cell(0,6,f"£{rd.get('amount',0):.2f}",ln=True)
        pdf.cell(50,6,"Status:"); pdf.cell(0,6,rd.get("status","").upper(),ln=True)
        if rd.get("decision_by"): pdf.cell(50,6,"Approved By:"); pdf.cell(0,6,clean(rd.get("decision_by","")),ln=True)
        pdf.ln(8); pdf.set_font("Courier","B",10); pdf.cell(0,6,"Description:",ln=True); pdf.set_font("Courier","",9)
        pdf.multi_cell(0,6,clean(rd.get("desc","")))
        fn = f"Request_{rd.get('id','')}_{datetime.now().strftime('%Y%m%d')}.pdf"
        path = os.path.join(PDF_DIR, fn)
        os.makedirs(PDF_DIR, exist_ok=True)
        with open(path, "wb") as f: f.write(bytes(pdf.output()))
        return True, bytes(pdf.output()), fn
    except Exception as e: return False, None, str(e)

def display_pdf_button(req, can_generate=False, key_suffix=""):
    uid = f"pdf_{req.get('id','')}_{key_suffix}"
    if can_generate and PDF_AVAILABLE:
        if st.button(f"📄 Generate PDF #{req.get('id')}", key=uid):
            ok, data, name = generate_approval_pdf(req)
            if ok:
                st.download_button(f"📥 Download {name}", data=data, file_name=name, mime="application/pdf", type="primary")
            else: st.error(f"❌ {name}")

# ============================================================
# ✅ CRITICAL: SUPER ADMIN FUNCTIONS DEFINED FIRST
# ============================================================
def show_dashboard(user, all_live_requests):
    role = user.get("role",""); dept = user.get("dept",""); name = user.get("full_name",user.get("username","User"))
    visible = all_live_requests if role in ["Director","Payroll","Super Admin"] else [r for r in all_live_requests if r.get("dept","")==dept]
    pending = [r for r in visible if r.get("status")=="pending"]
    approved = [r for r in visible if r.get("status")=="approved"]
    total_amt = sum(r.get("amount",0) for r in approved)
    st.subheader(f"👋 Welcome, {name}")
    st.markdown(f"**Role:** {role} | **Dept:** {dept} | 📅 {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    st.divider()
    c1,c2,c3,c4 = st.columns(4)
    with c1: st.metric("⏳ Pending", len(pending))
    with c2: st.metric("✅ Approved", len(approved))
    with c3: st.metric("❌ Rejected", len(visible)-len(pending)-len(approved))
    with c4: st.metric("💰 Total Approved", f"£{total_amt:.2f}")
    st.divider()

def user_management_panel():
    st.subheader("👤 User Management")
    st.info("🛡️ Create, edit, delete user accounts & set permissions.")
    st.divider()
    USERS = load_users(); ROLES = load_roles(); DEPTS = load_departments()
    t1,t2,t3 = st.tabs(["➕ Create","✏️ Edit","🗑️ Delete"])
    with t1:
        with st.form("create_user", border=True):
            fn = st.text_input("Full Name"); un = st.text_input("Username").lower().strip()
            pw = st.text_input("Password", type="password"); rl = st.selectbox("Role", ROLES)
            dt = st.selectbox("Department", DEPTS)
            defs = PERMISSION_DEFAULTS[rl]
            c1,c2 = st.columns(2)
            vw = c1.checkbox("View All Depts", value=defs["can_view_all_dept"])
            pdf = c1.checkbox("Generate PDFs", value=defs["can_generate_pdf"])
            dl = c2.checkbox("Download Data", value=defs["can_download_data"])
            ap = c2.checkbox("Approve Requests", value=defs["can_approve_requests"])
            if st.form_submit_button("✅ Create User", type="primary"):
                if not fn.strip() or not un or not pw: st.error("All fields required")
                elif un in USERS: st.error("Username exists")
                else:
                    USERS[un] = {"full_name":fn.strip(),"password":pw,"role":rl,"dept":dt,
                        "can_view_all_dept":vw,"can_generate_pdf":pdf,"can_download_data":dl,"can_approve_requests":ap}
                    save_users(USERS); st.success(f"✅ Created: {fn}"); st.rerun()
    with t2:
        sel = st.selectbox("Select User to Edit", list(USERS.keys()))
        if sel:
            u = USERS[sel]
            with st.form("edit_user", border=True):
                fn = st.text_input("Full Name", value=u["full_name"])
                unew = st.text_input("Username", value=sel).lower().strip()
                pw = st.text_input("New Password (leave blank to keep)", type="password")
                rl = st.selectbox("Role", ROLES, index=ROLES.index(u["role"]) if u["role"] in ROLES else 0)
                dt = st.selectbox("Department", DEPTS, index=DEPTS.index(u["dept"]) if u["dept"] in DEPTS else 0)
                vw = st.checkbox("View All Depts", value=u["can_view_all_dept"])
                pdf = st.checkbox("Generate PDFs", value=u["can_generate_pdf"])
                dl = st.checkbox("Download Data", value=u["can_download_data"])
                ap = st.checkbox("Approve Requests", value=u["can_approve_requests"])
                if st.form_submit_button("🔄 Update", type="primary"):
                    USERS = load_users()
                    if unew != sel:
                        USERS[unew] = {"full_name":fn.strip(),"password":pw if pw else u["password"],"role":rl,"dept":dt,
                            "can_view_all_dept":vw,"can_generate_pdf":pdf,"can_download_data":dl,"can_approve_requests":ap}
                        del USERS[sel]
                    else:
                        USERS[sel]["full_name"]=fn.strip()
                        if pw: USERS[sel]["password"]=pw
                        USERS[sel].update({"role":rl,"dept":dt,"can_view_all_dept":vw,"can_generate_pdf":pdf,"can_download_data":dl,"can_approve_requests":ap})
                    save_users(USERS); st.success("✅ Updated"); st.rerun()
    with t3:
        sel = st.selectbox("Select User to DELETE", [u for u in USERS if u != st.session_state.user_info.get("username")])
        if sel and st.button(f"🗑️ Delete {USERS[sel]['full_name']}", type="secondary"):
            del USERS[sel]; save_users(USERS); st.success("✅ Deleted"); st.rerun()

def settings_management_panel():
    st.subheader("⚙️ System Settings")
    st.info("🛡️ Manage categories, departments, roles."); st.divider()
    cats_tab, dept_tab, roles_tab = st.tabs(["🏷️ Categories","🏢 Departments","🎖️ Roles"])
    with cats_tab:
        cats = load_categories()
        with st.form("add_cat"):
            nc = st.text_input("New Category")
            if st.form_submit_button("✅ Add") and nc.strip() and nc.strip() not in cats:
                cats.append(nc.strip()); save_categories(cats); st.success("Added"); st.rerun()
        for i,c in enumerate(cats):
            if i>0:
                col1,col2,col3 = st.columns([4,1,1])
                col1.write(f"• {c}")
                if col3.button("🗑️", key=f"dc{i}"): cats.pop(i); save_categories(cats); st.rerun()
    with dept_tab:
        depts = load_departments()
        with st.form("add_dept"):
            nd = st.text_input("New Department")
            if st.form_submit_button("✅ Add") and nd.strip() and nd.strip() not in depts:
                depts.append(nd.strip()); save_departments(depts); st.success("Added"); st.rerun()
        for i,d in enumerate(depts):
            if i>0:
                col1,_,col3 = st.columns([4,1,1])
                col1.write(f"• {d}")
                if col3.button("🗑️", key=f"dd{i}"): depts.pop(i); save_departments(depts); st.rerun()
    with roles_tab:
        roles = load_roles()
        with st.form("add_role"):
            nr = st.text_input("New Role")
            if st.form_submit_button("✅ Add") and nr.strip() and nr.strip() not in roles:
                roles.append(nr.strip()); save_roles(roles); st.success("Added"); st.rerun()
        for i,r in enumerate(roles):
            if r!="Super Admin" and i>0:
                col1,_,col3 = st.columns([4,1,1])
                col1.write(f"• {r}")
                if col3.button("🗑️", key=f"dr{i}"): roles.pop(i); save_roles(roles); st.rerun()

def display_audit_log_panel():
    st.subheader("📜 Full Audit Log")
    logs = load_audit_log()
    if not logs: st.info("Empty"); return
    st.dataframe(pd.DataFrame(logs).sort_values("Timestamp", ascending=False), use_container_width=True)
    st.download_button("📥 Download CSV", pd.DataFrame(logs).to_csv(index=False).encode("utf-8"), "audit_log.csv", type="primary")

# ============================================================
# ✅ SESSION & LOGIN
# ============================================================
for k in ["logged_in","user_info","editing_request_id"]:
    if k not in st.session_state: st.session_state[k] = None if k=="editing_request_id" else (False if k=="logged_in" else {})

def display_company_header():
    if os.path.exists(LOGO_PATH): st.image(LOGO_PATH, width=300)
    else: st.title("⚡ ACOOLE ELECTRICAL LTD")
    st.caption("Approval Platform"); st.divider()

if not st.session_state.logged_in:
    display_company_header()
    with st.form("login", border=True):
        st.subheader("🔒 Login")
        u = st.text_input("Username").lower().strip()
        p = st.text_input("Password", type="password")
        if st.form_submit_button("🔐 Sign In", type="primary"):
            USERS = load_users()
            if u in USERS and USERS[u]["password"]==p:
                st.session_state.logged_in=True; st.session_state.user_info={**USERS[u],"username":u}; st.rerun()
            else: st.error("Invalid credentials")
    st.stop()

# ============================================================
# ✅ TOP BAR
# ============================================================
lft,rgt = st.columns([4,1])
with lft:
    if st.button("🔄 Refresh"): st.rerun()
with rgt:
    if st.button("🔒 Logout"): st.session_state.clear(); st.rerun()

display_company_header()
ui = st.session_state.user_info
full_name = ui.get("full_name",ui.get("username","User"))
dept = ui.get("dept",""); role = ui.get("role","")
st.info(f"👤 {full_name} | {dept} | {role}")
all_live_requests = load_records_from_excel()
CATEGORIES = load_categories()

# ============================================================
# ✅ ROLE CHAIN — NO CODE/FUNCTIONS IN BETWEEN!
# ============================================================

# ─── PAYROLL ───
if role == "Payroll":
    st.subheader("🧾 Payroll Portal")
    st.info("✅ View & download all approved requests"); st.divider()
    approved = [r for r in all_live_requests if r.get("status")=="approved"]
    if not approved: st.info("No approved requests yet.")
    else:
        st.metric("✅ Approved", len(approved)); st.divider()
        for r in reversed(approved):
            with st.expander(f"🟢 ID #{r['id']} | {r['emp_name']} | £{r['amount']:.2f}"):
                st.write(f"Dept: {r['dept']} | Date: {r['date']} | Approved by: {r.get('decision_by','—')}")
                display_pdf_button(r, can_generate=True)

# ─── MANAGER / STAFF ───
elif role in ["Manager","Staff","Team Member"]:
    dept_name = dept
    st.subheader(f"➕ New Request — {dept_name}")
    with st.form("new_req", clear_on_submit=True):
        c1,c2 = st.columns(2)
        with c1:
            en = st.text_input("Employee Name"); rt = st.selectbox("Type",["Addition","Deduction"])
            ct = st.selectbox("Category", CATEGORIES); amt = st.number_input("Amount (£)",0.01,step=10.0)
        with c2:
            from datetime import datetime as dt
            dv = st.date_input("Date"); mg = st.text_input("Line Manager"); desc = st.text_area("Description")
        if st.form_submit_button("📤 Submit", type="primary"):
            nid = max([x.get("id",0) for x in all_live_requests]+[0])+1
            payload = {"id":nid,"emp_name":en.strip(),"dept":dept_name,"type":rt,"category":ct,
                "date":str(dv),"amount":amt,"manager":mg.strip(),"desc":desc.strip(),
                "attachment_name":"None","status":"pending","director_comments":"",
                "decision_date":"","decision_by":"","pdf_path":"","edited_from_id":"","old_data":""}
            save_record_to_excel(payload); st.success(f"✅ Request #{nid} submitted!"); st.rerun()
    st.divider()
    my = [r for r in all_live_requests if r.get("dept")==dept_name]
    pend = [r for r in my if r.get("status")=="pending"]
    if pend:
        st.subheader("⏳ Pending"); st.divider()
        for r in reversed(pend): st.expander(f"🟡 #{r['id']} | {r['emp_name']} | £{r['amount']:.2f}").write(f"Desc: {r['desc']}")

# ─── DIRECTOR ───
elif role == "Director":
    st.subheader("🎬 Director Approval Portal")
    pend = [r for r in all_live_requests if r.get("status")=="pending"]
    appd = [r for r in all_live_requests if r.get("status")=="approved"]
    st.info(f"⏳ Pending: {len(pend)} | ✅ Approved: {len(appd)}"); st.divider()
    for r in reversed(pend):
        with st.expander(f"🟡 ID #{r['id']} | {r['emp_name']} | {r['dept']} | £{r['amount']:.2f}"):
            st.write(f"Desc: {r['desc']} | Manager: {r['manager']}")
            cmt = st.text_area("Comments", key=f"cmt{r['id']}")
            ap,rej = st.columns(2)
            if ap.button("✅ APPROVE", type="primary", key=f"ap{r['id']}") or rej.button("❌ REJECT", key=f"rj{r['id']}"):
                new_status = "approved" if ap.button else "rejected"
                recs = load_records_from_excel()
                for x in recs:
                    if x["id"]==r["id"]:
                        x["status"]=new_status; x["decision_by"]=full_name
                        x["decision_date"]=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        x["director_comments"]=cmt or "None"; break
                save_all_records(recs); st.success(f"✅ Request #{r['id']} {new_status.upper()}"); st.rerun()

# ─── SUPER ADMIN ───
elif role == "Super Admin":
    st.subheader("🛡️ Super Admin Control Centre")
    st.info("✅ Full system access — Dashboard, Users, Settings, Audit Logs & Data")
    st.divider()
    show_dashboard(ui, all_live_requests)
    st.divider()
    audit_tab, users_tab, settings_tab, data_tab = st.tabs(["📜 Audit Logs","👤 Users","⚙️ Settings","📊 Export"])
    with audit_tab: display_audit_log_panel()
    with users_tab: user_management_panel()
    with settings_tab: settings_management_panel()
    with data_tab:
        st.markdown("### 📊 Export All Data")
        if st.button("📥 Export Full CSV", type="primary"):
            df = pd.DataFrame(all_live_requests)
            st.download_button("📄 Download requests.csv", df.to_csv(index=False).encode("utf-8"),
                               f"requests_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", "text/csv")

# ─── UNKNOWN ROLE ───
else:
    st.warning(f"⚠️ Role '{role}' has no portal. Contact Admin.")
    if st.button("🔒 Logout"): st.session_state.clear(); st.rerun()


# ============================================================
# ✅ HELPER FUNCTIONS — Place these BEFORE or keep at BOTTOM
# ============================================================
def get_next_id(all_records):
    """Return next available request ID integer."""
    if not all_records:
        return 1
    return max(int(r.get("id", 0)) for r in all_records) + 1


def format_date(date_str, fallback="—"):
    """Format date strings consistently in UI."""
    if not date_str or str(date_str).strip() in ["", "None"]:
        return fallback
    s = str(date_str).strip()
    if " " in s:
        d, t = s.split(" ", 1)
        return f"{d} ⏰ {t[:5]}"
    return s


def display_attachments(req):
    """Show attached files with download links."""
    att_str = str(req.get("attachment_name", "None"))
    if att_str.strip().lower() in ["none", ""]:
        return
    fnames = [n.strip() for n in att_str.split(",") if n.strip()]
    if not fnames:
        return
    st.markdown("📎 **Attachments:**")
    for fn in fnames:
        fpath = os.path.join(UPLOAD_DIR, fn)
        if os.path.exists(fpath):
            with open(fpath, "rb") as f:
                st.download_button(f"📄 {fn}", data=f.read(), file_name=fn,
                                   key=f"dl_att_{req.get('id')}_{fn.replace('.','_')}")
        else:
            st.caption(f"⚠️ `{fn}` — file missing")


def show_old_new_comparison(old_raw, req):
    """Show side-by-side diff for edited requests."""
    try:
        old = json.loads(old_raw) if isinstance(old_raw, str) else old_raw
    except:
        old = {}
    if not old:
        st.info("ℹ️ No previous version data.")
        return

    st.markdown("| Field | Previous Value | Updated Value |")
    st.markdown("|---|---|---|")
    fields = [
        ("emp_name", "Employee Name"), ("dept", "Department"),
        ("type", "Type"), ("category", "Category"),
        ("date", "Date"), ("amount", "Amount"),
        ("manager", "Manager"), ("desc", "Description")
    ]
    for key, label in fields:
        o = old.get(key, "—")
        n = req.get(key, "—")
        if key == "amount":
            try: o = f"£{float(o):,.2f}"
            except: pass
            try: n = f"£{float(n):,.2f}"
            except: pass
        if str(o) != str(n):
            st.markdown(f"| **{label}** | `{o}` | 🟢 **`{n}`** |")
        else:
            st.markdown(f"| {label} | `{o}` | `{n}` |")


def refresh_data_button():
    """Clear and reload data cache."""
    if st.button("🔄 Refresh Data", key="refresh_all_data"):
        st.cache_data.clear()
        st.success("✅ Data refreshed!")
        st.rerun()
# ========================================================
# ✅ END OF ROLE-BASED PORTALS
# ========================================================
# Auto-save to GitHub after every page load
#github_auto_save()
# ============================================================
# ✅ END OF FILE — NOTHING AFTER THIS!
# ============================================================
