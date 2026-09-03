# ============================================================
# 🔄 GITHUB AUTO-SAVE — KEEPS DATA PERMANENT & SYNCED
# ============================================================
import streamlit as st
import os
import subprocess
from datetime import datetime

def github_auto_save():
    """Push changed Excel files to GitHub automatically"""
    if not os.path.exists("/mount/src/"):
        return
    try:
        GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")
        GITHUB_REPO = st.secrets.get("GITHUB_REPO", "")
        GITHUB_BRANCH = st.secrets.get("GITHUB_BRANCH", "main")
        if not GITHUB_TOKEN or not GITHUB_REPO:
            print("⚠️ GitHub secrets not set — skipping auto-save")
            return
        repo_url = f"https://{GITHUB_TOKEN}@github.com/{GITHUB_REPO}.git"
        os.system("git config --global user.name 'Streamlit Auto-Save'")
        os.system("git config --global user.email 'rahimi2027@users.noreply.github.com'")
        data_files = ["requests.xlsx", "user_database.xlsx", "settings.xlsx"]
        for f in data_files:
            if os.path.exists(f):
                os.system(f"git add {f}")
        status = subprocess.run(["git", "status", "--porcelain"], capture_output=True, text=True)
        if status.stdout.strip():
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
            commit_msg = f"🔄 Auto-save: data updated {timestamp}"
            os.system(f'git commit -m "{commit_msg}"')
            os.system(f"git push {repo_url} {GITHUB_BRANCH}")
            st.toast("✅ Data saved & synced to GitHub!", icon="✅")
    except Exception as e:
        print(f"⚠️ GitHub save error: {e}")

# ─── DATE FORMATTING HELPER ──────────────────────────────────────
def format_date(d):
    """Format dates safely — returns first 10 chars or placeholder"""
    if not d or str(d).strip() == "" or str(d).strip().lower() == "none":
        return "—"
    return str(d).strip()[:10]

def display_attachments(req):
    import streamlit as st
    import os
    att = req.get("attachment_name", "None")
    if not att or str(att).strip() == "" or str(att).strip().lower() == "none":
        st.info("📎 No attachments.")
        return
    try:
        attached_files = [n.strip() for n in str(att).split(",")]
        for idx, name in enumerate(attached_files):
            path = os.path.join(UPLOAD_DIR, name)
            if os.path.exists(path):
                with open(path, "rb") as f:
                    st.download_button(
                        f"⬇️ Download {name}",
                        f.read(),
                        file_name=name,
                        key=f"att_{req.get('id', idx)}_{idx}"
                    )
            else:
                st.warning(f"⚠️ File not found: {name}")
    except Exception as e:
        st.info(f"📎 Attachments: {att}")

# ─── PDF BUTTON HELPER ✅ FIXED WITH UNIQUE KEYS ──────────────────
def display_pdf_button(req, can_generate=False, key_suffix=""):
    import streamlit as st
    req_id = req["id"]
    unique_key = f"genpdf_{req_id}_{key_suffix}"
    if can_generate and PDF_AVAILABLE:
        if st.button(f"📄 Generate PDF for ID #{req_id}", type="primary", key=unique_key):
            ok, pdf_bytes, name = generate_approval_pdf(req)
            if ok:
                st.success(f"✅ Generated! Ready to download ↓")
                st.download_button(
                    f"📥 Download: {name}",
                    data=pdf_bytes,
                    file_name=name,
                    mime="application/pdf",
                    type="primary",
                    key=f"dl_{unique_key}"
                )
            else:
                st.error(f"❌ {name}")
    return False

# ─── GET NEXT REQUEST ID ────────────────────────────────────────────
def get_next_id(all_records):
    if not all_records:
        return 1
    return max(int(r.get("id", 0)) for r in all_records) + 1

# ─── UPDATE RECORD STATUS ──────────────────────────────────────────
def update_record_status_in_excel(req_id, new_status, comments, approved_by):
    from datetime import datetime
    records = load_records_from_excel()
    decision_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for r in records:
        if int(r["id"]) == int(req_id):
            r["status"] = new_status
            r["decision_date"] = decision_datetime
            r["decision_by"] = approved_by
            if comments.strip():
                ts = datetime.now().strftime("%Y-%m-%d %H:%M")
                r["director_comments"] = f"[{ts}] {comments.strip()}"
            r["pdf_path"] = ""
            break
    save_all_records(records)

# ─── DELETE RECORD BY ID ────────────────────────────────────────────
def delete_record_by_id(req_id):
    records = load_records_from_excel()
    records = [r for r in records if int(r["id"]) != int(req_id)]
    save_all_records(records)

# ─── SHOW OLD/NEW COMPARISON ────────────────────────────────────────
def show_old_new_comparison(old_json, new_rec):
    import json
    try:
        old = json.loads(old_json) if old_json and old_json != "{}" else {}
    except:
        old = {}
    if not old:
        st.info("📋 New request — no previous version.")
        return
    st.markdown("#### 🔄 Changes (Previous → New)")
    fields = [
        ("emp_name", "Employee Name"), ("dept", "Department"),
        ("type", "Transaction Type"), ("category", "Category"),
        ("date", "Date"), ("amount", "Amount (£)"),
        ("manager", "Line Manager"), ("desc", "Description")
    ]
    changed = False
    for key, label in fields:
        o = str(old.get(key, "")).strip()
        n = str(new_rec.get(key, "")).strip()
        if o != n:
            changed = True
            st.markdown(f"**{label}**: ~~`{o}`~~ → **`{n}`**")
    if not changed:
        st.info("✅ No changes detected.")

# ============================================================
# IMPORTS & SETUP
# ============================================================
import streamlit as st
import os
import pandas as pd
import shutil
from datetime import datetime

# ============================================================
# PDF GENERATION LIBRARY
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
# PAGE CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="Acoole Electrical Ltd - Portal",
    layout="wide"
)

# ============================================================
# FILE PATHS
# ============================================================
import sys
if "win32" in sys.platform:
    BASE_DIR = r"D:\Acoole_portal"
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploaded_attachments")
PDF_DIR = os.path.join(BASE_DIR, "approved_pdfs")
LOGO_PATH = os.path.join(BASE_DIR, "logo.png")
EXCEL_PATH = os.path.join(BASE_DIR, "requests.xlsx")
USER_DB_PATH = os.path.join(BASE_DIR, "user_database.xlsx")
SETTINGS_PATH = os.path.join(BASE_DIR, "settings.xlsx")
os.makedirs(BASE_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)

# ============================================================
# DEFAULT SETTINGS
# ============================================================
DEFAULT_CATEGORIES = ["Food Allowance", "Others", "Parking", "Parking Fine", "GYM Membership", "Item Not Returned", "Item Missing"]
DEFAULT_ROLES = ["Manager", "Director", "Payroll", "Super Admin"]
DEFAULT_DEPARTMENTS = ["National Grid", "Isolator", "Project", "Accounts", "Payroll Department", "ACoole Electrical Ltd"]

# ============================================================
# SETTINGS FUNCTIONS
# ============================================================
def init_settings():
    if not os.path.exists(SETTINGS_PATH):
        pd.DataFrame([
            {"setting": "categories", "value": "|".join(DEFAULT_CATEGORIES)},
            {"setting": "roles", "value": "|".join(DEFAULT_ROLES)},
            {"setting": "departments", "value": "|".join(DEFAULT_DEPARTMENTS)}
        ]).to_excel(SETTINGS_PATH, index=False)

def load_departments():
    init_settings()
    try:
        df = pd.read_excel(SETTINGS_PATH).fillna("")
        for _, r in df.iterrows():
            if r["setting"] == "departments":
                vals = [v.strip() for v in str(r["value"]).split("|") if v.strip()]
                return vals if vals else DEFAULT_DEPARTMENTS.copy()
        return DEFAULT_DEPARTMENTS.copy()
    except Exception as e:
        print(f"Load depts error: {e}")
        return DEFAULT_DEPARTMENTS.copy()

def save_departments(dept_list):
    init_settings()
    df = pd.read_excel(SETTINGS_PATH).fillna("")
    found = False
    for idx, r in df.iterrows():
        if r["setting"] == "departments":
            df.at[idx, "value"] = "|".join(dept_list)
            found = True
    if not found:
        df = pd.concat([df, pd.DataFrame([{"setting": "departments", "value": "|".join(dept_list)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False)

def load_categories():
    init_settings()
    try:
        df = pd.read_excel(SETTINGS_PATH).fillna("")
        for _, r in df.iterrows():
            if r["setting"] == "categories":
                vals = [v.strip() for v in str(r["value"]).split("|") if v.strip()]
                return vals if vals else DEFAULT_CATEGORIES
        return DEFAULT_CATEGORIES
    except:
        return DEFAULT_CATEGORIES

def save_categories(cat_list):
    init_settings()
    df = pd.read_excel(SETTINGS_PATH).fillna("")
    found = False
    for idx, r in df.iterrows():
        if r["setting"] == "categories":
            df.at[idx, "value"] = "|".join(cat_list)
            found = True
    if not found:
        df = pd.concat([df, pd.DataFrame([{"setting": "categories", "value": "|".join(cat_list)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False)

def load_roles():
    init_settings()
    try:
        df = pd.read_excel(SETTINGS_PATH).fillna("")
        for _, r in df.iterrows():
            if r["setting"] == "roles":
                vals = [v.strip() for v in str(r["value"]).split("|") if v.strip()]
                return vals if vals else DEFAULT_ROLES
        return DEFAULT_ROLES
    except:
        return DEFAULT_ROLES

def save_roles(roles_list):
    init_settings()
    df = pd.read_excel(SETTINGS_PATH).fillna("")
    found = False
    for idx, r in df.iterrows():
        if r["setting"] == "roles":
            df.at[idx, "value"] = "|".join(roles_list)
            found = True
    if not found:
        df = pd.concat([df, pd.DataFrame([{"setting": "roles", "value": "|".join(roles_list)}])], ignore_index=True)
    df.to_excel(SETTINGS_PATH, index=False)

# ============================================================
# USER DATABASE
# ============================================================
DEFAULT_USERS = [
    {"full_name": "National Grid Manager", "username": "national_grid", "password": "acoole123", "role": "Manager", "dept": "National Grid"},
    {"full_name": "Isolator Manager", "username": "isolator", "password": "acoole123", "role": "Manager", "dept": "Isolator"},
    {"full_name": "Project Manager", "username": "project", "password": "acoole123", "role": "Manager", "dept": "Project"},
    {"full_name": "Accounts Manager", "username": "accounts", "password": "acoole123", "role": "Manager", "dept": "Accounts"},
    {"full_name": "Andy Acoole", "username": "andy", "password": "andy2026", "role": "Director", "dept": "ACoole Electrical Ltd"},
    {"full_name": "System Administrator", "username": "wais", "password": "superadmin123", "role": "Super Admin", "dept": "System Administration"},
    {"full_name": "Payroll Team", "username": "payroll", "password": "payroll2026", "role": "Payroll", "dept": "Payroll Department"}
]

def init_user_db():
    if not os.path.exists(USER_DB_PATH):
        pd.DataFrame(DEFAULT_USERS).to_excel(USER_DB_PATH, index=False)

def load_users():
    init_user_db()
    try:
        df = pd.read_excel(USER_DB_PATH).fillna("")
        users = {}
        for _, r in df.iterrows():
            users[r["username"]] = {
                "full_name": str(r.get("full_name", r["username"])).strip(),
                "password": str(r["password"]),
                "role": str(r["role"]),
                "dept": str(r["dept"])
            }
        return users
    except Exception as e:
        st.error(f"User DB Load Error: {e}")
        return {}

def save_users(users_dict):
    export = []
    for uname, data in users_dict.items():
        export.append({
            "full_name": data.get("full_name", uname),
            "username": uname,
            "password": data["password"],
            "role": data["role"],
            "dept": data["dept"]
        })
    pd.DataFrame(export).to_excel(USER_DB_PATH, index=False)

# ============================================================
# REQUESTS EXCEL
# ============================================================
EXCEL_COLUMNS = [
    "ID", "Employee Name", "Department", "Transaction Type", "Category Reason",
    "Date", "Amount (£)", "Line Manager", "Description", "Attachment Name",
    "Status", "Director Comments", "Decision Date", "Decision By",
    "PDF File Path", "Edited From ID", "Old Data"
]

def initialise_excel():
    if not os.path.exists(EXCEL_PATH):
        pd.DataFrame(columns=EXCEL_COLUMNS).to_excel(EXCEL_PATH, index=False)
    else:
        df = pd.read_excel(EXCEL_PATH)
        for col in EXCEL_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df.to_excel(EXCEL_PATH, index=False)

initialise_excel()

def load_records_from_excel():
    try:
        if not os.path.exists(EXCEL_PATH): return []
        df = pd.read_excel(EXCEL_PATH).fillna("")
        if df.empty: return []
        records = df.to_dict(orient="records")
        parsed = []
        for r in records:
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
                "pdf_path": str(r.get("PDF File Path", "")).strip(),
                "edited_from_id": str(r.get("Edited From ID", "")).strip(),
                "old_data": str(r.get("Old Data", "")).strip()
            })
        return parsed
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
            "PDF File Path": str(r.get("pdf_path", "")),
            "Edited From ID": str(r.get("edited_from_id", "")),
            "Old Data": str(r.get("old_data", ""))
        })
    pd.DataFrame(export, columns=EXCEL_COLUMNS).to_excel(EXCEL_PATH, index=False)

def save_record_to_excel(new_record):
    current = load_records_from_excel()
    current.append(new_record)
    save_all_records(current)

# ============================================================
# PDF GENERATION — ✅ COURIER FONT + DOUBLE LINE + LOGO + STAMP
# ============================================================
def generate_approval_pdf(request_data):
    if not PDF_AVAILABLE:
        return False, None, "Install fpdf2: pip install fpdf2"
    try:
        # ✅ READ FRESH DATA FROM EXCEL
        req_id = request_data.get("id")
        all_recs = load_records_from_excel()
        fresh_data = next(
            (r for r in all_recs if int(str(r.get("id", "0"))) == int(str(req_id))),
            request_data
        )

        # ✅ CLEAN TEXT — REMOVE UNSUPPORTED CHARACTERS
        def clean_text(t):
            t = str(t)
            t = t.replace("\u2013", "-")
            t = t.replace("\u2014", "-")
            t = t.replace("\u2212", "-")
            t = t.replace("—", "-")
            t = t.replace("–", "-")
            for char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
                t = t.replace(char, " ")
            return t.strip()

        # ✅ SAFE DATE FORMAT
        def format_date(d):
            if not d or str(d).strip() == "" or str(d).strip().lower() in ["none", "nan"]:
                return "-"
            return str(d).strip()[:10]

        # ─── EXTRACT FIELDS ──────────────────────────────
        emp_name     = clean_text(fresh_data.get("emp_name", "Unknown"))
        dept         = clean_text(fresh_data.get("dept", ""))
        category     = clean_text(fresh_data.get("category", ""))
        amount       = clean_text(fresh_data.get("amount", "0"))
        req_date     = format_date(fresh_data.get("date", ""))
        desc         = clean_text(fresh_data.get("desc", ""))
        manager      = clean_text(fresh_data.get("manager", ""))
        status       = clean_text(fresh_data.get("status", "Pending")).strip().capitalize()
        dir_approve  = format_date(fresh_data.get("decision_date", ""))
        dir_name     = clean_text(fresh_data.get("decision_by", "Director"))

        # ─── CREATE PDF ───────────────────────────────────
        pdf = FPDF()
        pdf.add_page()

        # ✅ LOGO — centered at top
        LOGO_PATH = "logo.png"
        if os.path.exists(LOGO_PATH):
            pdf.image(LOGO_PATH, x=75, y=10, w=60)

        # ✅ MOVE DOWN AFTER LOGO
        pdf.ln(22)

        # ✅ FORM TITLE — Courier font
        pdf.set_font("Courier", "", 11)
        pdf.cell(0, 5, txt="Addition & Deduction Approval Form", ln=True, align="C")
        pdf.ln(3)

        # ✅ DOUBLE HORIZONTAL LINE — matches your screenshot
        line_y = pdf.get_y()
        pdf.line(10, line_y, 200, line_y)
        pdf.line(10, line_y + 1.5, 200, line_y + 1.5)
        pdf.ln(12)

        # ─── REQUEST DETAILS ─────────────────────────────
        pdf.set_font("Courier", "B", 10)
        pdf.cell(0, 5, txt="REQUEST DETAILS", ln=True)
        pdf.ln(2)
        pdf.set_font("Courier", "", 9)
        pdf.cell(52, 5, "Request ID:", 0, 0)
        pdf.cell(0, 5, str(fresh_data.get("id", "")), ln=True)
        pdf.cell(52, 5, "Employee Name:", 0, 0)
        pdf.cell(0, 5, emp_name, ln=True)
        pdf.cell(52, 5, "Department:", 0, 0)
        pdf.cell(0, 5, dept, ln=True)
        pdf.cell(52, 5, "Transaction Type:", 0, 0)
        pdf.cell(0, 5, clean_text(fresh_data.get("type", "")), ln=True)
        pdf.cell(52, 5, "Category / Reason:", 0, 0)
        pdf.cell(0, 5, category, ln=True)
        pdf.cell(52, 5, "Request Date:", 0, 0)
        pdf.cell(0, 5, req_date, ln=True)
        pdf.cell(52, 5, "Amount Approved:", 0, 0)
        pdf.cell(0, 5, f"£{amount}", ln=True)
        pdf.cell(52, 5, "Line Manager:", 0, 0)
        pdf.cell(0, 5, manager, ln=True)
        pdf.ln(6)

        # ─── DESCRIPTION / JUSTIFICATION ──────────────────
        pdf.set_font("Courier", "B", 10)
        pdf.cell(0, 5, txt="DESCRIPTION / JUSTIFICATION", ln=True)
        pdf.ln(2)
        pdf.set_font("Courier", "", 9)
        pdf.multi_cell(0, 5, desc)
        pdf.ln(8)

        # ─── DIRECTOR APPROVAL ────────────────────────────
        pdf.set_font("Courier", "B", 10)
        pdf.cell(0, 5, txt="DIRECTOR APPROVAL", ln=True)
        pdf.ln(2)
        pdf.set_font("Courier", "", 9)

        # ✅ APPROVED — Green text + Approved By + Date + Stamp
        if status.lower() == "approved" and dir_approve != "-":
            pdf.cell(52, 5, "Decision:", 0, 0)
            pdf.set_font("Courier", "B", 9)
            pdf.set_text_color(0, 128, 0)  # Green
            pdf.cell(0, 5, "APPROVED", ln=True)
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Courier", "", 9)
            pdf.cell(52, 5, "Approved By:", 0, 0)
            pdf.cell(0, 5, dir_name, ln=True)
            pdf.cell(52, 5, "Approval Date / Time:", 0, 0)
            pdf.cell(0, 5, dir_approve, ln=True)
            
            # ✅ Show Director Comments if provided
            dir_comments = fresh_data.get("director_comments", "").strip()
            if dir_comments and dir_comments.lower() != "none":
                pdf.ln(2)
                pdf.set_font("Courier", "B", 9)
                pdf.cell(52, 5, "Director Comments:", 0, 0)
                pdf.set_font("Courier", "", 9)
                pdf.ln(5)
                pdf.multi_cell(0, 5, dir_comments)

        # ✅ REJECTED — Red text + Rejected By + Date + REASON
        elif status.lower() == "rejected" and dir_approve != "-":
            pdf.cell(52, 5, "Decision:", 0, 0)
            pdf.set_font("Courier", "B", 9)
            pdf.set_text_color(200, 0, 0)  # Red
            pdf.cell(0, 5, "REJECTED", ln=True)
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Courier", "", 9)
            pdf.cell(52, 5, "Rejected By:", 0, 0)
            pdf.cell(0, 5, dir_name, ln=True)
            pdf.cell(52, 5, "Rejection Date / Time:", 0, 0)
            pdf.cell(0, 5, dir_approve, ln=True)
            
            # ✅ SHOW REJECTION REASON / COMMENT IN PDF
            dir_comments = fresh_data.get("director_comments", "").strip()
            if dir_comments and dir_comments.lower() != "none":
                pdf.ln(2)
                pdf.set_font("Courier", "B", 9)
                pdf.cell(52, 5, "Reason for Rejection:", 0, 0)
                pdf.set_font("Courier", "", 9)
                pdf.ln(5)
                pdf.multi_cell(0, 5, dir_comments)

        # ⏳ PENDING — Default
        else:
            pdf.cell(52, 5, "Decision:", 0, 0)
            pdf.cell(0, 5, "Pending", ln=True)

        pdf.ln(10)

        # ✅ DASHED / BROKEN LINE above signature
        dash_y = pdf.get_y()
        for x in range(10, 200, 4):
            pdf.line(x, dash_y, x + 2, dash_y)

        # ✅ APPROVED or REJECTED STAMP — centered on dashed line
        APPROVED_STAMP_PATH = "approved_stamp.png"
        REJECTED_STAMP_PATH = "rejected_stamp.png"
        if status.lower() == "approved" and os.path.exists(APPROVED_STAMP_PATH):
            pdf.image(APPROVED_STAMP_PATH, x=75, y=dash_y - 6, w=60)
        elif status.lower() == "rejected" and os.path.exists(REJECTED_STAMP_PATH):
            pdf.image(REJECTED_STAMP_PATH, x=75, y=dash_y - 6, w=60)

        pdf.ln(8)

        # ✅ SIGNATURE TEXT
        pdf.set_font("Courier", "", 8)
        pdf.cell(0, 5, txt="Authorised Signature / Director", ln=True)

        # ─── FILENAME ───
        safe_id = clean_text(str(req_id))
        safe_name = emp_name
        safe_category = category
        safe_date = datetime.now().strftime("%Y-%m-%d")
        filename = f"{safe_id}# {safe_name} - {safe_category} - {safe_date}.pdf"

        # ─── BYTES FIX — Works on ALL fpdf2 versions ───
        pdf_output = pdf.output()
        if isinstance(pdf_output, (bytes, bytearray)):
            pdf_bytes = bytes(pdf_output)
        else:
            pdf_bytes = pdf_output.encode("latin-1")

        # ✅ Save to disk
        os.makedirs(PDF_DIR, exist_ok=True)
        full_pdf_path = os.path.join(PDF_DIR, filename)
        with open(full_pdf_path, "wb") as f:
            f.write(pdf_bytes)

        return True, pdf_bytes, filename

    except Exception as e:
        return False, None, f"PDF Error: {str(e)}"

# ============================================================
# SESSION STATE
# ============================================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_info" not in st.session_state:
    st.session_state.user_info = None
if "editing_request_id" not in st.session_state:
    st.session_state.editing_request_id = None

# ============================================================
# REFRESH BUTTON
# ============================================================
def refresh_data_button():
    if st.button("🔄 Refresh Data", type="secondary", key="refresh_data_btn"):
        st.session_state["_last_refresh"] = datetime.now().isoformat()
        st.rerun()

# ============================================================
# SIDEBAR — PAYROLL PDF GENERATION
# ============================================================
if st.session_state.logged_in and st.session_state.user_info:
    if st.session_state.user_info.get("role") in ["Payroll", "Super Admin"]:
        if st.sidebar.button("🔧 Generate ALL Approved PDFs"):
            if PDF_AVAILABLE:
                all_live_requests = load_records_from_excel()
                count = 0
                for rec in all_live_requests:
                    if rec["status"] == "approved":
                        ok, pdf_bytes, name = generate_approval_pdf(rec)
                        if ok:
                            count += 1
                if count > 0:
                    st.sidebar.success(f"✅ Generated {count} PDFs!")
                    st.rerun()
                else:
                    st.sidebar.info("✅ No new PDFs needed.")
            else:
                st.sidebar.error("⚠️ Install fpdf2: pip install fpdf2")

# ============================================================
# SETTINGS PANEL
# ============================================================
def settings_management_panel():
    st.subheader("⚙️ System Settings — Categories, Departments & Roles")
    st.info("🛡️ Super Admin Only — Add, Edit, Delete Categories, Departments and Permission Roles.")
    st.divider()

    cats_tab, dept_tab, roles_tab = st.tabs([
        "🏷️ Manage Categories", "🏢 Manage Departments", "🎖️ Manage Roles / Permissions"
    ])

    with cats_tab:
        st.markdown("### 🏷️ Request Categories")
        st.info("These options appear in the request form dropdown.")
        current_cats = load_categories()
        with st.form("add_category_form", clear_on_submit=True):
            new_cat = st.text_input("➕ Add New Category", placeholder="e.g. Travel Allowance")
            if st.form_submit_button("✅ Add Category"):
                if new_cat.strip() and new_cat.strip() not in current_cats:
                    current_cats.append(new_cat.strip())
                    save_categories(current_cats)
                    st.success(f"✅ Added: {new_cat}")
                    st.rerun()
                elif new_cat.strip() in current_cats:
                    st.warning("⚠️ Category already exists!")
        st.divider()
        st.markdown("#### Current Categories")
        for i, cat in enumerate(current_cats):
            c1, c2, c3 = st.columns([4, 1, 1])
            with c1: st.markdown(f"• **{cat}**")
            with c2:
                if st.button(f"✏️ Edit", key=f"edit_cat_{i}"):
                    st.session_state[f"editing_cat_{i}"] = True
            with c3:
                if len(current_cats) > 1 and st.button(f"🗑️ Delete", key=f"del_cat_{i}"):
                    current_cats.pop(i)
                    save_categories(current_cats)
                    st.success(f"✅ Deleted: {cat}")
                    st.rerun()
            if st.session_state.get(f"editing_cat_{i}", False):
                with st.form(f"save_cat_form_{i}", clear_on_submit=True):
                    renamed = st.text_input("Rename Category", value=cat)
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.form_submit_button("💾 Save"):
                            current_cats[i] = renamed.strip()
                            save_categories(current_cats)
                            st.session_state[f"editing_cat_{i}"] = False
                            st.success(f"✅ Renamed to: {renamed}")
                            st.rerun()
                    with col2:
                        if st.form_submit_button("❌ Cancel"):
                            st.session_state[f"editing_cat_{i}"] = False
                            st.rerun()

    with dept_tab:
        st.markdown("### 🏢 Manage Departments")
        st.info("Create new departments, rename or remove existing ones.")
        st.divider()
        current_depts = load_departments()
        with st.form("add_dept_form", clear_on_submit=True):
            new_dept = st.text_input("➕ Add New Department", placeholder="e.g. Maintenance, HR")
            if st.form_submit_button("✅ Add Department"):
                if new_dept.strip() and new_dept.strip() not in current_depts:
                    current_depts.append(new_dept.strip())
                    save_departments(current_depts)
                    st.success(f"✅ Added: {new_dept}")
                    st.rerun()
                elif new_dept.strip() in current_depts:
                    st.warning("⚠️ Department already exists!")
        st.divider()
        st.markdown("#### Current Departments")
        for i, dept_name in enumerate(current_depts):
            c1, c2, c3 = st.columns([4, 1, 1])
            with c1: st.markdown(f"• **{dept_name}**")
            with c2:
                if st.button(f"✏️ Edit", key=f"edit_dept_{i}"):
                    st.session_state[f"editing_dept_{i}"] = True
            with c3:
                if len(current_depts) > 1 and st.button(f"🗑️ Delete", key=f"del_dept_{i}"):
                    current_depts.pop(i)
                    save_departments(current_depts)
                    st.success(f"✅ Deleted: {dept_name}")
                    st.rerun()
            if st.session_state.get(f"editing_dept_{i}", False):
                with st.form(f"save_dept_form_{i}", clear_on_submit=True):
                    renamed = st.text_input("Rename Department", value=dept_name)
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.form_submit_button("💾 Save"):
                            current_depts[i] = renamed.strip()
                            save_departments(current_depts)
                            st.session_state[f"editing_dept_{i}"] = False
                            st.success(f"✅ Renamed to: {renamed}")
                            st.rerun()
                    with col2:
                        if st.form_submit_button("❌ Cancel"):
                            st.session_state[f"editing_dept_{i}"] = False
                            st.rerun()

    with roles_tab:
        st.markdown("### 🎖️ User Roles / Permission Levels")
        st.info("⚠️ 'Super Admin' cannot be deleted to keep system access.")
        current_roles = load_roles()
        with st.form("add_role_form", clear_on_submit=True):
            new_role = st.text_input("➕ Add New Role", placeholder="e.g. HR Manager")
            if st.form_submit_button("✅ Add Role"):
                if new_role.strip() and new_role.strip() not in current_roles:
                    current_roles.append(new_role.strip())
                    save_roles(current_roles)
                    st.success(f"✅ Added: {new_role}")
                    st.rerun()
                elif new_role.strip() in current_roles:
                    st.warning("⚠️ Role already exists!")
        st.divider()
        st.markdown("#### Current Roles")
        for i, role in enumerate(current_roles):
            c1, c2, c3 = st.columns([4, 1, 1])
            with c1: st.markdown(f"• **{role}**")
            with c2:
                if st.button(f"✏️ Edit", key=f"edit_role_{i}"):
                    st.session_state[f"editing_role_{i}"] = True
            with c3:
                if role != "Super Admin" and len(current_roles) > 1 and st.button(f"🗑️ Delete", key=f"del_role_{i}"):
                    current_roles.pop(i)
                    save_roles(current_roles)
                    st.success(f"✅ Deleted: {role}")
                    st.rerun()
            if st.session_state.get(f"editing_role_{i}", False):
                with st.form(f"save_role_form_{i}", clear_on_submit=True):
                    renamed = st.text_input("Rename Role", value=role)
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.form_submit_button("💾 Save"):
                            current_roles[i] = renamed.strip()
                            save_roles(current_roles)
                            st.session_state[f"editing_role_{i}"] = False
                            st.success(f"✅ Renamed to: {renamed}")
                            st.rerun()
                    with col2:
                        if st.form_submit_button("❌ Cancel"):
                            st.session_state[f"editing_role_{i}"] = False
                            st.rerun()

# ============================================================
# USER MANAGEMENT PANEL
# ============================================================
def user_management_panel():
    st.subheader("👤 User Management — Create & Manage System Users")
    st.info("🛡️ Super Admin Only — Create, edit, or delete user accounts.")
    st.divider()
    USERS = load_users()
    ROLES = load_roles()
    tab1, tab2, tab3 = st.tabs(["➕ Create New User", "✏️ Edit User", "🗑️ Delete User"])

    with tab1:
        st.markdown("### ➕ Create New System User")
        with st.form("create_user_form", border=True, clear_on_submit=True):
            new_full_name = st.text_input("👤 Full Name", placeholder="e.g. John Smith")
            new_username = st.text_input("🔐 Username", placeholder="e.g. john_smith").lower().strip()
            new_password = st.text_input("🔑 Password", type="password")
            new_role = st.selectbox("🎖️ Role / Permission Level", ROLES)
            new_dept = st.selectbox("🏢 Department", load_departments())
            create_btn = st.form_submit_button("✅ Create User Account", type="primary")
            if create_btn:
                if not new_full_name.strip() or not new_username or not new_password:
                    st.error("❌ All fields required!")
                elif new_username in USERS:
                    st.error(f"❌ Username '{new_username}' already exists!")
                else:
                    USERS[new_username] = {
                        "full_name": new_full_name.strip(),
                        "password": new_password,
                        "role": new_role,
                        "dept": new_dept
                    }
                    save_users(USERS)
                    st.success(f"✅ User **'{new_full_name}'** created!")
                    st.balloons()

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
                if st.form_submit_button("🔄 Update User", type="primary"):
                    USERS = load_users()
                    if upd_username_new != edit_user_sel:
                        if upd_username_new in USERS:
                            st.error(f"❌ Username '{upd_username_new}' already exists!")
                            return
                        USERS[upd_username_new] = {
                            "full_name": upd_full_name.strip(),
                            "password": upd_password if upd_password else curr["password"],
                            "role": upd_role,
                            "dept": upd_dept
                        }
                        del USERS[edit_user_sel]
                    else:
                        USERS[edit_user_sel]["full_name"] = upd_full_name.strip()
                        if upd_password:
                            USERS[edit_user_sel]["password"] = upd_password
                        USERS[edit_user_sel]["role"] = upd_role
                        USERS[edit_user_sel]["dept"] = upd_dept
                    save_users(USERS)
                    st.success(f"✅ User updated: **{upd_full_name}**")
                    st.rerun()

    with tab3:
        st.markdown("### ⚠️ Delete User Account")
        st.warning("Existing requests remain safe.")
        del_user_sel = st.selectbox("Select User to DELETE", [u for u in USERS.keys() if u != st.session_state.user_info["username"]])
        if del_user_sel:
            del_name = USERS[del_user_sel].get("full_name", del_user_sel)
            if st.button(f"🗑️ DELETE: {del_name} ({del_user_sel})", type="secondary"):
                del USERS[del_user_sel]
                save_users(USERS)
                st.success(f"✅ User **{del_name}** deleted!")
                st.rerun()

# ============================================================
# LOGIN PAGE
# ============================================================
def display_company_header():
    import streamlit as st
    import os
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        LOGO_PATH = "logo.png"
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=300)
        else:
            st.title("⚡ ACOOLE ELECTRICAL LTD")
        st.caption("Addition & Deduction Approval Platform")
        st.divider()

if not st.session_state.logged_in:
    display_company_header()
    with st.form("login_form", border=True):
        st.markdown("### 🔒 Secure Gateway Login")
        st.caption("Enter your credentials to access the system")
        st.divider()
        username = st.text_input("🔐 Username", placeholder="e.g. andy, payroll, wais").lower().strip()
        password = st.text_input("🔑 Password", type="password", placeholder="Enter your password")
        login_btn = st.form_submit_button("🔐 Authenticate Portal", type="primary", use_container_width=True)
        if login_btn:
            USERS = load_users()
            if username in USERS and USERS[username]["password"] == password:
                st.session_state.logged_in = True
                st.session_state.user_info = {**USERS[username], "username": username}
                st.rerun()
            else:
                st.error("❌ Invalid Username or Password. Please try again.")

# ============================================================
# MAIN APPLICATION
# ============================================================
else:
    user = st.session_state.user_info
    FULL_NAME = user.get("full_name", user["username"])
    CATEGORIES = load_categories()
    refresh_data_button()
    all_live_requests = load_records_from_excel()
    display_company_header()

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"🟢 **Welcome:** {FULL_NAME} | {user['dept']} | **{user['role']}**")
    with c2:
        if st.button("🚪 Secure Logout"):
            st.session_state.logged_in = False
            st.session_state.user_info = None
            st.rerun()
    st.divider()

    # ========================================================
    # 🔐 ROLE-BASED PORTALS
    # ========================================================
    
    if user["role"] == "Super Admin":
        # ─── SUPER ADMIN PORTAL ───────────────────────────
        st.subheader("🔧 Super Admin Portal")
        st.info("Full system access — manage everything.")
        # ... keep your existing Super Admin code here ...

    elif user["role"] == "Director":
        # ========================================================
        # 🎛️ DIRECTOR PORTAL — WITH FULL STATUS CONTROL
        # ========================================================
        st.subheader("🎛️ Director Approval Portal — Andy Acoole")
        st.info("✅ Review all requests, Approve, Reject, OR Change Status. Decisions update automatically.")
        st.divider()
        tab_pending, tab_approved, tab_rejected = st.tabs([
            "⏳ Pending Requests",
            "✅ Approved Requests",
            "❌ Rejected Requests"
        ])
        
        with tab_pending:
            pending = [r for r in all_live_requests if r["status"] == "pending"]
            if not pending:
                st.success("✅ No pending requests — all reviewed!")
            else:
                st.metric("⏳ Pending Approval", len(pending))
                st.divider()
                for req in reversed(pending):
                    title = f"🟡 ID #{req['id']} | {req['emp_name']} | 📅 {format_date(req['date'])} | £{req['amount']:.2f} | {req['dept']}"
                    with st.expander(title):
                        st.write(f"👤 **Employee:** {req['emp_name']}")
                        st.write(f"🏢 **Department:** {req['dept']}")
                        st.write(f"🔄 **Transaction Type:** {req['type']}")
                        st.write(f"🏷️ **Category / Reason:** {req['category']}")
                        st.write(f"💷 **Amount:** £{req['amount']:.2f}")
                        st.write(f"👔 **Line Manager:** {req['manager']}")
                        st.write(f"📅 **Request Date:** {req['date']}")
                        st.info(f"📝 **Description:** {req['desc']}")
                        display_attachments(req)
                        st.divider()
                        
                        with st.form(f"change_status_pending_{req['id']}"):
                            st.subheader("🔧 Change Status")
                            comments = st.text_area("💬 Director Comments (Optional)")
                            col_approve, col_reject = st.columns(2)
                            with col_approve:
                                approve_btn = st.form_submit_button("✅ APPROVE", type="primary")
                            with col_reject:
                                reject_btn = st.form_submit_button("❌ REJECT", type="secondary")
                            
                            if approve_btn:
                                update_record_status_in_excel(req["id"], "approved", comments, FULL_NAME)
                                st.success(f"✅ Request #{req['id']} APPROVED! Status updated.")
                                st.rerun()
                            if reject_btn:
                                update_record_status_in_excel(req["id"], "rejected", comments, FULL_NAME)
                                st.warning(f"❌ Request #{req['id']} REJECTED! Status updated.")
                                st.rerun()
                        
                        st.divider()
                        display_pdf_button(req, can_generate=True)
        
        with tab_approved:
            approved = [r for r in all_live_requests if r["status"] == "approved"]
            if not approved:
                st.info("📋 No approved requests yet.")
            else:
                st.metric("✅ Total Approved", len(approved))
                st.divider()
                for req in reversed(approved):
                    approved_by_line = f"✅ Approved by {req.get('decision_by', 'Director')} on {format_date(req.get('decision_date', ''))}"
                    title = f"🟢 ID #{req['id']} | {req['emp_name']} | £{req['amount']:.2f} | {approved_by_line}"
                    with st.expander(title):
                        st.write(f"👤 **Employee:** {req['emp_name']}")
                        st.write(f"🏢 **Department:** {req['dept']}")
                        st.write(f"🔄 **Type:** {req['type']} | 🏷️ **Category:** {req['category']}")
                        st.write(f"💷 **Amount:** £{req['amount']:.2f}")
                        st.write(f"👔 **Line Manager:** {req['manager']}")
                        st.write(f"📅 **Request Date:** {req['date']}")
                        st.success(f"💬 **Director Comments:** {req.get('director_comments', 'None')}")
                        display_attachments(req)
                        st.divider()
                        
                        with st.form(f"change_status_approved_{req['id']}"):
                            st.subheader("🔧 Change Status")
                            comments = st.text_area("💬 Updated Comments (Optional)")
                            col_pending, col_reject = st.columns(2)
                            with col_pending:
                                pending_btn = st.form_submit_button("⏳ Move to Pending")
                            with col_reject:
                                reject_btn = st.form_submit_button("❌ Change to REJECTED", type="secondary")
                            
                            if pending_btn:
                                update_record_status_in_excel(req["id"], "pending", comments, FULL_NAME)
                                st.info(f"⏳ Request #{req['id']} moved back to PENDING!")
                                st.rerun()
                            if reject_btn:
                                update_record_status_in_excel(req["id"], "rejected", comments, FULL_NAME)
                                st.warning(f"❌ Request #{req['id']} changed to REJECTED!")
                                st.rerun()
                        
                        st.divider()
                        display_pdf_button(req, can_generate=True)
        
        with tab_rejected:
            rejected = [r for r in all_live_requests if r["status"] == "rejected"]
            if not rejected:
                st.info("📋 No rejected requests yet.")
            else:
                st.metric("❌ Total Rejected", len(rejected))
                st.divider()
                for req in reversed(rejected):
                    rejected_by_line = f"❌ Rejected by {req.get('decision_by', 'Director')} on {format_date(req.get('decision_date', ''))}"
                    title = f"🔴 ID #{req['id']} | {req['emp_name']} | £{req['amount']:.2f} | {rejected_by_line}"
                    with st.expander(title):
                        st.write(f"👤 **Employee:** {req['emp_name']}")
                        st.write(f"🏢 **Department:** {req['dept']}")
                        st.write(f"🔄 **Type:** {req['type']} | 🏷️ **Category:** {req['category']}")
                        st.write(f"💷 **Amount:** £{req['amount']:.2f}")
                        st.write(f"👔 **Line Manager:** {req['manager']}")
                        st.write(f"📅 **Request Date:** {req['date']}")
                        st.error(f"💬 **Director Comments:** {req.get('director_comments', 'None')}")
                        display_attachments(req)
                        st.divider()
                        
                        with st.form(f"change_status_rejected_{req['id']}"):
                            st.subheader("🔧 Change Status")
                            comments = st.text_area("💬 Updated Comments (Optional)")
                            col_pending, col_approve = st.columns(2)
                            with col_pending:
                                pending_btn = st.form_submit_button("⏳ Move to Pending")
                            with col_approve:
                                approve_btn = st.form_submit_button("✅ Change to APPROVED", type="primary")
                            
                            if pending_btn:
                                update_record_status_in_excel(req["id"], "pending", comments, FULL_NAME)
                                st.info(f"⏳ Request #{req['id']} moved back to PENDING!")
                                st.rerun()
                            if approve_btn:
                                update_record_status_in_excel(req["id"], "approved", comments, FULL_NAME)
                                st.success(f"✅ Request #{req['id']} changed to APPROVED!")
                                st.rerun()
                        
                        st.divider()
                        display_pdf_button(req, can_generate=True)

    elif user["role"] == "Payroll":
        # ========================================================
        # 🧾 PAYROLL PORTAL — VIEW & DOWNLOAD ONLY (NO REJECT)
        # ========================================================
        st.subheader("🧾 Payroll Portal")
        st.info("✅ View all requests and Download PDFs.")
        st.divider()
        tab_pending, tab_approved, tab_rejected = st.tabs([
            "⏳ Pending Requests",
            "✅ Approved Requests",
            "❌ Rejected Requests"
        ])
        
        with tab_pending:
            pending = [r for r in all_live_requests if r["status"] == "pending"]
            if not pending:
                st.success("✅ No pending requests!")
            else:
                st.metric("⏳ Pending", len(pending))
                st.divider()
                for req in reversed(pending):
                    title = f"🟡 ID #{req['id']} | {req['emp_name']} | £{req['amount']:.2f}"
                    with st.expander(title):
                        st.write(f"👤 **Employee:** {req['emp_name']}")
                        st.write(f"🏢 **Department:** {req['dept']}")
                        st.write(f"🔄 **Type:** {req['type']}")
                        st.write(f"🏷️ **Category:** {req['category']}")
                        st.write(f"💷 **Amount:** £{req['amount']:.2f}")
                        st.write(f"👔 **Line Manager:** {req['manager']}")
                        st.write(f"📅 **Date:** {req['date']}")
                        st.info(f"📝 **Description:** {req['desc']}")
                        display_attachments(req)
                        st.divider()
                        display_pdf_button(req, can_generate=True)
        
        with tab_approved:
            approved = [r for r in all_live_requests if r["status"] == "approved"]
            if not approved:
                st.info("📋 No approved requests yet.")
            else:
                st.metric("✅ Approved", len(approved))
                st.divider()
                for req in reversed(approved):
                    title = f"🟢 ID #{req['id']} | {req['emp_name']} | £{req['amount']:.2f}"
                    with st.expander(title):
                        st.write(f"👤 **Employee:** {req['emp_name']}")
                        st.write(f"🏢 **Department:** {req['dept']}")
                        st.write(f"💷 **Amount:** £{req['amount']:.2f}")
                        st.write(f"📅 **Date:** {req['date']}")
                        st.success(f"💬 **Comments:** {req.get('director_comments', 'None')}")
                        display_attachments(req)
                        st.divider()
                        display_pdf_button(req, can_generate=True)
        
        with tab_rejected:
            rejected = [r for r in all_live_requests if r["status"] == "rejected"]
            if not rejected:
                st.info("📋 No rejected requests yet.")
            else:
                st.metric("❌ Rejected", len(rejected))
                st.divider()
                for req in reversed(rejected):
                    title = f"🔴 ID #{req['id']} | {req['emp_name']} | £{req['amount']:.2f}"
                    with st.expander(title):
                        st.write(f"👤 **Employee:** {req['emp_name']}")
                        st.write(f"🏢 **Department:** {req['dept']}")
                        st.write(f"💷 **Amount:** £{req['amount']:.2f}")
                        st.write(f"📅 **Date:** {req['date']}")
                        st.error(f"💬 **Reason:** {req.get('director_comments', 'None')}")
                        display_attachments(req)
                        st.divider()
                        display_pdf_button(req, can_generate=True)

    # ========================================================
    # MANAGER PORTAL
    # ========================================================
    elif user["role"] == "Manager":
        if st.session_state.editing_request_id:
            eid = st.session_state.editing_request_id
            rec = next((r for r in all_live_requests if int(r["id"]) == int(eid)), None)
            if rec:
                st.subheader(f"✏️ Edit Request #{eid}")
                show_old_new_comparison("{}", rec)
                with st.form("edit_form"):
                    c1, c2 = st.columns(2)
                    with c1:
                        en = st.text_input("👤 Employee Name", rec["emp_name"])
                        rt = st.selectbox("🔄 Transaction Type", ["Addition", "Deduction"], index=["Addition", "Deduction"].index(rec["type"]))
                        ct = st.selectbox("🏷️ Category / Reason", CATEGORIES, index=CATEGORIES.index(rec["category"]) if rec["category"] in CATEGORIES else 0)
                        amt = st.number_input("💷 Amount (£)", min_value=0.01, step=10.0, value=float(rec["amount"]))
                    with c2:
                        from datetime import datetime as dt
                        try: d = dt.strptime(rec["date"][:10], "%Y-%m-%d")
                        except: d = dt.today()
                        dt_val = st.date_input("📅 Date", d)
                        mgr = st.text_input("👔 Line Manager", rec["manager"])
                        desc = st.text_area("📝 Description / Justification", rec["desc"])
                        files = st.file_uploader("📎 Add Documents", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)

                        if st.form_submit_button("✅ Submit Edit"):
                            old = str({"emp_name": rec["emp_name"], "dept": rec["dept"], "type": rec["type"], "category": rec["category"], "date": rec["date"], "amount": rec["amount"], "manager": rec["manager"], "desc": rec["desc"]})
                            records = load_records_from_excel()
                            for r in records:
                                if int(r["id"]) == int(eid):
                                    r["emp_name"] = en.strip()
                                    r["type"] = rt
                                    r["category"] = ct
                                    r["amount"] = amt
                                    r["date"] = str(dt_val)
                                    r["manager"] = mgr.strip()
                                    r["desc"] = desc.strip()
                                    r["status"] = "pending"
                                    r["old_data"] = old
                                    r["director_comments"] = ""
                                    r["decision_date"] = ""
                                    r["decision_by"] = ""
                                    if files:
                                        att_list = []
                                        if r["attachment_name"] and r["attachment_name"] != "None":
                                            att_list.extend([n.strip() for n in r["attachment_name"].split(",")])
                                        for i, f in enumerate(files):
                                            fn = f"ID_{eid}_EDIT_F{len(att_list)+1}_{f.name}"
                                            with open(os.path.join(UPLOAD_DIR, fn), "wb") as out:
                                                out.write(f.getbuffer())
                                            att_list.append(fn)
                                        r["attachment_name"] = ", ".join(att_list) if att_list else "None"
                            save_all_records(records)
                            st.success(f"✅ Updated & sent for approval!")
                            st.session_state.editing_request_id = None
                            st.rerun()
                if st.button("❌ Cancel"):
                    st.session_state.editing_request_id = None
                    st.rerun()
        
        st.subheader(f"➕ New Request — {user['dept']}")
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
                            for i, f in enumerate(files):
                                fn = f"ID_{nid}_F{i+1}_{f.name}"
                                with open(os.path.join(UPLOAD_DIR, fn), "wb") as out:
                                    out.write(f.getbuffer())
                                att_list.append(fn)
                        payload = {
                            "id": nid, "emp_name": en.strip(), "dept": user["dept"], "type": rt,
                            "category": ct, "date": str(dt_val), "amount": amt, "manager": mgr.strip(),
                            "desc": desc.strip(), "attachment_name": ", ".join(att_list) or "None",
                            "status": "pending", "director_comments": "", "decision_date": "",
                            "decision_by": "", "pdf_path": "", "edited_from_id": "", "old_data": ""
                        }
                        save_record_to_excel(payload)
                        st.success(f"✅ Request #{nid} sent for approval!")
                        st.rerun()
                    else:
                        st.error("⚠️ Please fill in: Employee Name, Line Manager, and Description")
        
        st.divider()
        st.subheader(f"📋 My Department Requests")
        my_reqs = [r for r in all_live_requests if r["dept"] == user["dept"]]
        if not my_reqs:
            st.info("📋 No requests yet.")
        else:
            for req in reversed(my_reqs):
                icon = "🟡" if req["status"] == "pending" else ("🟢" if req["status"] == "approved" else "🔴")
                title = f"{icon} ID #{req['id']} | {req['status'].upper()} | £{req['amount']:.2f} | 📅 {format_date(req['date'])}"
                with st.expander(title):
                    st.write(f"👤 Employee: {req['emp_name']} | 👔 Manager: {req['manager']}")
                    st.write(f"🔄 Type: {req['type']} | 🏷️ Category: {req['category']}")
                    st.info(f"📝 Description: {req['desc']}")
                    display_attachments(req)
                    if req["director_comments"]:
                        st.info(f"💬 Director Comments: {req['director_comments']}")
                    if req["status"] == "approved":
                        display_pdf_button(req, can_generate=False)
                    if req["status"] in ["pending", "rejected"]:
                        if st.button(f"✏️ Edit Request #{req['id']}", key=f"edit_{req['id']}"):
                            st.session_state.editing_request_id = req["id"]
                            st.rerun()

        # ========================================================
    # 🎛️ DIRECTOR PORTAL — WITH CHANGE STATUS CONTROLS
    # ========================================================
    elif user["role"] == "Director":
        st.subheader("🎛️ Director Approval Portal — Andy Acoole")
        st.info("✅ Review all requests, Approve, Reject, OR Change Status. Decisions update automatically.")
        st.divider()
        tab_pending, tab_approved, tab_rejected = st.tabs([
            "⏳ Pending Requests",
            "✅ Approved Requests",
            "❌ Rejected Requests"
        ])
        
        with tab_pending:
            pending = [r for r in all_live_requests if r["status"] == "pending"]
            if not pending:
                st.success("✅ No pending requests — all reviewed!")
            else:
                st.metric("⏳ Pending Approval", len(pending))
                st.divider()
                for req in reversed(pending):
                    title = f"🟡 ID #{req['id']} | {req['emp_name']} | 📅 {format_date(req['date'])} | £{req['amount']:.2f} | {req['dept']}"
                    with st.expander(title):
                        st.write(f"👤 **Employee:** {req['emp_name']}")
                        st.write(f"🏢 **Department:** {req['dept']}")
                        st.write(f"🔄 **Transaction Type:** {req['type']}")
                        st.write(f"🏷️ **Category / Reason:** {req['category']}")
                        st.write(f"💷 **Amount:** £{req['amount']:.2f}")
                        st.write(f"👔 **Line Manager:** {req['manager']}")
                        st.write(f"📅 **Request Date:** {req['date']}")
                        st.info(f"📝 **Description:** {req['desc']}")
                        display_attachments(req)
                        st.divider()
                        
                        # ✅ CHANGE STATUS — PENDING → APPROVE / REJECT
                        with st.form(f"change_status_pending_{req['id']}"):
                            st.subheader("🔧 Change Status")
                            comments = st.text_area("💬 Director Comments (Optional)")
                            col_approve, col_reject = st.columns(2)
                            with col_approve:
                                approve_btn = st.form_submit_button("✅ APPROVE", type="primary")
                            with col_reject:
                                reject_btn = st.form_submit_button("❌ REJECT", type="secondary")
                            
                            if approve_btn:
                                update_record_status_in_excel(req["id"], "approved", comments, FULL_NAME)
                                st.success(f"✅ Request #{req['id']} APPROVED! Status updated.")
                                st.rerun()
                            if reject_btn:
                                update_record_status_in_excel(req["id"], "rejected", comments, FULL_NAME)
                                st.warning(f"❌ Request #{req['id']} REJECTED! Status updated.")
                                st.rerun()
                        
                        st.divider()
                        display_pdf_button(req, can_generate=True)
        
        with tab_approved:
            approved = [r for r in all_live_requests if r["status"] == "approved"]
            if not approved:
                st.info("📋 No approved requests yet.")
            else:
                st.metric("✅ Total Approved", len(approved))
                st.divider()
                for req in reversed(approved):
                    approved_by_line = f"✅ Approved by {req.get('decision_by', 'Director')} on {format_date(req.get('decision_date', ''))}"
                    title = f"🟢 ID #{req['id']} | {req['emp_name']} | £{req['amount']:.2f} | {approved_by_line}"
                    with st.expander(title):
                        st.write(f"👤 **Employee:** {req['emp_name']}")
                        st.write(f"🏢 **Department:** {req['dept']}")
                        st.write(f"🔄 **Type:** {req['type']} | 🏷️ **Category:** {req['category']}")
                        st.write(f"💷 **Amount:** £{req['amount']:.2f}")
                        st.write(f"👔 **Line Manager:** {req['manager']}")
                        st.write(f"📅 **Request Date:** {req['date']}")
                        st.success(f"💬 **Director Comments:** {req.get('director_comments', 'None')}")
                        display_attachments(req)
                        st.divider()
                        
                        # ✅ CHANGE STATUS — APPROVED → PENDING / REJECTED
                        with st.form(f"change_status_approved_{req['id']}"):
                            st.subheader("🔧 Change Status")
                            comments = st.text_area("💬 Updated Comments (Optional)")
                            col_pending, col_reject = st.columns(2)
                            with col_pending:
                                pending_btn = st.form_submit_button("⏳ Move to Pending")
                            with col_reject:
                                reject_btn = st.form_submit_button("❌ Change to REJECTED", type="secondary")
                            
                            if pending_btn:
                                update_record_status_in_excel(req["id"], "pending", comments, FULL_NAME)
                                st.info(f"⏳ Request #{req['id']} moved back to PENDING!")
                                st.rerun()
                            if reject_btn:
                                update_record_status_in_excel(req["id"], "rejected", comments, FULL_NAME)
                                st.warning(f"❌ Request #{req['id']} changed to REJECTED!")
                                st.rerun()
                        
                        st.divider()
                        display_pdf_button(req, can_generate=True)
        
        with tab_rejected:
            rejected = [r for r in all_live_requests if r["status"] == "rejected"]
            if not rejected:
                st.info("📋 No rejected requests yet.")
            else:
                st.metric("❌ Total Rejected", len(rejected))
                st.divider()
                for req in reversed(rejected):
                    rejected_by_line = f"❌ Rejected by {req.get('decision_by', 'Director')} on {format_date(req.get('decision_date', ''))}"
                    title = f"🔴 ID #{req['id']} | {req['emp_name']} | £{req['amount']:.2f} | {rejected_by_line}"
                    with st.expander(title):
                        st.write(f"👤 **Employee:** {req['emp_name']}")
                        st.write(f"🏢 **Department:** {req['dept']}")
                        st.write(f"🔄 **Type:** {req['type']} | 🏷️ **Category:** {req['category']}")
                        st.write(f"💷 **Amount:** £{req['amount']:.2f}")
                        st.write(f"👔 **Line Manager:** {req['manager']}")
                        st.write(f"📅 **Request Date:** {req['date']}")
                        st.error(f"💬 **Director Comments:** {req.get('director_comments', 'None')}")
                        display_attachments(req)
                        st.divider()
                        
                        # ✅ CHANGE STATUS — REJECTED → PENDING / APPROVED
                        with st.form(f"change_status_rejected_{req['id']}"):
                            st.subheader("🔧 Change Status")
                            comments = st.text_area("💬 Updated Comments (Optional)")
                            col_pending, col_approve = st.columns(2)
                            with col_pending:
                                pending_btn = st.form_submit_button("⏳ Move to Pending")
                            with col_approve:
                                approve_btn = st.form_submit_button("✅ Change to APPROVED", type="primary")
                            
                            if pending_btn:
                                update_record_status_in_excel(req["id"], "pending", comments, FULL_NAME)
                                st.info(f"⏳ Request #{req['id']} moved back to PENDING!")
                                st.rerun()
                            if approve_btn:
                                update_record_status_in_excel(req["id"], "approved", comments, FULL_NAME)
                                st.success(f"✅ Request #{req['id']} changed to APPROVED!")
                                st.rerun()
                        
                        st.divider()
                        display_pdf_button(req, can_generate=True)
    
    # ========================================================
    # 🛡️ SUPER ADMIN PORTAL
    # ========================================================
elif user["role"] == "Super Admin":
    st.subheader("🛡️ Super Admin Control Panel")
    tab_settings, tab_users = st.tabs(["⚙️ System Settings", "👤 User Management"])
    
    with tab_settings:
        settings_management_panel()   # ✅ Create / Edit / Delete Categories, Departments, Roles
    
    with tab_users:
        user_management_panel()       # ✅ Create / Edit / Delete ALL User Accounts

    # ✅ DOWNLOAD ALL DATA BACKUPS — ONLY SUPER ADMIN CAN SEE THIS!
    st.subheader("📥 Download Data Backups")
    st.download_button("📥 Download Requests", ...)   # Excel backup
    st.download_button("📥 Download Users", ...)      # User list backup
    st.download_button("📥 Download Settings", ...)    # Settings backup
     # ========================================================
    # 📥 DOWNLOAD BACKUPS — SUPER ADMIN ONLY
    # ========================================================
    if user["role"] == "Super Admin":
        st.divider()
        st.subheader("📥 Download Data Backups")
        backup_col1, backup_col2, backup_col3 = st.columns(3)
        with backup_col1:
            if os.path.exists(EXCEL_PATH):
                with open(EXCEL_PATH, "rb") as f:
                    st.download_button(
                        "📥 Download Requests",
                        f.read(),
                        file_name=f"BACKUP_requests_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                        type="primary"
                    )
        with backup_col2:
            if os.path.exists(USER_DB_PATH):
                with open(USER_DB_PATH, "rb") as f:
                    st.download_button(
                        "📥 Download Users",
                        f.read(),
                        file_name=f"BACKUP_users_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                        type="primary"
                    )
        with backup_col3:
            if os.path.exists(SETTINGS_PATH):
                with open(SETTINGS_PATH, "rb") as f:
                    st.download_button(
                        "📥 Download Settings",
                        f.read(),
                        file_name=f"BACKUP_settings_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                        type="primary"
                    )
        st.caption("💾 Save these files to your computer for backup")

# ============================================================
# ✅ END OF FILE — NOTHING AFTER THIS!
# ============================================================
