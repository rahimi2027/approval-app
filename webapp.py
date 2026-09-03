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
# PDF GENERATION — ✅ FIXED: Reads LIVE Excel status + Stamp appears
# ============================================================
def generate_approval_pdf(request_data):
    import os
    import io
    if not PDF_AVAILABLE:
        return False, None, "Install fpdf2: pip install fpdf2"
    try:
        # ✅ ALWAYS READ FRESH DATA DIRECTLY FROM EXCEL — FIXES "PENDING" BUG!
        req_id = request_data.get("id")
        all_recs = load_records_from_excel()
        fresh_data = next((r for r in all_recs if int(r["id"]) == int(req_id)), request_data)
        
        def clean_text(t):
            return str(t).replace("—", "-").replace("–", "-").replace(":", "-").replace("/", "-").replace("\\", "-").strip()
        
        def format_date(d):
            if not d or str(d).strip() == "" or str(d).strip().lower() == "none":
                return "—"
            return str(d).strip()[:10]
        
        # ✅ USE FRESH STATUS FROM EXCEL — THIS IS THE MAIN FIX!
        fresh_status = str(fresh_data.get("status", "pending")).strip().lower()
        approved_by = clean_text(fresh_data.get("decision_by", fresh_data.get("approved_by", "")))
        approved_date = format_date(fresh_data.get("decision_date", fresh_data.get("approved_date", "")))
        
        emp_name = clean_text(fresh_data.get("emp_name", "Unknown"))
        category = clean_text(fresh_data.get("category", "-"))
        amount = clean_text(fresh_data.get("amount", "-"))
        reason = clean_text(fresh_data.get("desc", fresh_data.get("reason", "-")))
        dept = clean_text(fresh_data.get("dept", fresh_data.get("department", "-")))
        date_submitted = format_date(fresh_data.get("date", fresh_data.get("submitted_date", "")))
        transaction_type = clean_text(fresh_data.get("type", "-"))
        line_manager = clean_text(fresh_data.get("manager", "-"))
        
        # ✅ CREATE PDF
        from fpdf2 import FPDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        
        # ─── HEADER: LOGO + COMPANY NAME ──────────────────────────
        LOGO_PATH_PDF = os.path.join(BASE_DIR, "logo.png")
        if os.path.exists(LOGO_PATH_PDF):
            pdf.image(LOGO_PATH_PDF, x=15, y=12, w=50)
        pdf.ln(6)
        pdf.set_font("Courier", "B", 16)
        pdf.cell(0, 10, txt="ACOOLE ELECTRICAL LTD", ln=True, align="C")
        pdf.set_font("Courier", "", 11)
        pdf.cell(0, 8, txt="Addition & Deduction - APPROVAL FORM", ln=True, align="C")
        pdf.ln(4)
        
        # ─── DOUBLE LINE ───────────────────────────────────────────
        pdf.set_draw_color(0, 0, 0)
        pdf.line(20, pdf.get_y(), 190, pdf.get_y())
        pdf.ln(1)
        pdf.line(20, pdf.get_y(), 190, pdf.get_y())
        pdf.ln(12)
        
        # ─── REQUEST DETAILS ──────────────────────────────────────
        pdf.set_font("Courier", "B", 11)
        pdf.cell(0, 8, txt="REQUEST DETAILS", ln=True)
        pdf.ln(5)
        pdf.set_font("Courier", "", 11)
        label_w = 55
        pdf.cell(label_w, 7, txt="Request ID:", border=0)
        pdf.cell(0, 7, txt=f"#{req_id}", ln=True)
        pdf.cell(label_w, 7, txt="Employee Name:", border=0)
        pdf.cell(0, 7, txt=emp_name, ln=True)
        pdf.cell(label_w, 7, txt="Department:", border=0)
        pdf.cell(0, 7, txt=dept, ln=True)
        pdf.cell(label_w, 7, txt="Transaction Type:", border=0)
        pdf.cell(0, 7, txt=transaction_type, ln=True)
        pdf.cell(label_w, 7, txt="Category / Reason:", border=0)
        pdf.cell(0, 7, txt=category, ln=True)
        pdf.cell(label_w, 7, txt="Request Date:", border=0)
        pdf.cell(0, 7, txt=date_submitted, ln=True)
        pdf.cell(label_w, 7, txt="Amount Approved:", border=0)
        pdf.cell(0, 7, txt=f"£{amount}", ln=True)
        pdf.cell(label_w, 7, txt="Line Manager:", border=0)
        pdf.cell(0, 7, txt=line_manager, ln=True)
        pdf.ln(10)
        
        # ─── DESCRIPTION ──────────────────────────────────────────
        pdf.set_font("Courier", "B", 11)
        pdf.cell(0, 8, txt="DESCRIPTION / JUSTIFICATION", ln=True)
        pdf.set_font("Courier", "", 11)
        pdf.multi_cell(0, 7, txt=reason)
        pdf.ln(10)
        
        # ─── DIRECTOR APPROVAL — ✅ FRESH STATUS ───────────────────
        pdf.set_font("Courier", "B", 11)
        pdf.cell(0, 8, txt="DIRECTOR APPROVAL", ln=True)
        pdf.ln(5)
        pdf.set_font("Courier", "", 11)
        
        if fresh_status == "approved":
            decision_text = "APPROVED"
            pdf.set_font("Courier", "B", 11)
            pdf.set_text_color(0, 100, 0)
        elif fresh_status == "rejected":
            decision_text = "REJECTED"
            pdf.set_font("Courier", "B", 11)
            pdf.set_text_color(180, 0, 0)
        else:
            decision_text = "PENDING"
            pdf.set_font("Courier", "", 11)
            pdf.set_text_color(100, 100, 100)
        
        pdf.cell(label_w, 7, txt="Decision:", border=0)
        pdf.cell(0, 7, txt=decision_text, ln=True)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Courier", "", 11)
        
        pdf.cell(label_w, 7, txt="Approved By:", border=0)
        display_name = approved_by if approved_by and approved_by != "" else "—"
        pdf.cell(0, 7, txt=display_name, ln=True)
        pdf.cell(label_w, 7, txt="Approval Date / Time:", border=0)
        display_date = approved_date if approved_date and approved_date != "—" else "—"
        pdf.cell(0, 7, txt=display_date, ln=True)
        pdf.ln(22)
        
        # ─── DIRECTOR COMMENTS ─────────────────────────────────────
        prev_comments = fresh_data.get("director_comments", "").strip()
        if prev_comments:
            pdf.ln(5)
            pdf.set_font("Courier", "B", 10)
            pdf.set_text_color(80, 80, 80)
            pdf.cell(0, 7, txt="Director Comments:", ln=True)
            pdf.set_font("Courier", "", 10)
            pdf.multi_cell(0, 6, txt=prev_comments)
            pdf.set_text_color(0, 0, 0)
            pdf.ln(5)
        
        # ─── SIGNATURE LINE + APPROVED STAMP ✅ FIXED PATH ─────────
        signature_y = pdf.get_y()
        pdf.set_font("Courier", "", 10)
        pdf.set_text_color(80, 80, 80)
        pdf.dashed_line(20, signature_y, 190, signature_y, 2, 2)
        pdf.ln(5)
        pdf.cell(0, 7, txt="Authorized Signature / Director", ln=True)
        
        # ✅ STAMP — NOW USES CORRECT PATH
        if fresh_status == "approved":
            stamp_path = os.path.join(BASE_DIR, "approved_stamp.png")
            if os.path.exists(stamp_path):
                stamp_width = 55
                page_width = 210
                stamp_x = (page_width - stamp_width) / 2
                pdf.image(stamp_path, x=stamp_x, y=signature_y - 18, w=stamp_width, h=55)
            else:
                # Fallback text if image missing
                pdf.set_font("Courier", "B", 30)
                pdf.set_text_color(0, 150, 0)
                pdf.text(85, signature_y + 10, txt="APPROVED")
                pdf.set_text_color(0, 0, 0)
        
        # ─── OUTPUT ────────────────────────────────────────────────
        filename = f"ID No.{req_id} - {emp_name} - {category}.pdf"
        pdf_bytes = io.BytesIO()
        pdf.output(pdf_bytes)
        pdf_bytes.seek(0)
        return True, pdf_bytes.read(), filename
    except Exception as e:
        import traceback
        return False, None, f"PDF Error: {str(e)}"
# ============================================================
# DIRECTOR STATUS SWITCH — DEBUG VERSION
# ============================================================
def director_switch_status(request_data):
    """Director can flip status — shows Excel columns to fix 'id' error"""
    
    st.markdown("---")
    st.warning("🔧 ⚙️ DIRECTOR STATUS SWITCH PANEL")
    
    # ✅ Show what keys are in the request data
    #st.info(f"📋 Request data keys: {list(request_data.keys())}")
    
    # ✅ Try ALL possible id column names
    req_id = None
    for key in ["id", "ID", "request_id", "Id", "RequestID"]:
        if key in request_data and request_data[key]:
            req_id = request_data[key]
            break
    
    if not req_id:
        st.error("❌ Cannot find ID in request data!")
        return

    current_status = str(request_data.get("status", request_data.get("Status", "pending"))).strip().lower()

    col1, col2 = st.columns(2)
    with col1:
        new_status = st.selectbox(
            "Change Status To:",
            ["pending", "approved", "rejected"],
            index=["pending", "approved", "rejected"].index(current_status) if current_status in ["pending", "approved", "rejected"] else 0,
            key=f"switch_status_{req_id}"
        )
    with col2:
        confirm = st.checkbox("✅ I confirm this change", key=f"confirm_switch_{req_id}")

    if st.button("🔄 UPDATE STATUS", type="primary", disabled=not confirm, key=f"btn_switch_{req_id}"):
        if new_status == current_status:
            st.info("ℹ️ Status is already set to that — no change made.")
            return

        try:
            import pandas as pd
            from datetime import datetime

            excel_path = "requests.xlsx"
            df = pd.read_excel(excel_path)

            # ✅ Show Excel column names for debugging
            st.info(f"📊 Excel columns: {list(df.columns)}")

            # ✅ Try matching with ANY id column name
            id_col = None
            for possible_id in ["id", "ID", "request_id", "Id", "RequestID"]:
                if possible_id in df.columns:
                    id_col = possible_id
                    break

            if not id_col:
                st.error("❌ No 'id' column found in Excel!")
                return

            mask = df[id_col].astype(str) == str(req_id)
            if not mask.any():
                st.error(f"❌ Request #{req_id} NOT FOUND in Excel!")
                return

            # ✅ Update status
            status_col = "status" if "status" in df.columns else "Status"
            df.loc[mask, status_col] = new_status

            # ✅ Update decision fields
            if "decision_by" in df.columns:
                df.loc[mask, "decision_by"] = st.session_state.username
            if "decision_date" in df.columns:
                df.loc[mask, "decision_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            df.to_excel(excel_path, index=False)

            st.success(f"✅ ✅ Status CHANGED: **{current_status.upper()} → {new_status.upper()}**")
            st.info(f"📄 Saved to: {excel_path}")
            st.rerun()

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")

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
    # PAYROLL PORTAL
    # ========================================================
    if user["role"] == "Payroll":
        st.subheader("💰 Payroll Department Portal")
        st.info("✅ View approved & pending requests, review details, generate/download PDFs.")
        st.divider()
        tab_approved, tab_pending = st.tabs(["✅ Approved Requests", "🟡 Pending Requests"])

        with tab_approved:
            approved = [r for r in all_live_requests if r["status"] == "approved"]
            if not approved:
                st.info("📋 No approved requests yet.")
            else:
                st.metric("✅ Total Approved", len(approved))
                st.divider()
                for req in reversed(approved):
                    approved_by_display = req.get("decision_by", "Director")
                    approved_date_display = format_date(req.get("decision_date", ""))
                    if approved_date_display and approved_date_display != "—":
                        approved_by_line = f"✅ Approved by {approved_by_display} on {approved_date_display}"
                    else:
                        approved_by_line = f"✅ Approved by {approved_by_display}"
                    title = f"🟢 ID #{req['id']} | {req['emp_name']} | £{req['amount']:.2f} | {approved_by_line}"
                    with st.expander(title):
                        st.write(f"👤 Employee: {req['emp_name']}")
                        st.write(f"🏢 Department: {req['dept']}")
                        st.write(f"🔄 Type: {req['type']} | 🏷️ Category: {req['category']}")
                        st.write(f"💷 Amount: £{req['amount']:.2f}")
                        st.write(f"👔 Line Manager: {req['manager']}")
                        st.write(f"📅 Request Date: {req['date']}")
                        st.info(f"📝 Description: {req['desc']}")
                        if req["director_comments"]:
                            st.success(f"💬 Director Comments: {req['director_comments']}")
                        display_attachments(req)
                        st.divider()
                        display_pdf_button(req, can_generate=True, key_suffix="pay_approved")

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
                        
                        # ✅ PREVIOUS REJECTION COMMENT — NOW VISIBLE TO DIRECTOR!
                        prev_comments = req.get("director_comments", "").strip()
                        if prev_comments:
                            st.warning(f"📌 PREVIOUS REJECTION REASON:\n\n{prev_comments}")
                        
                        st.divider()
                        with st.form(f"decision_form_{req['id']}"):
                            comments = st.text_area("💬 Director Comments (Optional)")
                            col_approve, col_reject = st.columns(2)
                            with col_approve:
                                approve_btn = st.form_submit_button("✅ APPROVE", type="primary")
                            with col_reject:
                                reject_btn = st.form_submit_button("❌ REJECT", type="secondary")
                            if approve_btn:
                                update_record_status_in_excel(req["id"], "approved", comments, FULL_NAME)
                                st.success(f"✅ Request #{req['id']} APPROVED!")
                                st.rerun()
                            if reject_btn:
                                update_record_status_in_excel(req["id"], "rejected", comments, FULL_NAME)
                                st.warning(f"❌ Request #{req['id']} REJECTED!")
                                st.rerun()
                        st.divider()
                        display_pdf_button(req, can_generate=True)
                        director_switch_status(req)

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
                        old = str({"emp_name": rec["emp_name"], "dept": rec["dept"], "type": rec["type"], 
                                   "category": rec["category"], "date": rec["date"], "amount": rec["amount"], 
                                   "manager": rec["manager"], "desc": rec["desc"]})
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
                                
                                old_comments = r.get("director_comments", "").strip()
                                if old_comments and old_comments != "":
                                    if "📌 PREVIOUS REJECTION REASON:" not in old_comments:
                                        r["director_comments"] = f"📌 PREVIOUS REJECTION REASON:\n{old_comments}\n\n--- NEW REQUEST ---"
                                else:
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
                        st.success(f"✅ Updated & sent for approval! Previous rejection reason saved.")
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
                    if req.get("director_comments", "").strip():
                        st.warning(f"📌 **PREVIOUS COMMENT / REJECTION REASON:**\n\n{req['director_comments']}")
                    if req["status"] == "approved":
                        display_pdf_button(req, can_generate=False)
                    if req["status"] in ["pending", "rejected"]:
                        if st.button(f"✏️ Edit Request #{req['id']}", key=f"edit_{req['id']}"):
                            st.session_state.editing_request_id = req["id"]
                            st.rerun()

      # ========================================================
    # 🎛️ DIRECTOR PORTAL
    # ========================================================
    elif user["role"] == "Director":
        st.subheader("🎛️ Director Approval Portal — Andy Acoole")
        st.info("✅ Review all requests, Approve or Reject. Decisions update automatically.")
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
                        with st.form(f"decision_form_{req['id']}"):
                            comments = st.text_area("💬 Director Comments (Optional)")
                            col_approve, col_reject = st.columns(2)
                            with col_approve:
                                approve_btn = st.form_submit_button("✅ APPROVE", type="primary")
                            with col_reject:
                                reject_btn = st.form_submit_button("❌ REJECT", type="secondary")
                            if approve_btn:
                                update_record_status_in_excel(req["id"], "approved", comments, FULL_NAME)
                                st.success(f"✅ Request #{req['id']} APPROVED!")
                                st.rerun()
                            if reject_btn:
                                update_record_status_in_excel(req["id"], "rejected", comments, FULL_NAME)
                                st.warning(f"❌ Request #{req['id']} REJECTED!")
                                st.rerun()
                        st.divider()
                        display_pdf_button(req, can_generate=True)
                        director_switch_status(req)
        
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
                        display_pdf_button(req, can_generate=True)
                        director_switch_status(req)
        
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
                        display_pdf_button(req, can_generate=True)
                        director_switch_status(req)
    
    # ========================================================
    # 🛡️ SUPER ADMIN PORTAL
    # ========================================================
    elif user["role"] == "Super Admin":
        st.subheader("🛡️ Super Admin Control Panel")
        tab_settings, tab_users = st.tabs(["⚙️ System Settings", "👤 User Management"])
        with tab_settings:
            settings_management_panel()
        with tab_users:
            user_management_panel()
    # ========================================================
    # 🛡️ SUPER ADMIN PORTAL
    # ========================================================
    elif user["role"] == "Super Admin":
        st.subheader("🛡️ Super Admin Control Panel")
        tab_settings, tab_users = st.tabs(["⚙️ System Settings", "👤 User Management"])
        with tab_settings:
            settings_management_panel()
        with tab_users:
            user_management_panel()
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

