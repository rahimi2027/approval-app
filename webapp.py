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
BASE_DIR = r"D:\Acoole_portal"
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
DEPARTMENTS = ["National Grid", "Isolator", "Project", "Accounts", "Payroll Department", "ACoole Electrical Ltd"]

# ============================================================
# SETTINGS FUNCTIONS — Categories & Roles
# ============================================================
def init_settings():
    if not os.path.exists(SETTINGS_PATH):
        pd.DataFrame([{
            "setting": "categories",
            "value": "|".join(DEFAULT_CATEGORIES)
        }, {
            "setting": "roles",
            "value": "|".join(DEFAULT_ROLES)
        }]).to_excel(SETTINGS_PATH, index=False)

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

def save_categories(categories_list):
    init_settings()
    df = pd.read_excel(SETTINGS_PATH).fillna("")
    found = False
    for idx, r in df.iterrows():
        if r["setting"] == "categories":
            df.at[idx, "value"] = "|".join(categories_list)
            found = True
    if not found:
        df = pd.concat([df, pd.DataFrame([{"setting": "categories", "value": "|".join(categories_list)}])], ignore_index=True)
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

# ============================================================
# USER DATABASE FUNCTIONS
# ============================================================
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
# REQUESTS EXCEL COLUMNS
# ============================================================
EXCEL_COLUMNS = [
    "ID", "Employee Name", "Department", "Transaction Type", "Category Reason",
    "Date", "Amount (£)", "Line Manager", "Description", "Attachment Name",
    "Status", "Director Comments", "Decision Date", "Decision By",
    "PDF File Path", "Edited From ID", "Old Data"
]

# ============================================================
# INITIALISE REQUESTS EXCEL
# ============================================================
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

# ============================================================
# LOAD / SAVE REQUESTS
# ============================================================
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
# PDF GENERATION
# ============================================================
def generate_approval_pdf(request_data):
    if not PDF_AVAILABLE:
        return False, "Install fpdf2: pip install fpdf2", None
    try:
        # ✅ READ FRESH DATA DIRECTLY FROM EXCEL — BYPASS OLD CACHED DATA!
        req_id = request_data.get("id")
        all_recs = load_records_from_excel()
        fresh_data = next((r for r in all_recs if int(r["id"]) == int(req_id)), request_data)
        
        def clean_text(t):
            return str(t).replace("—", "-").replace("–", "-")
        
        # ✅ USE FRESH DATA FROM EXCEL — NOT OLD CACHED DATA!
        emp_name_safe = clean_text(fresh_data.get("emp_name", "Unknown"))
        category_safe = clean_text(fresh_data.get("category", "Approval"))
        amount = f"£{float(fresh_data.get('amount', 0)):.2f}"  # ✅ FRESH AMOUNT!
        req_date = clean_text(str(fresh_data.get("date", "Unknown")))
        decision_date = clean_text(str(fresh_data.get("decision_date", "Not Approved Yet")))
        comment_safe = clean_text(fresh_data.get("director_comments", ""))
        desc_safe = clean_text(fresh_data.get("desc", ""))
        dept_safe = clean_text(fresh_data.get("dept", ""))
        manager_safe = clean_text(fresh_data.get("manager", ""))
        decision_by_safe = clean_text(fresh_data.get("decision_by", ""))
        
        safe_emp = "".join(c for c in emp_name_safe if c.isalnum() or c in (" ", "-", "_")).strip()
        safe_cat = "".join(c for c in category_safe if c.isalnum() or c in (" ", "-", "_")).strip()
        safe_date = req_date.replace("/", "-").replace(":", "-")[:10]
        filename = f"{safe_emp} - {safe_cat} - {amount} - {safe_date}.pdf"
        full_pdf_path = os.path.join(PDF_DIR, filename)
        
        # ✅ DELETE ANY OLD PDF WITH SIMILAR NAME FIRST
        for f in os.listdir(PDF_DIR):
            if f.startswith(f"ID_{req_id}_") or f.startswith(f"{safe_emp} - {safe_cat}"):
                try:
                    os.remove(os.path.join(PDF_DIR, f))
                except:
                    pass
        
        pdf = FPDF("P", "mm", "A4")
        pdf.add_page()
        pdf.set_font("Courier", "", 12)
        if os.path.exists(LOGO_PATH):
            try: pdf.image(LOGO_PATH, x=10, y=8, w=40)
            except: pass
        pdf.set_xy(60, 10)
        pdf.set_font("Courier", "B", 18)
        pdf.cell(0, 10, "ACOOLE ELECTRICAL LTD", ln=True)
        pdf.set_font("Courier", "", 12)
        pdf.cell(0, 6, "Addition & Deduction - APPROVAL FORM", ln=True)
        pdf.ln(5)
        pdf.cell(0, 0, "_" * 85, ln=True)
        pdf.ln(8)
        def pdf_row(label, value):
            pdf.set_font("Courier", "B", 11)
            pdf.cell(60, 8, label + ":")
            pdf.set_font("Courier", "", 11)
            pdf.multi_cell(0, 8, clean_text(str(value)))
            pdf.ln(1)
        pdf.set_font("Courier", "B", 12)
        pdf.cell(0, 8, "REQUEST DETAILS", ln=True)
        pdf.ln(3)
        pdf_row("Request ID", f"#{req_id}")
        pdf_row("Employee Name", emp_name_safe)
        pdf_row("Department", dept_safe)
        pdf_row("Transaction Type", clean_text(fresh_data.get("type", "")))
        pdf_row("Category / Reason", category_safe)
        pdf_row("Request Date", req_date)
        pdf_row("Amount Approved", amount)  # ✅ FRESH AMOUNT IN PDF!
        pdf_row("Line Manager", manager_safe)
        pdf.ln(5)
        pdf.set_font("Courier", "B", 12)
        pdf.cell(0, 8, "DESCRIPTION / JUSTIFICATION", ln=True)
        pdf.set_font("Courier", "", 11)
        pdf.multi_cell(0, 7, desc_safe)
        pdf.ln(8)
        pdf.set_font("Courier", "B", 12)
        pdf.cell(0, 8, "DIRECTOR APPROVAL", ln=True)
        pdf_row("Decision", "APPROVED")
        pdf_row("Approved By", decision_by_safe)
        pdf_row("Approval Date / Time", decision_date)
        if comment_safe:
            pdf_row("Director Comments", comment_safe)
        pdf.ln(15)
        pdf.cell(0, 6, "-" * 50, ln=True)
        pdf.cell(0, 6, "Authorized Signature / Director", ln=True)
        
        att_names = fresh_data.get("attachment_name", "None")
        if att_names and att_names != "None":
            attached_files = [n.strip() for n in att_names.split(",")]
            for file_name in attached_files:
                file_path = os.path.join(UPLOAD_DIR, file_name)
                if os.path.exists(file_path):
                    pdf.add_page()
                    pdf.set_font("Courier", "B", 14)
                    pdf.cell(0, 12, "ATTACHED DOCUMENT", ln=True)
                    pdf.set_font("Courier", "", 11)
                    pdf.cell(0, 8, f"File: {file_name}", ln=True)
                    pdf.cell(0, 8, f"Request ID: #{req_id}", ln=True)
                    pdf.ln(5)
                    pdf.cell(0, 0, "_" * 60, ln=True)
                    pdf.ln(5)
                    try:
                        page_w = 190
                        pdf.image(file_path, x=10, y=pdf.get_y(), w=page_w)
                    except Exception:
                        pdf.set_font("Courier", "", 11)
                        pdf.multi_cell(0, 8, f"File saved at:\n{file_path}")
                else:
                    pdf.add_page()
                    pdf.set_font("Courier", "", 11)
                    pdf.cell(0, 10, f"Attachment not found: {file_name}", ln=True)
        
        pdf.output(full_pdf_path)
        return True, full_pdf_path, filename
    except Exception as e:
        return False, str(e), None

def save_pdf_path_to_request(req_id, pdf_path):
    records = load_records_from_excel()
    for rec in records:
        if int(rec["id"]) == int(req_id):
            rec["pdf_path"] = pdf_path
            save_all_records(records)
            return True
    return False

def update_record_status_in_excel(req_id, new_status, director_comments, decision_by):
    records = load_records_from_excel()
    for record in records:
        if int(record["id"]) == int(req_id):
            record["status"] = str(new_status).lower()
            old_comments = record["director_comments"].strip()
            new_comment_text = str(director_comments).strip()
            ts = datetime.now().strftime("%Y-%m-%d %H:%M")
            if new_comment_text:
                record["director_comments"] = f"{old_comments}\n---\n[{ts}] {decision_by}: {new_comment_text}" if old_comments else f"[{ts}] {decision_by}: {new_comment_text}"
            record["decision_date"] = ts
            record["decision_by"] = str(decision_by)
            save_all_records(records)
            return True
    return False

def delete_record_by_id(req_id):
    records = load_records_from_excel()
    new_records = [r for r in records if int(r["id"]) != int(req_id)]
    if len(new_records) < len(records):
        save_all_records(new_records)
        return True
    return False

def get_next_id(records):
    if not records: return 1
    ids = [int(r["id"]) for r in records if str(r.get("id", "")).isdigit()]
    return max(ids) + 1 if ids else 1

def display_company_header():
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        LOGO_PATH = "logo.png"
        if os.path.exists(LOGO_PATH):  
          st.image(LOGO_PATH, width=150)
        else:
         st.title("⚡ ACOOLE ELECTRICAL LTD")
        st.caption("Addition & Deduction Approval Platform")
        st.divider()

def display_attachments(req):
    att = req.get("attachment_name", "None")
    if not att or att == "None":
        st.info("📎 No attachments.")
        return
    for idx, name in enumerate([n.strip() for n in att.split(",")]):
        path = os.path.join(UPLOAD_DIR, name)
        if os.path.exists(path):
            with open(path, "rb") as f:
                st.download_button(f"⬇️ Download {name}", f.read(), file_name=name, key=f"att_{req['id']}_{idx}")
        else:
            st.warning(f"⚠️ Not found: {name}")

def display_pdf_button(req, can_generate=False):
    req_id = req["id"]
    pdf_path = req.get("pdf_path", "")
    if pdf_path and os.path.exists(pdf_path):
        with open(pdf_path, "rb") as f:
            st.download_button(f"📥 Download PDF", f.read(), file_name=os.path.basename(pdf_path), key=f"pdf_{req_id}")
        return True
    elif can_generate:
        if st.button(f"📄 Generate PDF for ID #{req_id}", type="primary", key=f"genpdf_{req_id}"):
            ok, path, name = generate_approval_pdf(req)
            if ok:
                save_pdf_path_to_request(req_id, path)
                st.success(f"✅ Generated: {name}")
                st.rerun()
            else:
                st.error(f"❌ Error: {path}")
    return False

def show_old_new_comparison(old_data, new_data):
    if not old_data or old_data == "{}": return
    try:
        old = eval(old_data) if isinstance(old_data, str) else old_data
        st.warning("🔄 EDITED — Changes:")
        fields = [("Employee Name", "emp_name"), ("Dept", "dept"), ("Type", "type"), ("Category", "category"), ("Date", "date"), ("Amount", "amount"), ("Manager", "manager"), ("Desc", "desc")]
        for label, key in fields:
            o, n = str(old.get(key, "")), str(new_data.get(key, ""))
            if o != n:
                st.markdown(f"**{label}:** ~~{o}~~ → **{n}**")
        st.divider()
    except: pass

def format_date(d):
    return str(d).strip()[:10] if d and str(d).strip() else "—"

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
    if st.button("🔄 Refresh Data", type="secondary"):
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
                    if rec["status"] == "approved" and not rec.get("pdf_path"):
                        ok, path, _ = generate_approval_pdf(rec)
                        if ok:
                            rec["pdf_path"] = path
                            count += 1
                if count > 0:
                    save_all_records(all_live_requests)
                    st.sidebar.success(f"✅ Generated {count} PDFs!")
                    st.rerun()
                else:
                    st.sidebar.info("✅ No new PDFs needed.")
            else:
                st.sidebar.error("⚠️ Install fpdf2: pip install fpdf2")

# ============================================================
# ⚙️ SUPER ADMIN SETTINGS PANEL
# ============================================================
def settings_management_panel():
    st.subheader("⚙️ System Settings — Categories & Roles")
    st.info("🛡️ Super Admin Only — Add, Edit, Delete Categories and Permission Roles.")
    st.divider()
    cats_tab, roles_tab = st.tabs(["🏷️ Manage Categories", "🎖️ Manage Roles / Permissions"])
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
            with c1:
                st.markdown(f"• **{cat}**")
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
            with c1:
                st.markdown(f"• **{role}**")
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
# 👤 USER MANAGEMENT PANEL
# ============================================================
def user_management_panel():
    st.subheader("👤 User Management — Create & Manage System Users")
    st.info("🛡️ Super Admin Only — Create, edit (Username & Full Name), or delete user accounts.")
    st.divider()
    USERS = load_users()
    ROLES = load_roles()
    tab1, tab2, tab3 = st.tabs(["➕ Create New User", "✏️ Edit User", "🗑️ Delete User"])
    with tab1:
        st.markdown("### ➕ Create New System User")
        with st.form("create_user_form", border=True, clear_on_submit=True):
            new_full_name = st.text_input("👤 Full Name", placeholder="e.g. John Smith")
            new_username = st.text_input("🔐 Username (Login ID)", placeholder="e.g. john_smith").lower().strip()
            new_password = st.text_input("🔑 Password", type="password")
            new_role = st.selectbox("🎖️ Role / Permission Level", ROLES)
            new_dept = st.selectbox("🏢 Department", DEPARTMENTS)
            st.markdown("#### 📋 Role Permissions:")
            st.info("""
            - **Manager** → Submit & edit requests for their department only
            - **Director** → Approve / reject ALL requests, change status anytime
            - **Payroll** → View approved requests, generate PDFs
            - **Super Admin** → Full system access + Create/edit/delete users & settings
            """)
            create_btn = st.form_submit_button("✅ Create User Account", type="primary")
            if create_btn:
                if not new_full_name.strip() or not new_username or not new_password:
                    st.error("❌ All fields required! (Full Name, Username, Password)")
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
                    st.success(f"✅ User **'{new_full_name}'** created!\n🔐 Username: **{new_username}** | Role: **{new_role}** | Dept: **{new_dept}**")
                    st.balloons()
    with tab2:
        st.markdown("### ✏️ Edit User — Change Everything")
        edit_user_sel = st.selectbox("Select User to Edit", list(USERS.keys()), key="edit_user_selector")
        if edit_user_sel:
            curr = USERS[edit_user_sel]
            st.info(f"Current: **{curr.get('full_name', edit_user_sel)}** | Login: `{edit_user_sel}` | {curr['role']} | {curr['dept']}")
            with st.form(f"edit_user_form_{edit_user_sel}", border=True, clear_on_submit=True):
                upd_full_name = st.text_input("👤 Full Name", value=curr.get("full_name", edit_user_sel))
                upd_username_new = st.text_input("🔐 Change Username (Login ID)", value=edit_user_sel).lower().strip()
                upd_password = st.text_input("🔑 New Password (leave blank to keep current)", type="password")
                upd_role = st.selectbox("🎖️ Change Role / Permission Level", ROLES, index=ROLES.index(curr["role"]) if curr["role"] in ROLES else 0)
                upd_dept = st.selectbox("🏢 Change Department", DEPARTMENTS, index=DEPARTMENTS.index(curr["dept"]))
                if st.form_submit_button("🔄 Update All Details", type="primary"):
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
                        st.success(f"✅ Username changed from **'{edit_user_sel}'** → **'{upd_username_new}'**")
                    else:
                        USERS[edit_user_sel]["full_name"] = upd_full_name.strip()
                        if upd_password:
                            USERS[edit_user_sel]["password"] = upd_password
                        USERS[edit_user_sel]["role"] = upd_role
                        USERS[edit_user_sel]["dept"] = upd_dept
                    save_users(USERS)
                    st.success(f"✅ User updated: **{upd_full_name}** | Role: **{upd_role}** | Dept: **{upd_dept}**")
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
# 🔐 LOGIN PAGE
# ============================================================
if not st.session_state.logged_in:
    display_company_header()
    _, c, _ = st.columns([1, 2, 1])
    with c:
        with st.form("login_form", border=True):
            st.markdown("### 🔒 Secure Gateway Login")
            st.caption("Enter your credentials to access the system")
            st.divider()
            username = st.text_input("🔐 Username", placeholder="e.g. andy, payroll, wais").lower().strip()
            password = st.text_input("🔑 Password", type="password", placeholder="Enter your password here")
            st.markdown("<br>", unsafe_allow_html=True)
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
# ✅ MAIN APPLICATION — ALL ROLES
# ============================================================
else:
    user = st.session_state.user_info
    CAN_GEN_PDF = user["role"] in ["Payroll", "Super Admin"]
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
    # 💰 PAYROLL PORTAL
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
                    title = f"🟢 ID #{req['id']} | {req['emp_name']} | 📅 {format_date(req['date'])} | £{req['amount']:.2f} | {approved_by_line}"
                    with st.expander(title):
                        st.write(f"👤 **Employee:** {req['emp_name']}")
                        st.write(f"🏢 **Department:** {req['dept']}")
                        st.write(f"🔄 **Type:** {req['type']} | 🏷️ **Category:** {req['category']}")
                        st.write(f"💷 **Amount:** £{req['amount']:.2f}")
                        st.write(f"👔 **Line Manager:** {req['manager']}")
                        st.write(f"📅 **Request Date:** {req['date']}")
                        st.write(f"🕒 **Decision:** {approved_by_line}")
                        st.info(f"📝 **Description:** {req['desc']}")
                        if req["director_comments"]:
                            st.success(f"💬 **Director Comments:** {req['director_comments']}")
                        display_attachments(req)
                        st.divider()
                        display_pdf_button(req, can_generate=True)
        with tab_pending:
            pending = [r for r in all_live_requests if r["status"] == "pending"]
            if not pending:
                st.info("✅ No pending requests.")
            else:
                st.metric("🟡 Pending Requests", len(pending))
                st.divider()
                for req in reversed(pending):
                    title = f"🟡 ID #{req['id']} | {req['emp_name']} | 📅 {format_date(req['date'])} | £{req['amount']:.2f} | ⏳ Pending"
                    with st.expander(title):
                        st.write(f"👤 **Employee:** {req['emp_name']}")
                        st.write(f"🏢 **Department:** {req['dept']}")
                        st.write(f"🔄 **Type:** {req['type']} | 🏷️ **Category:** {req['category']}")
                        st.write(f"💷 **Amount:** £{req['amount']:.2f}")
                        st.write(f"👔 **Line Manager:** {req['manager']}")
                        st.write(f"📅 **Request Date:** {req['date']}")
                        st.info(f"📝 **Description:** {req['desc']}")
                        st.warning("⏳ Waiting for Director Approval")
                        display_attachments(req)

    # ========================================================
    # 🛡️ SUPER ADMIN PORTAL
    # ========================================================
    elif user["role"] == "Super Admin":
        st.subheader("🛡️ Super Admin Control Centre")
        tab_users, tab_settings, tab_requests = st.tabs(["👤 User Management", "⚙️ System Settings", "📋 All Requests"])
        with tab_users:
            user_management_panel()
        with tab_settings:
            settings_management_panel()
        with tab_requests:
            # ===== SUPER ADMIN EDIT FORM =====
            if st.session_state.editing_request_id:
                eid = st.session_state.editing_request_id
                rec = next((r for r in all_live_requests if int(r["id"]) == int(eid)), None)
                if rec:
                    st.subheader(f"✏️ Edit Request #{eid} — Super Admin")
                    show_old_new_comparison("{}", rec)
                    with st.form("sa_edit_form"):
                        c1, c2 = st.columns(2)
                        with c1:
                            en = st.text_input("👤 Employee Name", rec["emp_name"])
                            rt = st.selectbox("🔄 Transaction Type", ["Addition", "Deduction"], index=["Addition", "Deduction"].index(rec["type"]))
                            ct = st.selectbox("🏷️ Category / Reason", CATEGORIES, index=CATEGORIES.index(rec["category"]) if rec["category"] in CATEGORIES else 0)
                            amt = st.number_input("💷 Amount (£)", min_value=0.01, step=10.0, value=float(rec["amount"]))
                        with c2:
                            from datetime import datetime as dt
                            try:
                                d = dt.strptime(rec["date"][:10], "%Y-%m-%d")
                            except:
                                d = dt.today()
                            dt_val = st.date_input("📅 Date", d)
                            mgr = st.text_input("👔 Line Manager", rec["manager"])
                            dept_sel = st.selectbox("🏢 Department", DEPARTMENTS, index=DEPARTMENTS.index(rec["dept"]) if rec["dept"] in DEPARTMENTS else 0)
                            desc = st.text_area("📝 Description / Justification", rec["desc"])
                            files = st.file_uploader("📎 Add Supporting Documents", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)
                        
                        st.divider()
                        new_status = st.selectbox("🔄 Change Status", ["pending", "approved", "rejected"], 
                            index=["pending", "approved", "rejected"].index(rec["status"]) if rec["status"] in ["pending", "approved", "rejected"] else 0)
                        admin_comments = st.text_area("💬 Comments / Notes")
                        
                        if st.form_submit_button("💾 Save All Changes", type="primary"):
                            # ✅ GET ALL NEW VALUES FROM FORM — CLEAR & DIRECT
                            new_emp_name = str(en.strip())
                            new_dept = str(dept_sel)
                            new_type = str(rt)
                            new_category = str(ct)
                            new_amount = round(float(amt), 2)  # ✅ FORCE 2 DECIMALS
                            new_date = str(dt_val)
                            new_manager = str(mgr.strip())
                            new_desc = str(desc.strip())
                            new_status_val = str(new_status)
                            
                            old = str({
                                "emp_name": rec["emp_name"], "dept": rec["dept"], 
                                "type": rec["type"], "category": rec["category"], 
                                "date": rec["date"], "amount": rec["amount"], 
                                "manager": rec["manager"], "desc": rec["desc"]
                            })
                            
                            # ✅ LOAD, UPDATE, AND SAVE — VERY CLEAR
                            records = load_records_from_excel()
                            updated = False
                            for r in records:
                                if int(r["id"]) == int(eid):
                                    # ✅ WRITE EVERY VALUE EXPLICITLY
                                    r["emp_name"] = new_emp_name
                                    r["dept"] = new_dept
                                    r["type"] = new_type
                                    r["category"] = new_category
                                    r["amount"] = new_amount  # ✅ DIRECT ASSIGNMENT
                                    r["date"] = new_date
                                    r["manager"] = new_manager
                                    r["desc"] = new_desc
                                    r["status"] = new_status_val
                                    r["old_data"] = old
                                    
                                    # ✅ DELETE OLD PDF & CLEAR PATH
                                    old_pdf_path = r.get("pdf_path", "")
                                    if old_pdf_path and os.path.exists(old_pdf_path):
                                        try:
                                            os.remove(old_pdf_path)
                                        except:
                                            pass
                                    r["pdf_path"] = ""
                                    
                                    if admin_comments.strip():
                                        ts = datetime.now().strftime("%Y-%m-%d %H:%M")
                                        r["director_comments"] = f"[{ts}] Super Admin Edit: {admin_comments.strip()}"
                                    
                                    # ✅ SAVE NEW ATTACHMENTS
                                    if files:
                                        att_list = []
                                        if r["attachment_name"] and r["attachment_name"] != "None":
                                            att_list.extend([n.strip() for n in r["attachment_name"].split(",")])
                                        for i, f in enumerate(files):
                                            fn = f"SA_ID_{eid}_EDIT_F{len(att_list)+1}_{f.name}"
                                            with open(os.path.join(UPLOAD_DIR, fn), "wb") as out:
                                                out.write(f.getbuffer())
                                            att_list.append(fn)
                                        r["attachment_name"] = ", ".join(att_list) if att_list else "None"
                                    
                                    updated = True
                                    break
                            
                            if updated:
                                save_all_records(records)
                                
                                # ✅ READ BACK AND SHOW EXACTLY WHAT IS IN EXCEL
                                verify = load_records_from_excel()
                                check = next((x for x in verify if int(x["id"]) == int(eid)), None)
                                if check:
                                    st.info(f"📊 **Excel ACTUAL Values:** Amount = **£{check['amount']:.2f}**, Name = **{check['emp_name']}**, Date = **{check['date']}**")
                                
                                st.success(f"✅ UPDATED! Now click **📄 Generate PDF** → it will use £{new_amount:.2f}")
                                st.session_state.editing_request_id = None
                                st.rerun()
                            else:
                                st.error("❌ Could not find request to update!")
                    

            # ===== SHOW ALL REQUESTS LIST =====
            if not all_live_requests:
                st.info("📋 No requests yet.")
            else:
                for req in reversed(all_live_requests):
                    icon = "🟡" if req["status"] == "pending" else ("🟢" if req["status"] == "approved" else "🔴")
                    title = f"{icon} ID #{req['id']} | {req['status'].upper()} | {req['emp_name']} | £{req['amount']:.2f} | {req['dept']}"
                    with st.expander(title):
                        if req.get("old_data"):
                            show_old_new_comparison(req["old_data"], req)
                        st.write(f"🏢 Dept: {req['dept']} | 🔄 Type: {req['type']} | 🏷️ Cat: {req['category']}")
                        st.write(f"📅 Date: {req['date']} | 👔 Manager: {req['manager']}")
                        st.info(f"📝 Desc: {req['desc']}")
                        display_attachments(req)
                        if req["status"] == "approved":
                            display_pdf_button(req, can_generate=True)
                        
                        col_edit, col_del = st.columns(2)
                        with col_edit:
                            if st.button(f"✏️ Edit #{req['id']}", key=f"e{req['id']}"):
                                st.session_state.editing_request_id = req["id"]
                                st.rerun()
                        with col_del:
                            if st.button(f"🗑️ Delete #{req['id']}", key=f"d{req['id']}", type="secondary"):
                                delete_record_by_id(req["id"])
                                st.success("✅ Deleted")
                                st.rerun()

    # ========================================================
    # 👔 MANAGER PORTAL
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
                        files = st.file_uploader("📎 Add Supporting Documents", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)
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
    # 🎛️ DIRECTOR PORTAL — ANDY ACOLE
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

        # ===== PENDING REQUESTS =====
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
                                update_record_status_in_excel(
                                    req["id"], "approved", comments, FULL_NAME
                                )
                                st.success(f"✅ Request #{req['id']} APPROVED!")
                                st.rerun()
                            if reject_btn:
                                update_record_status_in_excel(
                                    req["id"], "rejected", comments, FULL_NAME
                                )
                                st.warning(f"❌ Request #{req['id']} REJECTED!")
                                st.rerun()

        # ===== APPROVED REQUESTS =====
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
                        st.info(f"📝 **Description:** {req['desc']}")
                        st.success(f"💬 **Director Comments:** {req.get('director_comments', 'None')}")
                        display_attachments(req)
                        st.divider()
                        display_pdf_button(req, can_generate=True)

        # ===== REJECTED REQUESTS =====
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
                        st.info(f"📝 **Description:** {req['desc']}")
                        st.error(f"💬 **Director Comments:** {req.get('director_comments', 'None')}")
                        display_attachments(req)

# ============================================================
# END OF APPLICATION
# ============================================================
    # ========================================================