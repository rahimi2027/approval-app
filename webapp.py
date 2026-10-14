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
TRAINING_CHAT_PATH = os.path.join(APP_FOLDER, "training_support_chat.xlsx")
TRAINING_CHAT_COLUMNS = [
    "Message_ID", "Conversation_ID", "Timestamp", "Sender_Username", "Sender_Name",
    "Sender_Role", "Sender_Department", "Recipient_Username", "Recipient_Name",
    "Message", "Read_By_Recipient"
]
TRAINING_SUPPORT_USERNAME = "wais"
TRAINING_SUPPORT_CONTACT = "Wais Rahimi"
# Optional bundled user database used to bootstrap a fresh/empty APP_FOLDER.
# Once copied into APP_FOLDER, the persistent APP_FOLDER copy is authoritative.
BUNDLED_USER_DB_PATH = os.path.join(BASE_DIR, "user_database.xlsx")
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
