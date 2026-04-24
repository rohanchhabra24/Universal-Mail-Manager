"""
Outlook Mail Manager — Streamlit App (Windows Native)
=====================================================
Single-file Streamlit application to read and send emails
via local Windows Outlook client using COM and pywin32.

THIS SCRIPT IS DESIGNED EXCLUSIVELY FOR WINDOWS.

Requirements:
    pip install streamlit pywin32 pytz
"""

import os
import time
import math
import re
import streamlit as st
import pytz
from datetime import datetime, timedelta
import streamlit.components.v1 as components

try:
    import win32com.client
    import pythoncom
except ImportError:
    st.error("🚨 `pywin32` library not found or you are not on Windows.\nPlease run `pip install pywin32` on a Windows machine.")
    win32com = None

# -----------------------------------------------------------------------------
# [TOP BLOCK] Initialization & Config
# -----------------------------------------------------------------------------

st.set_page_config(page_title="Outlook Manager", page_icon="📬", layout="wide")

st.markdown("""
<style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    .email-row {
        padding: 10px; border-radius: 8px; margin-bottom: 5px;
        background-color: transparent; transition: background-color 0.2s;
        border-bottom: 1px solid #f0f0f0;
    }
    .email-row:hover { background-color: #EBF3FB; cursor: pointer; }
    .unread-row { border-left: 4px solid #0078D4; background-color: #f9f9f9; }
    .metric-card {
        border: 1px solid #e0e0e0; border-radius: 8px; padding: 15px;
        text-align: center; box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .metric-title { font-size: 14px; color: #666; }
    .metric-value { font-size: 24px; font-weight: bold; color: #333; }
    .badge {
        display: inline-block; padding: 2px 8px; border-radius: 12px;
        font-size: 11px; font-weight: 600; margin-right: 5px;
    }
    .badge-unread { background-color: #EBF3FB; color: #0078D4; }
    .badge-flagged { background-color: #FEF3C7; color: #D97706; }
    .badge-attach { background-color: #F3F4F6; color: #4B5563; }
    .avatar {
        display: inline-block; width: 32px; height: 32px; border-radius: 50%;
        background-color: #0078D4; color: white; text-align: center;
        line-height: 32px; font-weight: bold; font-size: 12px;
    }
    .preview-text {
        color: #777; font-size: 13px; white-space: nowrap;
        overflow: hidden; text-overflow: ellipsis; max-width: 100%; display: inline-block;
    }
    div.row-widget.stRadio > div{ flex-direction:row; align-items: stretch; }
    div.row-widget.stRadio > div[role="radiogroup"] > label[data-baseweb="radio"] {
        background-color: #f9f9f9; padding: 10px 20px; border-radius: 5px;
        margin-right: 10px; border: 1px solid #ddd;
    }
    div.row-widget.stRadio > div[role="radiogroup"] > label[data-baseweb="radio"]:has(input[checked]) {
        background-color: #0078D4; color: white; border-color: #0078D4;
    }
    div.row-widget.stRadio > div[role="radiogroup"] > label[data-baseweb="radio"] div {
        color: inherit; font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# [SESSION STATE]
# -----------------------------------------------------------------------------
def initialize_session_state():
    session_defaults = {
        "emails": [],
        "selected_email_id": None,
        "selected_email_detail": None,
        "current_page": 0,
        "active_tab": "Read Mails",
        "compose_to": "",
        "compose_cc": "",
        "compose_bcc": "",
        "compose_subject": "",
        "compose_body": "",
        "selected_folder": "Inbox",
        "total_email_count": 0,
        "last_sync": None
    }
    for key, value in session_defaults.items():
        st.session_state.setdefault(key, value)

initialize_session_state()

# -----------------------------------------------------------------------------
# [OUTLOOK COM HELPERS]
# -----------------------------------------------------------------------------
def get_outlook():
    if not win32com: return None
    try:
        # Initialize COM in this thread
        pythoncom.CoInitialize()
        return win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        st.error(f"Failed to connect to local Outlook client: {e}")
        return None

# Outlook MAPI Folder IDs
FOLDER_MAP = {
    "Deleted Items": 3,
    "Outbox": 4,
    "Sent Items": 5,
    "Inbox": 6,
    "Drafts": 16,
    "Junk Email": 23
}

def resolve_folder(namespace, folder_name):
    # Tries default MAPI standard folders first, fallback to iterating folders
    if folder_name in FOLDER_MAP:
        try:
            return namespace.GetDefaultFolder(FOLDER_MAP[folder_name])
        except:
            pass
    
    # Custom folder fallback logic over root folders
    for f in namespace.Folders:
        for sub_f in f.Folders:
            if sub_f.Name.lower() == folder_name.lower():
                return sub_f
    return None

def fetch_emails(folder_name="Inbox", search_query="", top=25, skip=0):
    outlook = get_outlook()
    if not outlook: return
    
    mapi = outlook.GetNamespace("MAPI")
    folder = resolve_folder(mapi, folder_name)
    
    if not folder:
        st.warning(f"Could not find folder specified: {folder_name}")
        st.session_state.total_email_count = 0
        st.session_state.emails = []
        return
        
    items = folder.Items
    items.Sort("[ReceivedTime]", True) # Sort by date descending
    
    if search_query:
        search_filter = f"@SQL=urn:schemas:httpmail:subject LIKE '%{search_query}%' OR urn:schemas:httpmail:sendername LIKE '%{search_query}%'"
        try:
            items = items.Restrict(search_filter)
        except Exception as e:
            st.error(f"Search failed: {e}")
    
    st.session_state.total_email_count = len(items)
    parsed_emails = []
    
    # Range handling
    end_idx = min(skip + top, len(items))
    start_idx = skip
    
    current_idx = 0
    item = items.GetFirst()
    
    while item:
        if current_idx >= end_idx:
            break
            
        if current_idx >= start_idx:
            # We must ensure it's a MailItem (Class == 43) and not a Calendar Item, etc.
            if getattr(item, "Class", None) == 43:
                try:
                    date_str = str(item.ReceivedTime)
                except:
                    date_str = ""
                    
                preview = item.Body[:100] if item.Body else ""
                clean_preview = preview.replace('\r', '').replace('\n', ' ')
                
                parsed_emails.append({
                    "id": item.EntryID,
                    "subject": item.Subject,
                    "from": item.SenderName,
                    "to": item.To,
                    "date": date_str,
                    "is_read": not item.UnRead,
                    "is_flagged": item.IsMarkedAsTask,
                    "has_attachments": item.Attachments.Count > 0,
                    "preview": clean_preview,
                    "body_html": item.HTMLBody,
                    "body_text": item.Body
                })
        
        item = items.GetNext()
        current_idx += 1

    st.session_state.emails = parsed_emails
    st.session_state.last_sync = datetime.now()

def fetch_email_detail(entry_id):
    outlook = get_outlook()
    if not outlook: return None
    mapi = outlook.GetNamespace("MAPI")
    try:
        # Fetch directly using EntryID globally
        msg = mapi.GetItemFromID(entry_id)
        
        # Prepare attachments cache (we only list sizes/names because dragging raw COM bytes requires SaveAsFile)
        attachments = []
        for att in msg.Attachments:
            # Type 1 is explicit attachment file
            if att.Type == 1:
                attachments.append({
                    "filename": att.FileName,
                    "size": getattr(att, "Size", 0),
                    "index": att.Index
                })
                
        return {
            "id": msg.EntryID,
            "subject": msg.Subject,
            "from": msg.SenderName,
            "to": msg.To,
            "date": str(msg.ReceivedTime),
            "body_html": msg.HTMLBody,
            "body_text": msg.Body,
            "attachments_meta": attachments
        }
    except Exception as e:
        st.error(f"Failed to fetch item details: {e}")
        return None

def mark_as_read(entry_id):
    outlook = get_outlook()
    if not outlook: return
    mapi = outlook.GetNamespace("MAPI")
    try:
        msg = mapi.GetItemFromID(entry_id)
        msg.UnRead = False
        msg.Save()
    except:
        pass
        
def send_email(to_str, cc_str, bcc_str, subject, body, format_html, priority, attachments, is_draft=False):
    outlook = get_outlook()
    if not outlook: return False
    
    try:
        mail = outlook.CreateItem(0) # 0 = MailItem
        mail.To = to_str
        if cc_str: mail.CC = cc_str
        if bcc_str: mail.BCC = bcc_str
        mail.Subject = subject
        
        if format_html == "HTML":
            mail.BodyFormat = 2 # olFormatHTML
            mail.HTMLBody = body
        else:
            mail.BodyFormat = 1 # olFormatPlain
            mail.Body = body
            
        if priority == "High":
            mail.Importance = 2
        elif priority == "Low":
            mail.Importance = 0
        else:
            mail.Importance = 1
            
        # Due to security restrictions in browsers passing files directly into local COM objects,
        # we have to write the uploaded file temporarily to disk and attach via filepath.
        if attachments:
            temp_dir = os.path.join(os.getcwd(), "temp_attachments")
            os.makedirs(temp_dir, exist_ok=True)
            for file in attachments:
                temp_path = os.path.join(temp_dir, file.name)
                with open(temp_path, "wb") as f:
                    f.write(file.getvalue())
                mail.Attachments.Add(temp_path)
                
        if is_draft:
            mail.Save()
            return True
        else:
            mail.Send()
            return True
            
    except Exception as e:
        st.error(f"Failed to action email: {e}")
        return False


# -----------------------------------------------------------------------------
# [HELPER FUNCTIONS]
# -----------------------------------------------------------------------------
def format_date(date_str):
    if not date_str: return ""
    try:
        # Date string formatting cleanup for pywin32 localized dates
        return date_str[:16] # basic slicing to look clean if tz issues arise
    except:
        return date_str

def get_initials(name):
    if not name: return "?"
    name = re.sub(r'<.*?>', '', name).strip()
    parts = name.split()
    if len(parts) >= 2:
        return (parts[0][0] + parts[-1][0]).upper()
    return name[0:2].upper() if name else "?"

def safe_html(html_str):
    return re.sub(r'<script.*?>.*?</script>', '', str(html_str), flags=re.IGNORECASE | re.DOTALL)

# -----------------------------------------------------------------------------
# [UI BLOCKS]
# -----------------------------------------------------------------------------
def sidebar():
    with st.sidebar:
        st.title("📬 Windows Outlook Manager")
        st.markdown("---")
        
        if not win32com:
            st.error("COM Failed - Run on Windows")
            return
            
        st.success("✅ COM Connected to Outlook Desktop")
        
        outlook = get_outlook()
        try:
            current_acc = outlook.Session.Accounts.Item(1).SmtpAddress
        except Exception:
            current_acc = "Active Local Session"
        st.write(f"**Account:** {current_acc}")
        
        st.markdown("---")
        if st.button("🔄 Clear Viewer Cache", use_container_width=True):
            st.session_state.emails = []
            st.session_state.selected_email_id = None
            st.rerun()
            
        if st.session_state.last_sync:
            sync_diff = datetime.now() - st.session_state.last_sync
            mins = int(sync_diff.total_seconds() / 60)
            st.caption(f"Last synced: {'Just now' if mins == 0 else f'{mins} mins ago'}")
        
        with st.expander("⚙️ Settings"):
            st.slider("Max emails to load", 10, 100, 25, key="cfg_max_emails")
            
def tab_read_mails():
    if not win32com:
        st.info("👋 This script requires Windows with Outlook installed natively.")
        return

    emails = st.session_state.emails
    unread_count = sum(1 for e in emails if not e.get("is_read"))
    total_count = st.session_state.total_email_count

    m1, m2, m3 = st.columns(3)
    m1.markdown(f'<div class="metric-card"><div class="metric-title">Unread (Loaded)</div><div class="metric-value">{unread_count}</div></div>', unsafe_allow_html=True)
    m2.markdown(f'<div class="metric-card"><div class="metric-title">Loaded Size</div><div class="metric-value">{len(emails)}</div></div>', unsafe_allow_html=True)
    m3.markdown(f'<div class="metric-card"><div class="metric-title">Total in Folder</div><div class="metric-value">{total_count}</div></div>', unsafe_allow_html=True)
    st.write("")

    # Controls
    folder_list = list(FOLDER_MAP.keys())
    
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        selected_folder = st.selectbox("Folder", folder_list, index=folder_list.index(st.session_state.selected_folder) if st.session_state.selected_folder in folder_list else 3)
        st.session_state.selected_folder = selected_folder
    with c2:
        st.write("")
        st.write("")
        if st.button("Fetch / Refresh", use_container_width=True):
            with st.spinner("Talking to Outlook..."):
                fetch_emails(selected_folder, st.session_state.search_query, top=st.session_state.get("cfg_max_emails", 25))
                st.session_state.current_page = 0
                st.session_state.selected_email_id = None
                st.session_state.selected_email_detail = None
    with c3:
        search_query = st.text_input("Search (Subject/Sender)", key="search_query")

    st.markdown("---")

    if not emails and st.session_state.last_sync is None:
        fetch_emails(selected_folder, "", top=st.session_state.get("cfg_max_emails", 25))
        emails = st.session_state.emails

    if not emails:
        st.info("📭 No emails found in this folder based on your criteria.")
        return

    # List
    items_per_page = 10
    total_pages = math.ceil(len(emails) / items_per_page)
    if st.session_state.current_page >= total_pages: st.session_state.current_page = 0
    start_idx = st.session_state.current_page * items_per_page
    page_emails = emails[start_idx:start_idx+items_per_page]

    for e in page_emails:
        is_read = e.get("is_read", True)
        sender = e.get("from", "Unknown")
        
        row_class = "email-row" if is_read else "email-row unread-row"
        font_weight = "normal" if is_read else "bold"
        
        badges = []
        if not is_read: badges.append('<span class="badge badge-unread">Unread</span>')
        if e.get("is_flagged"): badges.append('<span class="badge badge-flagged">Flagged</span>')
        if e.get("has_attachments"): badges.append('<span class="badge badge-attach">📎 Attachment</span>')
        
        avatar = get_initials(sender)
        
        c1, c2, c3 = st.columns([0.5, 7, 2])
        with c1:
            st.markdown(f'<div class="avatar">{avatar}</div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f"""
                <div style="font-weight: {font_weight}; color: #333; margin-bottom: 2px;">{re.sub(r'<.*?>', '', sender).strip()} <span style="font-weight:normal; color:#666; font-size:14px;">- {e['subject']}</span></div>
                <div class="preview-text">{e['preview']}...</div>
                <div style="margin-top: 4px;">{" ".join(badges)}</div>
            """, unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div style="text-align:right; font-size:12px; color:#999; margin-bottom: 5px;">{format_date(e["date"])}</div>', unsafe_allow_html=True)
            if st.button("Read", key=f"read_{e['id']}", use_container_width=True):
                st.session_state.selected_email_id = e['id']
                with st.spinner("Fetching full item data via COM..."):
                    st.session_state.selected_email_detail = fetch_email_detail(e['id'])
                if not e.get("is_read"):
                    mark_as_read(e['id'])
                    e["is_read"] = True
        
        st.markdown(f'<div style="border-bottom: 1px solid #eee; margin-top: 10px;"></div>', unsafe_allow_html=True)

    # Pagination controls
    pc1, pc2, pc3 = st.columns([2, 6, 2])
    with pc1:
        if st.button("⬅️ Previous", disabled=(st.session_state.current_page == 0)):
            st.session_state.current_page -= 1
            st.rerun()
    with pc2:
        st.write(f"<div style='text-align:center;'>Page {st.session_state.current_page + 1} of {max(1, total_pages)}</div>", unsafe_allow_html=True)
    with pc3:
        if st.button("Next ➡️", disabled=(st.session_state.current_page >= total_pages - 1)):
            st.session_state.current_page += 1
            st.rerun()

    # Detail View
    if st.session_state.selected_email_id and st.session_state.selected_email_detail:
        detail = st.session_state.selected_email_detail
        
        st.markdown("---")
        with st.container(border=True):
            st.subheader(detail.get("subject", "(No Subject)"))
            st.markdown(f"**From:** {detail.get('from')}")
            st.markdown(f"**To:** {detail.get('to')}")
            st.markdown(f"**Date:** {format_date(detail.get('date'))}")
            
            bc1, bc2, bc3, bc4 = st.columns([1,1,1,5])
            with bc1:
                if st.button("⬅️ Close"):
                    st.session_state.selected_email_id = None
                    st.session_state.selected_email_detail = None
                    st.rerun()
            with bc2:
                if st.button("↩️ Reply"):
                    st.session_state.compose_to = detail.get('from', '')
                    st.session_state.compose_subject = f"Re: {detail.get('subject')}"
                    st.session_state.compose_body = f"\n\n---Original Message---\nFrom: {detail.get('from')}\n\n"
                    st.session_state.active_tab = "Send Mail"
                    st.rerun()
            
            st.markdown("---")
            
            if detail.get("body_html"):
                components.html(safe_html(detail["body_html"]), height=400, scrolling=True)
            else:
                st.text_area("Content", detail.get("body_text", ""), height=400, disabled=True)
            
            if detail.get("attachments_meta"):
                st.markdown("#### Attachments Meta Details")
                st.info("File download blocks stream directly via COM. Review files in Outlook Desktop Directly to securely save items.")
                for att in detail["attachments_meta"]:
                    st.write(f"- 📎 {att['filename']} ({att['size'] // 1024} KB)")

def tab_send_mail():
    if not win32com:
        st.warning("Please run on Windows.")
        return
        
    st.header("Compose Email")
    
    with st.form("compose_form", clear_on_submit=False):
        to_input = st.text_input("To (comma separated)", value=st.session_state.get("compose_to", ""))
        cc_input = st.text_input("CC (comma separated)", value=st.session_state.get("compose_cc", ""))
        subject_input = st.text_input("Subject", value=st.session_state.get("compose_subject", ""))
        
        body_format = st.radio("Format", ["Plain Text", "HTML"], index=0, horizontal=True)
        body_input = st.text_area("Body", value=st.session_state.get("compose_body", ""), height=250)
        
        priority = st.selectbox("Priority", ["Normal", "High", "Low"], index=0)
        
        attachments = st.file_uploader("Attachments", accept_multiple_files=True)
        
        b1, b2, b3 = st.columns([2, 2, 8])
        if b1.form_submit_button("📤 Send Email natively", type="primary"):
            if not to_input or not subject_input or not body_input:
                st.error("To, Subject, and Body are required.")
            else:
                with st.spinner("Dispatching via Outlook COM..."):
                    success = send_email(to_input, cc_input, "", subject_input, body_input, body_format, priority, attachments, is_draft=False)
                    if success:
                        st.success(f"✅ Email passed to desktop Outlook outbox successfully!")
                        st.session_state.compose_to = ""
                        st.session_state.compose_subject = ""
                        st.session_state.compose_body = ""
        
        if b2.form_submit_button("💾 Save as Draft natively"):
            if not subject_input:
                st.error("Subject is required.")
            else:
                with st.spinner("Filing away Draft via COM..."):
                    success = send_email(to_input, cc_input, "", subject_input, body_input, body_format, priority, attachments, is_draft=True)
                    if success:
                        st.success("Draft securely saved directly in your Desktop Outlook application!")

# -----------------------------------------------------------------------------
def main():
    sidebar()
    
    # Navigation mechanism using radio buttons mimicking tabs
    nav_selected = st.radio("Navigation", ["Read Mails", "Send Mail"], 
                            horizontal=True, 
                            label_visibility="collapsed",
                            index=0 if st.session_state.active_tab == "Read Mails" else 1)
    
    if nav_selected != st.session_state.active_tab:
        st.session_state.active_tab = nav_selected
        st.rerun()

    st.markdown("---")

    if st.session_state.active_tab == "Read Mails":
        tab_read_mails()
    else:
        tab_send_mail()

if __name__ == "__main__":
    main()
