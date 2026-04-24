"""
Mail Manager — Streamlit App
=====================================
Single-file Streamlit application to read and send emails
via IMAP and SMTP using standard Python libraries.

Supports Outlook and Gmail using App Passwords.

Requirements:
    pip install streamlit pytz
"""

import imaplib
import smtplib
import email
from email.header import decode_header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import time
import base64
import math
import re
import streamlit as st
import pytz
from datetime import datetime, timedelta
import streamlit.components.v1 as components

# -----------------------------------------------------------------------------
# [TOP BLOCK] Initialization & Config
# -----------------------------------------------------------------------------

st.set_page_config(page_title="Mail Manager", page_icon="📬", layout="wide")

# Custom CSS
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
        "logged_in": False,
        "email_address": "",
        "app_password": "",
        "provider": "Outlook",
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
        "folder_list": [],
        "selected_folder": "INBOX",
        "total_email_count": 0,
        "last_sync": None
    }
    for key, value in session_defaults.items():
        st.session_state.setdefault(key, value)

initialize_session_state()

# -----------------------------------------------------------------------------
# [LOGIN FLOW]
# -----------------------------------------------------------------------------
def get_servers(provider):
    if provider == "Gmail":
        return "imap.gmail.com", "smtp.gmail.com"
    else:
        return "outlook.office365.com", "smtp.office365.com"

@st.dialog("🔒 Login to Mail Manager")
def login_dialog():
    st.write("Bypass standard OAuth by generating a 16-digit **App Password** from your security settings!")
    
    provider = st.selectbox("Select Provider", ["Outlook", "Gmail"])
    email_inp = st.text_input("Email Address")
    pwd_inp = st.text_input("App Password", type="password")
    
    if st.button("Connect Account", type="primary"):
        with st.spinner("Authenticating..."):
            imap_host, _ = get_servers(provider)
            try:
                mail = imaplib.IMAP4_SSL(imap_host)
                mail.login(email_inp, pwd_inp)
                mail.logout()
                st.session_state.email_address = email_inp
                st.session_state.app_password = pwd_inp
                st.session_state.provider = provider
                st.session_state.logged_in = True
                st.success("✅ Logged in successfully!")
                time.sleep(1)
                st.rerun()
            except imaplib.IMAP4.error as e:
                st.error("Login failed. Check your App Password, and ensure IMAP is enabled on your account.")

def get_imap_conn():
    if not st.session_state.logged_in: return None
    imap_host, _ = get_servers(st.session_state.provider)
    try:
        mail = imaplib.IMAP4_SSL(imap_host)
        mail.login(st.session_state.email_address, st.session_state.app_password)
        return mail
    except:
        st.session_state.logged_in = False
        return None

def get_smtp_conn():
    if not st.session_state.logged_in: return None
    _, smtp_host = get_servers(st.session_state.provider)
    try:
        server = smtplib.SMTP(smtp_host, 587)
        server.starttls()
        server.login(st.session_state.email_address, st.session_state.app_password)
        return server
    except:
        return None

# -----------------------------------------------------------------------------
# [HELPERS & PARSING]
# -----------------------------------------------------------------------------
def decode_mime_str(s):
    if not s: return ""
    decoded_string = []
    try:
        for part, enc in decode_header(s):
            if isinstance(part, bytes):
                decoded_string.append(part.decode(enc or 'utf-8', errors='replace'))
            else:
                decoded_string.append(str(part))
        return "".join(decoded_string)
    except:
        return str(s)

def extract_body(msg):
    body_text = ""
    body_html = ""
    attachments = []
    
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            cdispo = str(part.get('Content-Disposition'))
            
            if ctype == 'text/plain' and 'attachment' not in cdispo:
                try: body_text += part.get_payload(decode=True).decode()
                except: pass
            elif ctype == 'text/html' and 'attachment' not in cdispo:
                try: body_html += part.get_payload(decode=True).decode()
                except: pass
            elif 'attachment' in cdispo or part.get_filename():
                fname = decode_mime_str(part.get_filename())
                if fname:
                    attachments.append({
                        "filename": fname,
                        "data": part.get_payload(decode=True),
                        "content_type": ctype
                    })
    else:
        ctype = msg.get_content_type()
        try:
            payload = msg.get_payload(decode=True).decode()
            if ctype == 'text/plain': body_text = payload
            elif ctype == 'text/html': body_html = payload
        except:
            pass
            
    return body_html if body_html else body_text, "HTML" if body_html else "Text", attachments

def format_date(date_str):
    if not date_str: return ""
    try:
        from email.utils import parsedate_to_datetime
        dt = parsedate_to_datetime(date_str)
        now = datetime.now(dt.tzinfo)
        diff = now - dt
        
        if diff < timedelta(minutes=5): return "Just now"
        elif diff < timedelta(hours=24) and now.date() == dt.date(): return dt.strftime("%I:%M %p")
        elif diff < timedelta(days=2) and now.date() - dt.date() == timedelta(days=1): return "Yesterday"
        elif diff < timedelta(days=7): return dt.strftime("%a")
        else: return dt.strftime("%d %b")
    except:
        return date_str

def get_initials(name):
    if not name: return "?"
    # Strip email if name comes as "Name <email>"
    name = re.sub(r'<.*?>', '', name).strip()
    parts = name.split()
    if len(parts) >= 2:
        return (parts[0][0] + parts[-1][0]).upper()
    return name[0:2].upper() if name else "?"

def safe_html(html_str):
    # Basic script tag removal
    return re.sub(r'<script.*?>.*?</script>', '', html_str, flags=re.IGNORECASE | re.DOTALL)

# -----------------------------------------------------------------------------
# [IMAP LOGIC]
# -----------------------------------------------------------------------------
@st.cache_data(ttl=300, show_spinner=False)
def get_folders_cached(auth_trigger): 
    # auth_trigger is passed to bust cache dynamically
    mail = get_imap_conn()
    if not mail: return ["INBOX"]
    status, folders = mail.list()
    mail.logout()
    folder_names = []
    for folder in folders:
        # Decode IMAP folder list string
        parts = folder.decode().split(' "/" ')
        if len(parts) == 2:
            fname = parts[1].strip('"')
            if fname not in ["[Gmail]", "INBOX"]: # Clean up some defaults
                folder_names.append(fname)
    return ["INBOX"] + sorted(folder_names)

def fetch_emails(folder="INBOX", search_query="", top=25, skip=0):
    mail = get_imap_conn()
    if not mail: return
    
    # Try selecting folder
    status, data = mail.select(f'"{folder}"', readonly=True)
    if status != "OK":
        st.error(f"Could not open folder {folder}")
        mail.logout()
        return

    # Handle Search
    if search_query:
        # Simple subject/from search
        status, messages = mail.search(None, f'(OR SUBJECT "{search_query}" FROM "{search_query}")')
    else:
        status, messages = mail.search(None, 'ALL')

    mail_ids = messages[0].split()
    st.session_state.total_email_count = len(mail_ids)
    
    if not mail_ids:
        st.session_state.emails = []
        mail.logout()
        return
        
    mail_ids.reverse() # Newest first
    target_ids = mail_ids[skip:skip+top]
    
    results = []
    
    # Fetch in bulk if possible, or iterate
    # Using iteration since target_ids is small (top=25 max usually)
    for m_id in target_ids:
        # Fetch FLAGS and BODY.PEEK to keep them unread until explicitly opened
        res, msg_data = mail.fetch(m_id, '(FLAGS BODY.PEEK[])')
        
        flags = []
        raw_email = b""
        
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                # response_part[0] contains the flags and headers
                # response_part[1] contains the actual body bytes
                parse_head = response_part[0].decode(errors='ignore')
                raw_email = response_part[1]
                
                # Extract flags
                flag_match = re.search(r'FLAGS \((.*?)\)', parse_head)
                if flag_match:
                    flags = flag_match.group(1).split()
        
        if raw_email:
            msg = email.message_from_bytes(raw_email)
            subject = decode_mime_str(msg.get("Subject"))
            sender = decode_mime_str(msg.get("From"))
            date_str = msg.get("Date")
            
            body_content, content_type, attachments = extract_body(msg)
            
            clean_text_preview = re.sub('<[^<]+?>', '', body_content) if content_type == "HTML" else body_content
            clean_text_preview = clean_text_preview.replace('\r', '').replace('\n', ' ')
            
            results.append({
                "id": m_id.decode(),
                "subject": subject,
                "from": sender,
                "to": decode_mime_str(msg.get("To")),
                "date": date_str,
                "is_read": '\\Seen' in flags,
                "is_flagged": '\\Flagged' in flags,
                "has_attachments": len(attachments) > 0,
                "preview": clean_text_preview[:100],
                "body": body_content,
                "body_type": content_type,
                "attachments": attachments
            })

    mail.logout()
    st.session_state.emails = results
    st.session_state.last_sync = datetime.now()

def mark_as_read(uid):
    mail = get_imap_conn()
    if not mail: return
    mail.select(f'"{st.session_state.selected_folder}"')
    mail.store(uid, '+FLAGS', '\\Seen')
    mail.logout()

# -----------------------------------------------------------------------------
# [UI BLOCKS]
# -----------------------------------------------------------------------------
def sidebar():
    with st.sidebar:
        st.title("📬 Mail Manager")
        st.markdown("---")
        
        if not st.session_state.logged_in:
            st.warning("Not Connected")
            if st.button("Log In (App Password)", use_container_width=True):
                login_dialog()
            return
            
        st.success("✅ Connected")
        st.write(f"**Account:** {st.session_state.email_address}")
        st.write(f"**Provider:** {st.session_state.provider}")
        
        st.markdown("---")
        if st.button("🗑️ Log Out", use_container_width=True):
            for k in list(st.session_state.keys()): del st.session_state[k]
            st.rerun()
            
        if st.button("🔄 Clear App Cache", use_container_width=True):
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
    if not st.session_state.logged_in:
        st.info("👋 Welcome! Use the sidebar to sign in using your App Password.")
        return

    emails = st.session_state.emails
    unread_count = sum(1 for e in emails if not e.get("is_read"))
    total_count = st.session_state.total_email_count

    m1, m2, m3 = st.columns(3)
    m1.markdown(f'<div class="metric-card"><div class="metric-title">Unread (Loaded)</div><div class="metric-value">{unread_count}</div></div>', unsafe_allow_html=True)
    m2.markdown(f'<div class="metric-card"><div class="metric-title">Pagination Size</div><div class="metric-value">{len(emails)}</div></div>', unsafe_allow_html=True)
    m3.markdown(f'<div class="metric-card"><div class="metric-title">Total in Folder</div><div class="metric-value">{total_count}</div></div>', unsafe_allow_html=True)
    st.write("")

    # Controls
    folder_list = get_folders_cached(st.session_state.email_address)
    
    c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
    with c1:
        selected_folder = st.selectbox("Folder", folder_list, index=folder_list.index(st.session_state.selected_folder) if st.session_state.selected_folder in folder_list else 0)
        st.session_state.selected_folder = selected_folder
    with c4:
        st.write("")
        st.write("")
        if st.button("Fetch / Refresh", use_container_width=True):
            with st.spinner("Fetching emails..."):
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
                st.session_state.selected_email_detail = e
                if not e.get("is_read"):
                    mark_as_read(e['id'])
                    e["is_read"] = True
        
        st.markdown(f'<div style="border-bottom: 1px solid #eee; margin-top: 10px;"></div>', unsafe_allow_html=True)

    # Detail View
    if st.session_state.selected_email_id and st.session_state.selected_email_detail:
        detail = st.session_state.selected_email_detail
        
        st.markdown("---")
        with st.container(border=True):
            st.subheader(detail.get("subject", "(No Subject)"))
            st.markdown(f"**From:** {detail.get('from')}")
            st.markdown(f"**Date:** {format_date(detail.get('date'))}")
            
            bc1, bc2, bc3, bc4 = st.columns([1,1,1,5])
            with bc1:
                if st.button("⬅️ Close"):
                    st.session_state.selected_email_id = None
                    st.session_state.selected_email_detail = None
                    st.rerun()
            with bc2:
                if st.button("↩️ Reply"):
                    # Basic extraction of email address for reply
                    from_email = re.search(r'<([^>]+)>', detail.get('from', ''))
                    st.session_state.compose_to = from_email.group(1) if from_email else detail.get('from')
                    st.session_state.compose_subject = f"Re: {detail.get('subject')}"
                    st.session_state.compose_body = f"\n\n---Original Message---\nFrom: {detail.get('from')}\n\n"
                    st.session_state.active_tab = "Send Mail"
                    st.rerun()
            
            st.markdown("---")
            if detail.get("body_type") == "HTML":
                components.html(safe_html(detail["body"]), height=400, scrolling=True)
            else:
                st.text_area("Content", detail["body"], height=400, disabled=True)
            
            if detail.get("has_attachments"):
                st.markdown("#### Attachments")
                for att in detail["attachments"]:
                    st.download_button(
                        label=f"⬇️ {att['filename']} ({len(att['data']) // 1024} KB)",
                        data=att['data'],
                        file_name=att['filename'],
                        mime=att['content_type']
                    )

def tab_send_mail():
    if not st.session_state.logged_in:
        st.warning("Please login first.")
        return
        
    st.header("Compose Email")
    
    with st.form("compose_form", clear_on_submit=False):
        to_input = st.text_input("To (comma separated)", value=st.session_state.get("compose_to", ""))
        cc_input = st.text_input("CC (comma separated)", value=st.session_state.get("compose_cc", ""))
        subject_input = st.text_input("Subject", value=st.session_state.get("compose_subject", ""))
        
        body_format = st.radio("Format", ["Plain Text", "HTML"], index=0, horizontal=True)
        body_input = st.text_area("Body", value=st.session_state.get("compose_body", ""), height=250)
        
        attachments = st.file_uploader("Attachments", accept_multiple_files=True)
        
        b1, b2 = st.columns([2, 8])
        if b1.form_submit_button("📤 Send Email", type="primary"):
            if not to_input or not subject_input or not body_input:
                st.error("To, Subject, and Body are required.")
            else:
                with st.spinner("Sending via SMTP..."):
                    server = get_smtp_conn()
                    if server:
                        msg = MIMEMultipart()
                        msg['From'] = st.session_state.email_address
                        msg['To'] = to_input
                        msg['Cc'] = cc_input
                        msg['Subject'] = subject_input
                        
                        msg.attach(MIMEText(body_input, 'html' if body_format == "HTML" else 'plain'))
                        
                        for file in attachments:
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(file.getvalue())
                            encoders.encode_base64(part)
                            part.add_header('Content-Disposition', f'attachment; filename="{file.name}"')
                            msg.attach(part)
                        
                        all_recipients = [e.strip() for e in to_input.split(",")] + [e.strip() for e in cc_input.split(",") if e.strip()]
                        
                        try:
                            server.sendmail(st.session_state.email_address, all_recipients, msg.as_string())
                            server.quit()
                            st.success(f"✅ Email sent successfully to {to_input}!")
                            st.session_state.compose_to = ""
                            st.session_state.compose_subject = ""
                            st.session_state.compose_body = ""
                        except Exception as e:
                            st.error(f"Failed to send: {e}")
                    else:
                        st.error("SMTP Connection failed.")

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
