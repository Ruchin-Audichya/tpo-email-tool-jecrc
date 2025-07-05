import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
import os
import tempfile
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re
from datetime import datetime

# ========== Session State Setup ==========
if 'sender_email' not in st.session_state:
    st.session_state.sender_email = ""
if 'sender_password' not in st.session_state:
    st.session_state.sender_password = ""
if 'subject' not in st.session_state:
    st.session_state.subject = ""
if 'body' not in st.session_state:
    st.session_state.body = ""
if 'html_mode' not in st.session_state:
    st.session_state.html_mode = True
if 'footer_name' not in st.session_state:
    st.session_state.footer_name = "Sijo Joji"

# ========== Streamlit Setup ==========
st.set_page_config(page_title="TPO Email Sender - JECRC", layout="wide")
st.title("📧 TPO Email Automation Tool – JECRC University")

with st.expander("🛠️ How to Use This Tool (Start Here)", expanded=False):
    st.markdown("""
    ### Step-by-Step Guide

    1. **Login with Gmail App Password** in the sidebar.
       - Use 2-Step Verification in Gmail.
       - Generate an App Password for this tool.

    2. **Upload Google Sheet**:
       - Must contain at least an `email` column.
       - Optional placeholders like `{name}`, `{gender}`, `{company}` will be replaced automatically.
       - Example row:
         | name | gender | email              |
         |------|--------|--------------------|
         | Ravi | male   | ravi@example.com   |

    3. **Compose Email**:
       - Use placeholders like:
         - `{name}` → Recipient's name (from sheet)
         - `{title}` → Mr./Ms. (based on gender)
         - `{footer}` → Automatically replaced with the selected footer block

    4. **Footer Selection**:
       - You can choose a default TPO footer (e.g. "Sijo Mam")
       - Or create your **own custom footer** with name, role, phone, LinkedIn, etc.
       - This footer will be appended to every email.

    5. **Preview Email**:
       - Use the 👀 Preview Email button to see what the final email will look like for one recipient.

    6. **Send Emails**:
       - Click "📨 Send Emails" when you're ready.
       - In **Test Mode**, emails go to the Admin only.
       - In real mode, emails are sent to all recipients.

    7. **Download Log**:
       - Enable "Save Send Log" to download a CSV after sending.
       - A full log is also saved internally by date for audit.

    ---
    ⚠️ Gmail Limit: Avoid sending to more than 400 people per day with one Gmail account to prevent blocking.

    🧠 Tip: Always preview and test before final send!
    """)

# ========== Sidebar ==========
with st.sidebar:
    st.header("🔐 Sender Authentication")
    st.session_state.sender_email = st.text_input("Gmail Address", value=st.session_state.sender_email, placeholder="e.g. yourname@gmail.com")
    st.session_state.sender_password = st.text_input("Gmail App Password", type="password", value=st.session_state.sender_password, placeholder="Paste your Gmail App Password")

    with st.expander("❓ How to get Gmail App Password"):
        st.markdown("""
        1. Visit [Google Account Security](https://myaccount.google.com/security)
        2. Turn on **2-Step Verification**
        3. Scroll to **App Passwords** → Select app: Mail → Device: Other → Generate
        4. Paste the generated 16-char password here
        """)

    st.header("🧪 Send Mode")
    test_mode = st.checkbox("Test Mode (Send to Admin Only)", value=False)
    if test_mode:
        st.warning("🧪 Test Mode is ON. Emails will be sent to admin email only.")
    admin_email = st.text_input("Admin/Test Email", value="yourtest@example.com", placeholder="Where test email goes")

    st.header("📌 Attachment Options")
    common_attachment = st.file_uploader("Common Attachment (optional)", type=None)

    st.header("📥 Save Send Log")
    log_download = st.checkbox("Enable Email Status Log Download")

# ========== Compose Email ==========
st.subheader("📝 Compose Email")
col1, col2 = st.columns(2)
with col1:
    st.session_state.subject = st.text_input("Email Subject", value=st.session_state.subject, placeholder="e.g. Internship Offer for {name}")
with col2:
    st.session_state.html_mode = st.checkbox("Use HTML Formatting", value=st.session_state.html_mode)

st.session_state.body = st.text_area("Email Body (Use {name}, {title}, {footer})", height=250, value=st.session_state.body, placeholder="e.g. Dear {title} {name},\n\nWe are glad to offer you...")

# ========== Footer Selection ==========
st.subheader("✍️ Signature Block")
use_custom_footer = st.checkbox("Use Custom Footer")

if use_custom_footer:
    st.markdown("**🖊️ Custom Footer Editor**")
    left, right = st.columns(2)
    with left:
        st.session_state.footer_name = st.text_input("Name", st.session_state.footer_name)
        footer_designation = st.text_input("Designation", "Head - Corporate Relations, Training & Placement")
        footer_phone = st.text_input("Phone", "+91-7229845674")
        footer_email = st.text_input("Email", "sijo.joji@jecrc.edu.in")
        footer_website = st.text_input("Website", "www.jecrcuniversity.edu.in")
        footer_address = st.text_area("Address", "Plot No IS –2036 to 2039, Sitapura Extension, Jaipur")
        linkedin_url = st.text_input("LinkedIn URL", "https://www.linkedin.com")
    with right:
        custom_footer_html = f'''
        <div style="font-family: Arial, sans-serif; font-size: 13px; line-height: 1.4;">
            <p><strong>{st.session_state.footer_name}</strong><br>{footer_designation}</p>
            <p><b>M:</b> {footer_phone} | <b>E:</b> <a href="mailto:{footer_email}">{footer_email}</a><br>
            <b>W:</b> <a href="http://{footer_website}">{footer_website}</a></p>
            <p><b>A:</b> {footer_address}</p>
            <p style="margin-top:8px;"><a href="{linkedin_url}" target="_blank">🔗 LinkedIn</a></p>
            <p style="font-size: 11px; color: gray;"><strong>Note:</strong> This email may contain confidential information.</p>
        </div>
        '''
        st.markdown("**Live Preview:**")
        st.markdown(custom_footer_html, unsafe_allow_html=True)
    selected_footer = custom_footer_html
else:
    footer_option = st.selectbox("Select Footer", ["Sijo Mam"])
    footers = {
        "Sijo Mam": '''<div style="font-family: Arial, sans-serif; font-size: 13px;">
          <p><strong>Sijo Joji</strong><br>Head - Corporate Relations, Training & Placement at JECRC University</p>
          <p><b>M</b> +91-7229845674 &nbsp; <b>E</b> <a href="mailto:sijo.joji@jecrc.edu.in">sijo.joji@jecrc.edu.in</a><br>
          <b>W</b> <a href="http://www.jecrcuiversity.edu.in/">www.jecrcuiversity.edu.in</a><br>
          <b>A</b> Sitapura Extension, Jaipur</p>
          <p><a href="https://www.linkedin.com" target="_blank">🔗 LinkedIn</a></p>
          <p style="font-size: 11px; color: gray;">Important: Confidential and intended for the recipient only.</p>
        </div>'''
    }
    selected_footer = footers[footer_option]

# ========== Google Sheet ==========
st.subheader("📂 Load Recipient List from Google Sheet")
sheet_url = st.text_input("Paste Google Sheet URL")
json_file = st.file_uploader("Upload Google Service Account JSON", type="json")

def validate_columns(df):
    required = ['email']
    return [col for col in required if col not in df.columns]

data_loaded = False
if sheet_url and json_file:
    try:
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/drive']
        with tempfile.NamedTemporaryFile(delete=False) as tmp_json:
            tmp_json.write(json_file.read())
            json_path = tmp_json.name
        creds = ServiceAccountCredentials.from_json_keyfile_name(json_path, scope)
        client = gspread.authorize(creds)

        sheet_id = sheet_url.split("/d/")[1].split("/")[0]
        gfile = client.open_by_key(sheet_id)
        sheet_names = [ws.title for ws in gfile.worksheets()]
        sheet_name = st.selectbox("Select Sheet Tab", sheet_names)
        sheet = gfile.worksheet(sheet_name)
        data = sheet.get_all_records()
        df = pd.DataFrame(data)

        missing_cols = validate_columns(df)
        if missing_cols:
            st.error(f"Missing required columns: {', '.join(missing_cols)}")
        else:
            st.success(f"Loaded {len(df)} records from '{sheet_name}'")
            st.markdown(f"**Available Placeholders:** {', '.join(df.columns)}")
            st.dataframe(df.head())
            data_loaded = True

    except Exception as e:
        st.error(f"Error loading sheet: {e}")

# ========== Preview Email ==========
if data_loaded and st.button("👀 Preview Email"):
    sample = df.iloc[0]
    name = sample.get('name', '')
    gender = sample.get('gender', '').lower()
    title = "Mr." if gender == "male" else "Ms." if gender == "female" else ""
    preview_body = st.session_state.body.replace("{title}", title).replace("{footer}", selected_footer)
    for key, value in sample.items():
        if pd.notna(value):
            preview_body = preview_body.replace(f"{{{key}}}", str(value))
    st.markdown("### Sample Email Preview")
    st.markdown(f"**To:** {sample['email']}  |  **Subject:** {st.session_state.subject}")
    if st.session_state.html_mode:
        st.markdown(preview_body, unsafe_allow_html=True)
    else:
        st.text(preview_body)

# ========== Send Emails ==========
if data_loaded and st.button("📨 Send Emails"):
    if not st.session_state.sender_email or not st.session_state.sender_password or not st.session_state.subject or not st.session_state.body:
        st.error("Please fill all required fields.")
    else:
        status_log = []
        try:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(st.session_state.sender_email, st.session_state.sender_password)
                progress = st.progress(0)
                status = st.empty()
                target_df = df.head(1) if test_mode else df

                for i, row in target_df.iterrows():
                    recipient = admin_email if test_mode else row['email']
                    name = str(row.get('name', '')).strip()
                    gender = str(row.get('gender', '')).strip().lower()
                    title = "Mr." if gender == "male" else "Ms." if gender == "female" else ""

                    final_body = st.session_state.body.replace("{title}", title).replace("{footer}", selected_footer)
                    for key, value in row.items():
                        if pd.notna(value):
                            final_body = final_body.replace(f"{{{key}}}", str(value))

                    msg = EmailMessage()
                    msg['From'] = st.session_state.sender_email
                    msg['To'] = recipient
                    msg['Subject'] = st.session_state.subject

                    if st.session_state.html_mode:
                        msg.set_content(final_body, subtype='html')
                    else:
                        msg.set_content(final_body)

                    attach_ok = False
                    if 'attachment_path' in row and pd.notna(row['attachment_path']):
                        try:
                            with open(row['attachment_path'], 'rb') as f:
                                msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename=os.path.basename(row['attachment_path']))
                                attach_ok = True
                        except Exception as e:
                            status.warning(f"Attachment error for {recipient}: {e}")

                    if not attach_ok and common_attachment:
                        with tempfile.NamedTemporaryFile(delete=False) as tmp:
                            tmp.write(common_attachment.read())
                            temp_path = tmp.name
                        with open(temp_path, 'rb') as f:
                            msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename=common_attachment.name)
                        os.unlink(temp_path)

                    smtp.send_message(msg)
                    progress.progress((i + 1) / len(target_df))
                    status.text(f"Sent to: {recipient} ({i+1}/{len(target_df)})")
                    status_log.append({"email": recipient, "status": "Sent"})

            st.success("🎉 All emails sent successfully!")

            # Save log
            if status_log:
                log_df = pd.DataFrame(status_log)
                date_str = datetime.now().strftime("%Y-%m-%d_%H-%M")
                os.makedirs("email_logs", exist_ok=True)
                log_df.to_csv(f"email_logs/send_log_{date_str}.csv", index=False)

                if log_download:
                    st.download_button("📄 Download Send Log", log_df.to_csv(index=False), file_name="send_log.csv")

        except smtplib.SMTPAuthenticationError:
            st.error("Login failed. Check email or app password.")
        except Exception as e:
            st.error(f"Error: {e}")

# ========== Log Viewer ==========
st.subheader("📅 Sent Email Log History")
log_dir = "email_logs"
os.makedirs(log_dir, exist_ok=True)
all_logs = sorted([f for f in os.listdir(log_dir) if f.endswith(".csv")], reverse=True)

if all_logs:
    selected_log = st.selectbox("Select Log Date", all_logs)
    if selected_log:
        df_log = pd.read_csv(os.path.join(log_dir, selected_log))
        st.dataframe(df_log)
else:
    st.info("No previous logs found.")
