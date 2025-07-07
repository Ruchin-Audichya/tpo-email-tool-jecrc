# === TPO EMAIL TOOL WITH OFFICIAL FOOTER (MATCHES SIJO MAM STYLE) ===

import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
import os
import tempfile
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import re

# ========== Session State Setup ==========
def init_session():
    defaults = {
        'sender_email': "",
        'sender_password': "",
        'subject': "",
        'body': "",
        'html_mode': True,
        'footer_name': "Sijo Joji",
        'footer_image_url': "https://i.imgur.com/8Zz4pMc.png",  # Example JECRC T&P image
        'linkedin_url': "https://www.linkedin.com/in/sijojoji"
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

init_session()

# ========== Streamlit Setup ==========
st.set_page_config(page_title="TPO Email Sender - JECRC", layout="wide")
st.title("📧 TPO Email Automation Tool – JECRC University")

# ========== Sidebar ==========
with st.sidebar:
    st.header("🔐 Sender Authentication")
    st.session_state.sender_email = st.text_input("Gmail Address", value=st.session_state.sender_email)
    st.session_state.sender_password = st.text_input("Gmail App Password", type="password", value=st.session_state.sender_password)

    st.header("🧪 Send Mode")
    test_mode = st.checkbox("Test Mode (Send to Admin Only)", value=False)
    admin_email = st.text_input("Admin/Test Email", value="yourtest@example.com")

    st.header("📎 Attachment")
    common_attachment = st.file_uploader("Upload common attachment", type=None)

    log_download = st.checkbox("Enable Email Log Download")

# ========== Compose Email ==========
st.subheader("📝 Compose Email")
col1, col2 = st.columns(2)
with col1:
    st.session_state.subject = st.text_input("Email Subject", value=st.session_state.subject)
with col2:
    st.session_state.html_mode = st.checkbox("Use HTML Formatting", value=st.session_state.html_mode)

st.session_state.body = st.text_area("Email Body (Use {name}, {title}, {footer})", height=250, value=st.session_state.body)

# ========== Official JECRC Footer ==========
st.subheader("✍️ Signature Block")
use_custom_footer = True  # Always using JECRC official footer

# Editable values for dynamic footer
left, right = st.columns(2)
with left:
    st.session_state.footer_name = st.text_input("Name", st.session_state.footer_name)
    footer_designation = st.text_input("Designation", "Head - Corporate Relations, Training & Placement at JECRC University")
    footer_phone = st.text_input("Mobile", "+91-7229845674")
    footer_landline = st.text_input("Phone", "01412770322")
    footer_email = st.text_input("Email", "sijo.joji@jecrcu.edu.in")
    footer_website = st.text_input("Website", "http://www.jecrcuniversity.edu.in/")
with right:
    footer_address = st.text_area("Address", "Plot No IS -2036 to 2039, Ramchandrapura, Vidhani, Sitapura Extension, Jaipur - 303905 (Rajasthan), India")
    st.session_state.footer_image_url = st.text_input("Image URL (T&P photo)", st.session_state.footer_image_url)
    st.session_state.linkedin_url = st.text_input("LinkedIn URL", st.session_state.linkedin_url)

# Final footer HTML
official_footer_html = f'''
<table style="font-family: Arial, sans-serif; font-size: 13px; line-height: 1.4;">
  <tr>
    <td style="vertical-align: top;">
      <img src="{st.session_state.footer_image_url}" alt="TPO Office" width="120" style="margin-right: 15px;">
    </td>
    <td>
      <p style="margin-bottom: 5px;"><strong style="color: #003366;">{st.session_state.footer_name}</strong><br>
      <span style="color: #336699;">{footer_designation}</span></p>
      <hr style="border: 1px solid #336699; width: 100%; margin: 5px 0;">
      <p><strong>M</strong> {footer_phone} &nbsp;&nbsp; <strong>P</strong> {footer_landline} &nbsp;&nbsp;
      <strong>E</strong> <a href="mailto:{footer_email}">{footer_email}</a></p>
      <p><strong>W</strong> <a href="{footer_website}">{footer_website}</a></p>
      <p><strong>A</strong> {footer_address}</p>
      <p><a href="{st.session_state.linkedin_url}" target="_blank" style="background-color: #0077b5; color: white; padding: 4px 8px; text-decoration: none; border-radius: 3px;">LinkedIn</a></p>
    </td>
  </tr>
</table>
'''
selected_footer = official_footer_html
st.markdown("**Live Footer Preview:**")
st.markdown(official_footer_html, unsafe_allow_html=True)

# ========== Google Sheet Upload ==========
st.subheader("📂 Load Recipient List from Google Sheet")
sheet_url = st.text_input("Google Sheet URL")
json_file = st.file_uploader("Service Account JSON", type="json")

def validate_columns(df):
    return [col for col in ['email'] if col not in df.columns]

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

        missing = validate_columns(df)
        if missing:
            st.error(f"Missing columns: {', '.join(missing)}")
        else:
            st.success(f"Loaded {len(df)} records from '{sheet_name}'")
            st.dataframe(df.head())
            data_loaded = True

    except Exception as e:
        st.error(f"Sheet Load Error: {e}")

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

    st.markdown(f"**To:** {sample['email']}  |  **Subject:** {st.session_state.subject}")
    if st.session_state.html_mode:
        st.markdown(preview_body, unsafe_allow_html=True)
    else:
        st.text(preview_body)

# ========== Send Emails ==========
if data_loaded and st.button("📨 Send Emails"):
    if not st.session_state.sender_email or not st.session_state.sender_password or not st.session_state.subject or not st.session_state.body:
        st.error("Missing required sender or email fields.")
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

                    if common_attachment:
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

            st.success("✅ All emails sent successfully!")
            if status_log:
                log_df = pd.DataFrame(status_log)
                date_str = datetime.now().strftime("%Y-%m-%d_%H-%M")
                os.makedirs("email_logs", exist_ok=True)
                log_df.to_csv(f"email_logs/send_log_{date_str}.csv", index=False)
                if log_download:
                    st.download_button("📄 Download Send Log", log_df.to_csv(index=False), file_name="send_log.csv")

        except smtplib.SMTPAuthenticationError:
            st.error("Authentication failed. Please check email and app password.")
        except Exception as e:
            st.error(f"Sending Error: {e}")