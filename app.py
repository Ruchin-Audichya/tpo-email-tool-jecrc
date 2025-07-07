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

# ========== INIT ==========
def init_session():
    defaults = {
        'sender_email': "",
        'sender_password': "",
        'subject': "",
        'body': "",
        'html_mode': True,
        'footer_name': "Sijo Joji",
        'footer_image_url': "https://i.imgur.com/8Zz4pMc.png",
        'linkedin_url': "https://www.linkedin.com/in/sijojoji"
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

init_session()

# ========== PAGE SETUP ==========
st.set_page_config(page_title="TPO Email Sender - JECRC", layout="wide")
st.title("📧 TPO Email Automation Tool – JECRC University")

# ========== HOW TO USE GUIDE ==========
with st.expander("📘 How to Use This Tool (Click to Expand Guide)", expanded=False):
    st.markdown("""
This tool helps the TPO Cell send bulk emails with proper personalization, attachments, and official footer.

#### ✅ Step-by-Step Setup

1. **Enable 2-Step Verification on Gmail**  
   🔗 [Google 2-Step Verification](https://myaccount.google.com/security)

2. **Generate App Password (for Gmail)**  
   🔗 [App Passwords](https://myaccount.google.com/apppasswords)

3. **Create Google Service Account (.json file)**  
   🔗 [Google Cloud Console](https://console.cloud.google.com/)  
   - Enable Sheets API  
   - Create service account → Add Key → JSON  
   - Share your Google Sheet with:  
     `service-account-name@project-id.iam.gserviceaccount.com`

4. **Prepare Google Sheet Columns**  
   - Required: `email`  
   - Optional: `name`, `gender`  
   - Use `{name}`, `{title}`, `{footer}` in your message

5. **Run the App**
   - Upload your Gmail/app password  
   - Load the sheet + json file  
   - Use test mode to verify  
   - Click **Send Emails**

---

🎯 Placeholders: `{name}`, `{gender}`, `{title}`, `{footer}`  
📎 Attachment: One common file  
📝 HTML Mode: Enables styling, logo, LinkedIn button
""")

# ========== SIDEBAR ==========
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

# ========== COMPOSE ==========
st.subheader("📝 Compose Email")
col1, col2 = st.columns(2)
with col1:
    st.session_state.subject = st.text_input("Email Subject", value=st.session_state.subject)
with col2:
    st.session_state.html_mode = st.checkbox("Use HTML Formatting", value=st.session_state.html_mode)

st.session_state.body = st.text_area("Email Body (Use {name}, {title}, {footer})", height=250, value=st.session_state.body)

# ========== FOOTER ==========
st.subheader("✍️ Signature Block")
left, right = st.columns(2)
with left:
    st.session_state.footer_name = st.text_input("Name", st.session_state.footer_name)
    designation = st.text_input("Designation", "Head - Corporate Relations, Training & Placement at JECRC University")
    phone = st.text_input("Mobile", "+91-7229845674")
    landline = st.text_input("Phone", "01412770322")
    email = st.text_input("Email", "sijo.joji@jecrcu.edu.in")
    website = st.text_input("Website", "http://www.jecrcuniversity.edu.in/")
with right:
    address = st.text_area("Address", "Plot No IS -2036 to 2039, Ramchandrapura, Vidhani, Sitapura Extension, Jaipur - 303905 (Rajasthan), India")
    st.session_state.footer_image_url = st.text_input("Image URL (T&P photo)", st.session_state.footer_image_url)
    st.session_state.linkedin_url = st.text_input("LinkedIn URL", st.session_state.linkedin_url)

official_footer_html = f'''
<table style="font-family: Arial, sans-serif; font-size: 13px; line-height: 1.4;">
  <tr>
    <td style="vertical-align: top;">
      <img src="{st.session_state.footer_image_url}" alt="TPO Office" width="120" style="margin-right: 15px;">
    </td>
    <td>
      <p style="margin-bottom: 5px;"><strong style="color: #003366;">{st.session_state.footer_name}</strong><br>
      <span style="color: #336699;">{designation}</span></p>
      <hr style="border: 1px solid #336699; width: 100%; margin: 5px 0;">
      <p><strong>M</strong> {phone} &nbsp;&nbsp; <strong>P</strong> {landline} &nbsp;&nbsp;
      <strong>E</strong> <a href="mailto:{email}">{email}</a></p>
      <p><strong>W</strong> <a href="{website}">{website}</a></p>
      <p><strong>A</strong> {address}</p>
      <p><a href="{st.session_state.linkedin_url}" target="_blank" style="background-color: #0077b5; color: white; padding: 4px 8px; text-decoration: none; border-radius: 3px;">LinkedIn</a></p>
    </td>
  </tr>
</table>
'''
selected_footer = official_footer_html
st.markdown("**Live Footer Preview:**")
st.markdown(selected_footer, unsafe_allow_html=True)

# ========== LOAD SHEET ==========
st.subheader("📂 Load Recipient List from Google Sheet")
sheet_url = st.text_input("Google Sheet URL")
json_file = st.file_uploader("Upload Service Account JSON", type="json")

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

# ========== PREVIEW ==========
if data_loaded and st.button("👀 Preview Email"):
    sample = df.iloc[0]
    gender = str(sample.get('gender', '')).lower()
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

# ========== SEND ==========
if data_loaded and st.button("📨 Send Emails"):
    if not st.session_state.sender_email or not st.session_state.sender_password or not st.session_state.subject or not st.session_state.body:
        st.error("Missing required fields.")
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
                    gender = str(row.get('gender', '')).lower()
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
            if log_download and status_log:
                log_df = pd.DataFrame(status_log)
                st.download_button("📄 Download Send Log", log_df.to_csv(index=False), file_name="send_log.csv")

        except smtplib.SMTPAuthenticationError:
            st.error("Authentication failed. Check Gmail ID or App Password.")
        except Exception as e:
            st.error(f"Sending Error: {e}")
