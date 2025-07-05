# 📧 JECRC TPO Email Automation Tool

A lightweight web-based email automation tool built with Streamlit for JECRC University’s TPO cell.

## 🔹 Features
- Google Sheet integration (load students’ data)
- Custom & dynamic HTML email generation
- Secure Gmail-based email sending (App Password only)
- Attachment support
- Custom footer builder
- Logs and downloadable history
- Test mode to preview emails safely

## 🔒 Gmail Setup
1. Enable **2-Step Verification**: https://myaccount.google.com/security
2. Generate an **App Password** from “App Passwords” → Mail → Other

## 📋 Sheet Format
Your Google Sheet should have at least an `email` column.
Optional placeholders:
- `{name}`, `{gender}`, `{company}`, `{title}`


## 📂 Logs
All email logs are saved under `email_logs/` with timestamped CSVs.

---

## 🧠 Created by: Ruchin Audichya
For internal use only – JECRC University
# tpo-email-tool-jecrc
