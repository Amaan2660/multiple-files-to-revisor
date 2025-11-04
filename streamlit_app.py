import io
import os
import zipfile
import smtplib
import time
from email.utils import formataddr
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

import streamlit as st

# =========================================
# APP CONFIG
# =========================================
st.set_page_config(page_title="Multiple files to revisor", page_icon="📧", layout="centered")
st.title("📧 Multiple files to revisor")
st.caption("Upload PDFs (or a ZIP of PDFs), choose a recipient, and send one email per PDF — each in its own thread.")

# =========================================
# SIDEBAR: SMTP / SENDER SETTINGS
# =========================================
with st.sidebar:
    st.header("SMTP / Sender Settings")

    # Prefill from Streamlit Secrets if available
    default_sender = st.secrets.get("SENDER_EMAIL", "") if hasattr(st, "secrets") else ""
    default_pass = st.secrets.get("APP_PASSWORD", "") if hasattr(st, "secrets") else ""

    sender_name = st.text_input("From Name", value="Me")
    sender_email = st.text_input("From Email (Gmail recommended)", value=default_sender, placeholder="you@gmail.com")
    app_password = st.text_input("Gmail App Password", type="password", value=default_pass, placeholder="16-character app password")

    st.markdown(
        "Get an **App Password**:\n\n"
        "1) Two-factor auth must be enabled\n"
        "2) Visit: https://myaccount.google.com/apppasswords\n"
        "3) App: **Mail**, Device: **Other** → copy the 16-char password"
    )

    st.divider()
    st.caption("Advanced SMTP (optional)")
    smtp_server = st.text_input("SMTP Server", value="smtp.gmail.com")
    smtp_port = st.number_input("SMTP Port", value=587, step=1)
    use_tls = st.toggle("Use TLS", value=True)

# =========================================
# INPUT SOURCE
# =========================================
st.subheader("1) Choose source")

upload_mode = st.radio(
    "How do you want to provide PDFs?",
    ["Upload files / ZIP (web-safe)", "Local folder (desktop only)"],
    horizontal=False
)

uploads = []
folder_path = ""

if upload_mode == "Upload files / ZIP (web-safe)":
    uploads = st.file_uploader(
        "Drag & drop multiple PDFs, or a single ZIP containing PDFs",
        type=["pdf", "zip"],
        accept_multiple_files=True
    )
else:
    folder_path = st.text_input("Local folder path (e.g., C:\\invoices or /Users/me/pdfs)")
    st.caption("Works only when running Streamlit locally on your machine. Not supported on Streamlit Cloud.")

# =========================================
# RECIPIENT (AUTO SUBJECT/BODY)
# =========================================
st.subheader("2) Recipient")
recipient = st.text_input("Send all emails to this recipient", value="", placeholder="recipient@example.com")

# Automatic subject/body based on filename
AUTO_BODY_TEMPLATE = "Hi,\n\nPlease find attached: {filename}.\n\nBest regards."

def build_subject_auto(filename: str) -> str:
    return filename[:-4] if filename.lower().endswith(".pdf") else filename

def build_body_auto(filename: str) -> str:
    base = filename[:-4] if filename.lower().endswith(".pdf") else filename
    return AUTO_BODY_TEMPLATE.format(filename=base)

dry_run = st.checkbox("Dry run (don’t actually send; just preview)", value=False)

st.divider()
send_btn = st.button("🚀 Send emails", type="primary", use_container_width=True)

# =========================================
# HELPERS
# =========================================
def is_pdf(name: str) -> bool:
    return name.lower().endswith(".pdf")

def list_pdfs_from_uploads(files):
    """
    Accepts a list of UploadedFile objects (PDF or ZIP),
    returns a list of dicts: [{name: str, data: bytes}, ...] for PDFs only.
    """
    results = []
    for f in files or []:
        fname = f.name
        if is_pdf(fname):
            results.append({"name": fname, "data": f.read()})
        elif fname.lower().endswith(".zip"):
            try:
                b = io.BytesIO(f.read())
                with zipfile.ZipFile(b) as z:
                    for info in z.infolist():
                        if not info.is_dir() and is_pdf(info.filename):
                            with z.open(info, "r") as pdf_file:
                                results.append({"name": os.path.basename(info.filename), "data": pdf_file.read()})
            except zipfile.BadZipFile:
                st.warning(f"⚠️ '{fname}' is not a valid ZIP file.")
    return results

def list_pdfs_from_folder(path: str):
    """
    Reads all PDFs from a local folder (desktop-only).
    """
    results = []
    if path and os.path.isdir(path):
        for name in os.listdir(path):
            if is_pdf(name):
                try:
                    with open(os.path.join(path, name), "rb") as f:
                        results.append({"name": name, "data": f.read()})
                except Exception as e:
                    st.warning(f"Could not read '{name}': {e}")
    return results

def collect_pdfs(upload_mode, uploads, folder_path):
    if upload_mode == "Upload files / ZIP (web-safe)":
        return list_pdfs_from_uploads(uploads)
    return list_pdfs_from_folder(folder_path)

def send_one_email(
    smtp, sender_email, sender_name, recipient, subject, body, attachment_name, attachment_bytes
):
    msg = MIMEMultipart()
    msg["From"] = formataddr((sender_name, sender_email))
    msg["To"] = recipient
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    part = MIMEApplication(attachment_bytes, _subtype="pdf")
    part.add_header("Content-Disposition", "attachment", filename=attachment_name)
    msg.attach(part)

    smtp.send_message(msg)

# =========================================
# ACTION
# =========================================
if send_btn:
    errors = []
    if not recipient:
        errors.append("Please enter a recipient email.")
    if upload_mode == "Upload files / ZIP (web-safe)" and not uploads:
        errors.append("Please upload at least one PDF or a ZIP containing PDFs.")
    if upload_mode == "Local folder (desktop only)" and not folder_path:
        errors.append("Please provide a local folder path.")
    if not app_password and not dry_run:
        errors.append("Please enter your Gmail App Password (or enable Dry run).")
    if not sender_email:
        errors.append("Please enter your sender email.")

    pdfs = collect_pdfs(upload_mode, uploads, folder_path)

    if (uploads or folder_path) and not pdfs:
        errors.append("No PDFs detected. Upload PDFs directly / ZIPs or point to a folder with PDFs.")

    # Gmail ~25MB limit per email (attachment slightly less)
    too_big = [p for p in pdfs if len(p["data"]) > 24 * 1024 * 1024]
    if too_big:
        st.warning("Some PDFs exceed ~24MB and may fail to send on Gmail. Consider splitting or compressing.")

    if errors:
        for e in errors:
            st.error(f"• {e}")
    else:
        st.info(f"Found {len(pdfs)} PDF(s) to send to **{recipient}**.")
        results = []

        try:
            smtp = None
            if not dry_run:
                smtp = smtplib.SMTP(smtp_server, smtp_port, timeout=60)
                if use_tls:
                    smtp.starttls()
                smtp.login(sender_email, app_password)

            for i, pdf in enumerate(pdfs, start=1):
                filename = pdf["name"]
                subject = build_subject_auto(filename)      # auto: filename (no .pdf)
                body = build_body_auto(filename)            # auto: includes filename

                if dry_run:
                    results.append({
                        "File": filename,
                        "Subject": subject,
                        "Body (first 80 chars)": (body[:80] + "…") if len(body) > 80 else body,
                        "Status": "DRY RUN ✅ (not sent)"
                    })
                else:
                    try:
                        send_one_email(
                            smtp=smtp,
                            sender_email=sender_email,
                            sender_name=sender_name,
                            recipient=recipient,
                            subject=subject,
                            body=body,
                            attachment_name=filename,
                            attachment_bytes=pdf["data"],
                        )
                        results.append({"File": filename, "Subject": subject, "Status": "Sent ✅"})
                        # small delay to be polite with SMTP servers
                        time.sleep(0.3)
                    except Exception as e:
                        results.append({"File": filename, "Subject": subject, "Status": f"Failed ❌: {e}"})

            if smtp:
                try:
                    smtp.quit()
                except Exception:
                    pass

            st.success("Done.")
            st.dataframe(results, use_container_width=True)
        except Exception as e:
            st.error(f"Unexpected error: {e}")
