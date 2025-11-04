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

# ======= HARD-CODED ADDRESSES =======
SENDER_NAME = "LimoExpress CPH"
SENDER_EMAIL = "limoexpresscph@gmail.com"
RECIPIENT = "849bilag1790627@e-conomic.dk"

# ======= SMTP (GMAIL) SETTINGS =======
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587      # TLS
USE_TLS = True

# ======= APP UI =======
st.set_page_config(page_title="Multiple files to revisor", page_icon="📧", layout="centered")
st.title("📧 Multiple files to revisor")
st.caption("Upload PDFs (or a ZIP). Each PDF is emailed separately to the accounting address.")

# Read App Password from secrets if available
default_pass = ""
try:
    default_pass = st.secrets.get("APP_PASSWORD", "")
except Exception:
    pass

with st.sidebar:
    st.header("Sender (hard-coded)")
    st.write(f"**From:** {SENDER_NAME} <{SENDER_EMAIL}>")
    st.write(f"**To:** {RECIPIENT}")
    APP_PASSWORD = st.text_input("Gmail App Password", value=default_pass, type="password",
                                 placeholder="16-character app password")

    st.markdown(
        "Get an **App Password**:\n"
        "1) Turn on 2-Step Verification\n"
        "2) Visit https://myaccount.google.com/apppasswords\n"
        "3) App: **Mail**, Device: **Other** → copy the 16-char password"
    )

st.subheader("Upload PDFs (or a ZIP of PDFs)")
uploads = st.file_uploader(
    "Drag & drop multiple PDFs, or a single ZIP containing PDFs",
    type=["pdf", "zip"],
    accept_multiple_files=True
)

dry_run = st.checkbox("Dry run (preview only, do not send)", value=False)
send_btn = st.button("🚀 Send emails", type="primary", use_container_width=True)

# ======= HELPERS =======
def is_pdf(name: str) -> bool:
    return name.lower().endswith(".pdf")

def collect_pdfs(files):
    """Return [{'name': str, 'data': bytes}, ...] for PDFs (including those inside ZIPs)."""
    out = []
    for f in files or []:
        if is_pdf(f.name):
            out.append({"name": f.name, "data": f.read()})
        elif f.name.lower().endswith(".zip"):
            try:
                buf = io.BytesIO(f.read())
                with zipfile.ZipFile(buf) as z:
                    for info in z.infolist():
                        if not info.is_dir() and is_pdf(info.filename):
                            with z.open(info, "r") as pdf_file:
                                out.append({"name": os.path.basename(info.filename), "data": pdf_file.read()})
            except zipfile.BadZipFile:
                st.warning(f"⚠️ '{f.name}' is not a valid ZIP file.")
    return out

def subject_from_filename(filename: str) -> str:
    return filename[:-4] if filename.lower().endswith(".pdf") else filename

def body_from_filename(filename: str) -> str:
    base = subject_from_filename(filename)
    return f"Hi,\n\nPlease find attached: {base}.\n\nBest regards,\n{SENDER_NAME}"

def send_one_email(smtp, attachment_name: str, attachment_bytes: bytes):
    msg = MIMEMultipart()
    msg["From"] = formataddr((SENDER_NAME, SENDER_EMAIL))
    msg["To"] = RECIPIENT
    msg["Subject"] = subject_from_filename(attachment_name)
    msg.attach(MIMEText(body_from_filename(attachment_name), "plain"))

    part = MIMEApplication(attachment_bytes, _subtype="pdf")
    part.add_header("Content-Disposition", "attachment", filename=attachment_name)
    msg.attach(part)

    smtp.send_message(msg)

# ======= ACTION =======
if send_btn:
    errors = []
    pdfs = collect_pdfs(uploads)
    if not uploads:
        errors.append("Please upload at least one PDF or a ZIP containing PDFs.")
    if not pdfs:
        errors.append("No PDFs detected in your upload.")
    if not APP_PASSWORD and not dry_run:
        errors.append("Enter the Gmail App Password or enable Dry run.")

    # Gmail limit ~25MB per email (slightly less for attachment)
    too_big = [p for p in pdfs if len(p["data"]) > 24 * 1024 * 1024]
    if too_big:
        st.warning("Some PDFs exceed ~24MB and may fail to send via Gmail. Consider compressing/splitting.")

    if errors:
        for e in errors:
            st.error(f"• {e}")
    else:
        st.info(f"Preparing to send {len(pdfs)} email(s) from {SENDER_EMAIL} to {RECIPIENT}.")
        results = []
        smtp = None

        try:
            if not dry_run:
                smtp = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=60)
                if USE_TLS:
                    smtp.starttls()
                smtp.login(SENDER_EMAIL, APP_PASSWORD)

            for i, pdf in enumerate(pdfs, start=1):
                name = pdf["name"]
                if dry_run:
                    results.append({
                        "File": name,
                        "Subject": subject_from_filename(name),
                        "Status": "DRY RUN ✅ (not sent)"
                    })
                else:
                    try:
                        send_one_email(smtp, name, pdf["data"])
                        results.append({"File": name, "Subject": subject_from_filename(name), "Status": "Sent ✅"})
                        time.sleep(0.3)  # gentle pacing
                    except Exception as e:
                        results.append({"File": name, "Subject": subject_from_filename(name), "Status": f"Failed ❌: {e}"})
        except Exception as e:
            st.error(f"SMTP error: {e}")
        finally:
            if smtp:
                try:
                    smtp.quit()
                except Exception:
                    pass

        st.success("Done.")
        st.dataframe(results, use_container_width=True)

