import os
import re
import tempfile
import zipfile
import pandas as pd
import streamlit as st
from datetime import datetime

# ReportLab for PDF export
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# extract_msg for parsing .msg files
import extract_msg

# ----------------------------------------
# .msg Parsing
# ----------------------------------------
@st.cache_data
 def parse_msg_files(_msg_files):
    """
    Parse a list of .msg files (UploadedFile-like objects).
    Returns (messages, attachments_storage).
    - messages: list of dicts with keys:
        date (datetime), subject, sender, recipients, body,
        emails_in_body, phones_in_body, attachments (list of (filename, key))
    - attachments_storage: dict mapping key -> raw bytes
    """
    messages = []
    attachments_storage = {}

    for uploaded in _msg_files:
        # Write to temp file so extract_msg can process
        with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
            tmp.write(uploaded.read())
            tmp_path = tmp.name

        msg = extract_msg.Message(tmp_path)
        msg_sender  = msg.sender or ""
        msg_subject = msg.subject or ""

        try:
            msg_date = msg.date
            if isinstance(msg_date, str):
                msg_date = datetime.strptime(msg_date, "%m/%d/%Y %I:%M:%S %p")
        except Exception:
            msg_date = datetime.fromtimestamp(0)

        to_field  = msg.to or ""
        cc_field  = msg.cc or ""
        bcc_field = msg.bcc or ""
        recipients = ", ".join(filter(None, [to_field, cc_field, bcc_field]))

        body = msg.body or ""
        emails_in_body = re.findall(r"\b[\w\.-]+@[\w\.-]+\.\w+\b", body)
        phones_in_body = re.findall(
            r"(\+?\d{1,3}[-\.\s]?)?\(?\d{3}\)?[-\.\s]?\d{3}[-\.\s]?\d{4}", body
        )
        emails_in_body = list(set(emails_in_body))
        phones_in_body = ["".join(p) for p in phones_in_body]
        phones_in_body = list(set(phones_in_body))

        attachments = []
        for att in msg.attachments:
            fname = att.longFilename or att.shortFilename or "attachment"
            data  = att.data
            key   = f"{len(attachments_storage)}_{fname}"
            attachments_storage[key] = data
            attachments.append((fname, key))

        messages.append({
            "date": msg_date,
            "subject": msg_subject,
            "sender": msg_sender,
            "recipients": recipients,
            "body": body,
            "emails_in_body": ", ".join(emails_in_body),
            "phones_in_body": ", ".join(phones_in_body),
            "attachments": attachments,
        })

        msg.close()
        os.unlink(tmp_path)

    return messages, attachments_storage

# ----------------------------------------
# ZIP Handling: extract .msg files inside
# ----------------------------------------
@st.cache_data
 def parse_zip_file(uploaded_zip):
    """
    Extract all .msg files from uploaded ZIP and parse them.
    Returns (messages, attachments_storage).
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, uploaded_zip.name)
        with open(zip_path, "wb") as f:
            f.write(uploaded_zip.read())

        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(tmpdir)

        msg_file_paths = []
        for root, _, files in os.walk(tmpdir):
            for name in files:
                if name.lower().endswith(".msg"):
                    msg_file_paths.append(os.path.join(root, name))

        class _TmpUploaded:
            def __init__(self, path):
                self.name = os.path.basename(path)
                self._path = path
            def read(self):
                with open(self._path, "rb") as f:
                    return f.read()

        msg_uploads = [_TmpUploaded(p) for p in msg_file_paths]
        return parse_msg_files(msg_uploads)

# ----------------------------------------
# CSV / PDF Export Helpers
# ----------------------------------------
@st.cache_data
 def generate_csv_download(messages):
    df = pd.DataFrame([{
        "Date": msg.get("date"),
        "Subject": msg.get("subject"),
        "Sender": msg.get("sender"),
        "Recipients": msg.get("recipients"),
        "EmailsInBody": msg.get("emails_in_body"),
        "PhonesInBody": msg.get("phones_in_body"),
        "Attachments": ";".join([
            att if isinstance(att, str) else att[0]
            for att in msg.get("attachments", [])
        ]),
        "Body": msg.get("body"),
    } for msg in messages])

    if not df.empty and isinstance(df.loc[0, "Date"], datetime):
        df["Date"] = df["Date"].dt.strftime("%Y-%m-%d %H:%M:%S")

    return df.to_csv(index=False).encode("utf-8")

@st.cache_data
 def generate_pdf_download(messages):
    buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    doc = SimpleDocTemplate(buffer.name, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()

    title = Paragraph("Filtered Email Report", styles["Title"])
    elements.append(title)
    elements.append(Spacer(1, 12))

    table_data = [[
        "Date", "Subject", "Sender", "Recipients",
        "Emails In Body", "Phones In Body", "Attachments"
    ]]

    for msg in messages:
        date_str        = msg["date"].strftime("%Y-%m-%d %H:%M:%S") if msg.get("date") else ""
        attachments_text = ";".join([
            att if isinstance(att, str) else att[0]
            for att in msg.get("attachments", [])
        ])

        row = [
            date_str,
            msg.get("subject", ""),
            msg.get("sender", ""),
            msg.get("recipients", ""),
            msg.get("emails_in_body", ""),
            msg.get("phones_in_body", ""),
            attachments_text,
        ]
        table_data.append(row)

    tbl = Table(table_data, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    elements.append(tbl)

    doc.build(elements)
    with open(buffer.name, "rb") as f:
        data = f.read()
    os.unlink(buffer.name)
    return data

 def generate_single_pdf(msg):
    buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    doc = SimpleDocTemplate(buffer.name, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()

    elements.append(Paragraph(f"Date: {msg['date'].strftime('%Y-%m-%d %H:%M:%S')}", styles["Normal"]))
    elements.append(Paragraph(f"Subject: {msg['subject']}", styles["Normal"]))
    elements.append(Paragraph(f"Sender: {msg['sender']}", styles["Normal"]))
    elements.append(Paragraph(f"Recipients: {msg['recipients']}", styles["Normal"]))

    if msg["attachments"]:
        elements.append(Spacer(1, 12))
        elements.append(Paragraph("Attachments:", styles["Normal"]))
        for fn in msg["attachments"]:
            elements.append(Paragraph(f"• {fn}", styles["Normal"]))

    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Body:", styles["Normal"]))
    for line in msg["body"].split("\n"):
        if line.strip() == "":
            elements.append(Spacer(1, 6))
        else:
            elements.append(Paragraph(line, styles["Normal"]))

    doc.build(elements)
    with open(buffer.name, "rb") as f:
        data = f.read()
    os.unlink(buffer.name)
    return data

# ----------------------------------------
# Streamlit Interface
# ----------------------------------------
 st.title("Email Forensic Tool (.msg ZIP)")

 st.write(
     "Upload a ZIP containing `.msg` files (eDiscovery export). "
     "The app will extract and parse every .msg, then present a searchable table, "
     "including each email’s body and any attachments."
 )

 uploaded = st.file_uploader(
     "Upload a ZIP of .msg files (single ZIP only)", type=["zip"]
 )

 if uploaded:
     with st.spinner("Extracting and parsing .msg files…"):
         try:
             messages, attachments_storage = parse_zip_file(uploaded)
         except Exception as e:
             st.error(f"Failed to parse ZIP: {e}")
             st.stop()

     if not messages:
         st.info("No `.msg` files were found in the uploaded ZIP.")
         st.stop()

     df = pd.DataFrame([{ 
         "Date": msg.get("date"),
         "Subject": msg.get("subject"),
         "Sender": msg.get("sender"),
         "Recipients": msg.get("recipients"),
         "EmailsInBody": msg.get("emails_in_body"),
         "PhonesInBody": msg.get("phones_in_body"),
