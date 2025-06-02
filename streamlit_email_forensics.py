import os
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

# msg_parser for parsing .msg files (incl. RTF→HTML)
from msg_parser import MsOxMessage

# ----------------------------------------
# .msg Parsing via msg_parser
# ----------------------------------------
@st.cache_data
def parse_msg_files(_msg_files):
    """
    Parse a list of .msg files (UploadedFile-like objects) using msg_parser.
    Returns (messages, attachments_storage).
    - messages: list of dicts with keys:
        date (datetime), subject, sender, recipients, body,
        emails_in_body, phones_in_body, attachments (list of (filename, key))
    - attachments_storage: dict mapping key -> raw bytes
    """
    messages = []
    attachments_storage = {}

    for uploaded in _msg_files:
        # Save UploadedFile to a temp .msg
        with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
            tmp.write(uploaded.read())
            tmp_path = tmp.name

        # Parse with msg_parser
        msg = MsOxMessage(tmp_path)
        props = msg.get_properties()

        # Date
        msg_date = props.get("DeliveryTime") or props.get("SentOn") or None
        if isinstance(msg_date, str):
            try:
                msg_date = datetime.fromisoformat(msg_date)
            except Exception:
                msg_date = None
        if not isinstance(msg_date, datetime):
            msg_date = None

        # Subject, Sender
        msg_subject = props.get("Subject", "")
        msg_sender = props.get("SenderName", "") or props.get("FromDisplayName", "")

        # Recipients: 'To', 'Cc', 'Bcc'
        to_list = props.get("To", []) or []
        cc_list = props.get("Cc", []) or []
        bcc_list = props.get("Bcc", []) or []
        recipients = ", ".join(to_list + cc_list + bcc_list)

        # Body: prefer HTML, fallback to plain text
        html_body = props.get("Html", "").strip() or None
        text_body = props.get("Body", "").strip() or None
        # Use HTML if available; otherwise plain text
        body = html_body or text_body or ""

        # Extract email addresses & phone numbers from the body text
        import re
        emails_in_body = re.findall(r"\b[\w\.-]+@[\w\.-]+\.\w+\b", body)
        phones_in_body = re.findall(
            r"(\+?\d{1,3}[-\.\s]?)?\(?\d{3}\)?[-\.\s]?\d{3}[-\.\s]?\d{4}", body
        )
        emails_in_body = list(dict.fromkeys(emails_in_body))
        phones_in_body = ["".join(p) for p in phones_in_body]
        phones_in_body = list(dict.fromkeys(phones_in_body))

        # Attachments: msg_parser returns a list of dicts with 'filename' and 'data'
        for att in msg.attachments:
            fname = att.get("filename", "attachment")
            data = att.get("content", b"")
            key = f"{len(attachments_storage)}_{fname}"
            attachments_storage[key] = data

        attachment_list = [(att["filename"], f"{i}_{att['filename']}") 
                           for i, att in enumerate(msg.attachments)]

        messages.append({
            "date": msg_date,
            "subject": msg_subject,
            "sender": msg_sender,
            "recipients": recipients,
            "body": body,
            "emails_in_body": ", ".join(emails_in_body),
            "phones_in_body": ", ".join(phones_in_body),
            "attachments": attachment_list,
        })

        msg.close()
        os.unlink(tmp_path)

    return messages, attachments_storage


# ----------------------------------------
# ZIP Handling: extract .msg files
# ----------------------------------------
@st.cache_data
def parse_zip_file(uploaded_zip):
    """
    Extract all .msg files from the uploaded ZIP and parse them via msg_parser.
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
        "Attachments": ";".join([att[0] for att in msg.get("attachments", [])]),
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
        date_str = msg["date"].strftime("%Y-%m-%d %H:%M:%S") if msg.get("date") else ""
        attachments_text = ";".join([att[0] for att in msg.get("attachments", [])])

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
    """
    Create a PDF (bytes) for a single message, formatted as a detailed report.
    """
    buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    doc = SimpleDocTemplate(buffer.name, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()

    # Headers
    elements.append(Paragraph(f"Date: {msg['date'].strftime('%Y-%m-%d %H:%M:%S')}", styles["Normal"]))
    elements.append(Paragraph(f"Subject: {msg['subject']}", styles["Normal"]))
    elements.append(Paragraph(f"Sender: {msg['sender']}", styles["Normal"]))
    elements.append(Paragraph(f"Recipients: {msg['recipients']}", styles["Normal"]))

    if msg["attachments"]:
        elements.append(Spacer(1, 12))
        elements.append(Paragraph("Attachments:", styles["Normal"]))
        for fn, _ in msg["attachments"]:
            elements.append(Paragraph(f"• {fn}", styles["Normal"]))

    # Body
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
# Split-Zip Helper for Attachments
# ----------------------------------------
def create_split_zips(attachments_storage, size_limit=190 * 1024 * 1024):
    """
    Given attachments_storage (key -> bytes), create multiple ZIP files beneath size_limit.
    Returns list of (zip_filename, zip_bytes).
    """
    items = list(attachments_storage.items())
    zips = []
    current_zip_file = None
    current_zip = None
    current_size = 0
    part_index = 1

    def finalize_zip(zf, tmp_path):
        zf.close()
        with open(tmp_path, "rb") as f:
            data = f.read()
        os.unlink(tmp_path)
        return data

    # Start first ZIP
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    current_zip_file = tmp.name
    zf = zipfile.ZipFile(current_zip_file, mode="w", compression=zipfile.ZIP_DEFLATED)
    current_size = 0

    for key, data in items:
        fname = key.split("_", 1)[1]
        file_size = len(data)
        # If adding this would exceed limit, finalize current and start a new one
        if current_size + file_size > size_limit and current_size > 0:
            zip_bytes = finalize_zip(zf, current_zip_file)
            zips.append((f"attachments_part{part_index}.zip", zip_bytes))
            part_index += 1
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
            current_zip_file = tmp.name
            zf = zipfile.ZipFile(current_zip_file, mode="w", compression=zipfile.ZIP_DEFLATED)
            current_size = 0

        zf.writestr(fname, data)
        current_size += file_size

    # Finalize last
    if zf:
        zip_bytes = finalize_zip(zf, current_zip_file)
        zips.append((f"attachments_part{part_index}.zip", zip_bytes))

    return zips


# ----------------------------------------
# Streamlit Interface
# ----------------------------------------
st.title("Email Forensic Tool (.msg ZIP)")

st.write(
    "Upload a ZIP containing `.msg` files (eDiscovery export). "
    "The app will extract and parse every .msg using msg_parser (so bodies appear correctly), "
    "then present a searchable table including each email’s body and any attachments. "
    "You can also download all attachments split into multiple ZIP parts under 190 MB each."
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

    # Build DataFrame for display/filtering
    df = pd.DataFrame([{
        "Date": msg.get("date"),
        "Subject": msg.get("subject"),
        "Sender": msg.get("sender"),
        "Recipients": msg.get("recipients"),
        "EmailsInBody": msg.get("emails_in_body"),
        "PhonesInBody": msg.get("phones_in_body"),
        "AttachmentsCount": len(msg.get("attachments", [])),
        "Index": i
    } for i, msg in enumerate(messages)])

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    # Sidebar filters
    st.sidebar.header("Filters")
    subj_filter = st.sidebar.text_input("Subject contains")
    sender_filter = st.sidebar.text_input("Sender contains")
    rec_filter = st.sidebar.text_input("Communicated with (email/domain)")
    email_filter = st.sidebar.text_input("Email in body contains")
    phone_filter = st.sidebar.text_input("Phone in body contains")
    body_filter = st.sidebar.text_input("Body contains (any text/address/etc.)")
    has_attach = st.sidebar.checkbox("Only show messages with attachments")
    start_date = st.sidebar.date_input("Start date", value=datetime(2000, 1, 1).date())
    end_date = st.sidebar.date_input("End date", value=datetime.today().date())

    # Apply filters
    filtered = []
    for msg in messages:
        if subj_filter and subj_filter.lower() not in msg["subject"].lower():
            continue
        if sender_filter and sender_filter.lower() not in msg["sender"].lower():
            continue
        if rec_filter:
            low = rec_filter.lower()
            if low not in msg["sender"].lower() and low not in msg["recipients"].lower():
                continue
        if email_filter and email_filter.lower() not in msg["emails_in_body"].lower():
            continue
        if phone_filter and phone_filter.lower() not in msg["phones_in_body"].lower():
            continue
        if body_filter and body_filter.lower() not in msg["body"].lower():
            continue
        if has_attach and len(msg["attachments"]) == 0:
            continue
        msg_date = msg["date"]
        if msg_date:
            if msg_date.date() < start_date or msg_date.date() > end_date:
                continue
        filtered.append(msg)

    if filtered:
        disp_df = pd.DataFrame([{
            "Date": msg.get("date"),
            "Subject": msg.get("subject"),
            "Sender": msg.get("sender"),
            "Recipients": msg.get("recipients"),
            "EmailsInBody": msg.get("emails_in_body"),
            "PhonesInBody": msg.get("phones_in_body"),
            "AttachmentsCount": len(msg.get("attachments", [])),
            "Index": i
        } for i, msg in enumerate(messages) if msg in filtered])

        disp_df["Date"] = disp_df["Date"].dt.strftime("%Y-%m-%d %H:%M:%S")
        st.dataframe(disp_df.set_index("Index"), height=400)

        # Download filtered as CSV / PDF
        st.download_button(
            "Download Filtered as CSV",
            data=generate_csv_download(filtered),
            file_name="filtered_emails.csv",
            mime="text/csv",
            key="download_filtered_csv"
        )
        st.download_button(
            "Download Filtered as PDF",
            data=generate_pdf_download(filtered),
            file_name="filtered_emails.pdf",
            mime="application/pdf",
            key="download_filtered_pdf"
        )

        # Split & download all attachments for filtered messages
        filtered_keys = []
        for msg in filtered:
            for fname_key in msg["attachments"]:
                _, key = fname_key
                filtered_keys.append(key)

        if filtered_keys:
            filtered_attachments = {k: attachments_storage[k] for k in filtered_keys}
            parts = create_split_zips(filtered_attachments)
            st.write("## Download Attachments (Split ZIPs)")
            for part_name, part_bytes in parts:
                st.download_button(
                    f"Download {part_name}",
                    data=part_bytes,
                    file_name=part_name,
                    mime="application/zip",
                    key=f"download_attach_part_{part_name}"
                )
        else:
            st.info("No attachments to download for the filtered messages.")

        # Detailed view for a single message
        st.write("## Message Details & Attachments")
        idx_list = disp_df.index.tolist()
        selected_index = st.selectbox("Select message by Index", options=idx_list, key="select_message")
        msg = [m for i, m in enumerate(messages) if i == selected_index][0]

        st.write(f"**Date:** {msg['date'].strftime('%Y-%m-%d %H:%M:%S')}")
        st.write(f"**Subject:** {msg['subject']}")
        st.write(f"**Sender:** {msg['sender']}")
        st.write(f"**Recipients:** {msg['recipients']}")

        if msg["attachments"]:
            st.write("**Attachments:**")
            for att_index, att in enumerate(msg["attachments"]):
                fn, key = att
                data = attachments_storage.get(key)
                if data:
                    st.download_button(
                        f"Download {fn}",
                        data=data,
                        file_name=fn,
                        key=f"download_attach_{selected_index}_{att_index}"
                    )

        st.write("**Body:**")
        st.write(msg["body"])

        # Download this message as PDF
        single_pdf = generate_single_pdf(msg)
        st.download_button(
            "Download This Message as PDF",
            data=single_pdf,
            file_name=f"message_{selected_index}.pdf",
            mime="application/pdf",
            key=f"download_msg_pdf_{selected_index}"
        )
    else:
        st.info("No messages match the current filters.")
