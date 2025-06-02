import os
import re
import tempfile
import pandas as pd
import streamlit as st
from datetime import datetime

# ReportLab for PDF export
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# ----------------------------------------
# CSV Parsing with Auto‐Column Detection
# ----------------------------------------
@st.cache_data
def parse_csv(file_obj):
    """
    Parse a CSV report with flexible column detection.
    Detects columns (case‐insensitive) for:
      - date (any column containing 'date' or 'sent')
      - subject (containing 'subject')
      - sender (containing 'from' or 'sender')
      - recipients: 'to', 'cc', 'bcc' (combines all found)
      - body (containing 'body', 'content', or 'text')
      - attachments (containing 'attach')

    Returns:
      - messages: list of dicts with fields:
          date (datetime or timestamp 0),
          subject, sender, recipients, body,
          emails_in_body, phones_in_body, attachments (list of filenames)
      - column_map: mapping of detected column names (for reference)
    """
    # Read the entire CSV into a DataFrame
    df = pd.read_csv(file_obj, encoding="utf-8")
    cols = list(df.columns)
    cols_lower = [c.lower() for c in cols]

    def find_col(substrings):
        for substr in substrings:
            for i, lower in enumerate(cols_lower):
                if substr in lower:
                    return cols[i]
        return None

    date_col = find_col(["deliverydate", "date", "sent"])
    subject_col = find_col(["subject"])
    sender_col = find_col(["from", "sender"])
    to_col = find_col([" to", " to ", " to$", "^to$", "recipient"])
    cc_col = find_col(["cc"])
    bcc_col = find_col(["bcc"])
    body_col = find_col(["body", "content", "text"])
    attach_col = find_col(["attach"])

    # Convert the detected date column to datetime if it exists
    if date_col:
        try:
            df[date_col] = pd.to_datetime(df[date_col])
        except Exception:
            df[date_col] = pd.to_datetime(df[date_col].astype(str), errors="coerce")

    messages = []
    for _, row in df.iterrows():
        # Date
        if date_col and not pd.isna(row.get(date_col)):
            date = row.get(date_col)
            if not isinstance(date, datetime):
                try:
                    date = pd.to_datetime(date).to_pydatetime()
                except Exception:
                    date = datetime.fromtimestamp(0)
        else:
            date = datetime.fromtimestamp(0)

        # Subject
        subject = str(row.get(subject_col, "")) if subject_col else ""

        # Sender
        sender = str(row.get(sender_col, "")) if sender_col else ""

        # Recipients: combine To, CC, BCC if present
        rec_list = []
        if to_col:
            rec_to = str(row.get(to_col, "")) or ""
            if rec_to:
                rec_list.append(rec_to)
        if cc_col:
            rec_cc = str(row.get(cc_col, "")) or ""
            if rec_cc:
                rec_list.append(rec_cc)
        if bcc_col:
            rec_bcc = str(row.get(bcc_col, "")) or ""
            if rec_bcc:
                rec_list.append(rec_bcc)
        recipients = ", ".join(rec_list)

        # Body
        body = str(row.get(body_col, "")) if body_col else ""

        # Extract email addresses and phone numbers from the body
        emails_in_body = re.findall(r"\b[\w\.-]+@[\w\.-]+\.\w+\b", body)
        phones_in_body = re.findall(
            r"(\+?\d{1,3}[-\.\s]?)?\(?\d{3}\)?[-\.\s]?\d{3}[-\.\s]?\d{4}", body
        )
        emails_in_body = list(set(emails_in_body))
        phones_in_body = ["".join(p) for p in phones_in_body]
        phones_in_body = list(set(phones_in_body))

        # Attachments
        attachments = []
        if attach_col:
            raw = str(row.get(attach_col, "")) or ""
            for fn in raw.split(";"):
                fn = fn.strip()
                if fn:
                    attachments.append(fn)

        messages.append({
            "date": date,
            "subject": subject,
            "sender": sender,
            "recipients": recipients,
            "body": body,
            "emails_in_body": ", ".join(emails_in_body),
            "phones_in_body": ", ".join(phones_in_body),
            "attachments": attachments,
        })

    column_map = {
        "date": date_col,
        "subject": subject_col,
        "sender": sender_col,
        "to": to_col,
        "cc": cc_col,
        "bcc": bcc_col,
        "body": body_col,
        "attachments": attach_col
    }
    return messages, column_map


# ----------------------------------------
# CSV / PDF Export Helpers
# ----------------------------------------
@st.cache_data
def generate_csv_download(filtered_messages):
    df = pd.DataFrame([{
        "Date": msg.get("date"),
        "Subject": msg.get("subject"),
        "Sender": msg.get("sender"),
        "Recipients": msg.get("recipients"),
        "EmailsInBody": msg.get("emails_in_body"),
        "PhonesInBody": msg.get("phones_in_body"),
        "Attachments": ";".join(msg.get("attachments", [])),
        "Body": msg.get("body"),
    } for msg in filtered_messages])

    if not df.empty and isinstance(df.loc[0, "Date"], datetime):
        df["Date"] = df["Date"].dt.strftime("%Y-%m-%d %H:%M:%S")

    return df.to_csv(index=False).encode("utf-8")


@st.cache_data
def generate_pdf_download(filtered_messages):
    buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    doc = SimpleDocTemplate(buffer.name, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()

    title = Paragraph("Filtered Email Report", styles["Title"])
    elements.append(title)
    elements.append(Spacer(1, 12))

    table_data = [[
        "Date",
        "Subject",
        "Sender",
        "Recipients",
        "Emails In Body",
        "Phones In Body",
        "Attachments"
    ]]

    for msg in filtered_messages:
        date_str = msg["date"].strftime("%Y-%m-%d %H:%M:%S") if msg.get("date") else ""
        attachments_text = ";".join(msg.get("attachments", []))

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
    Create a PDF (bytes) for a single message, with detailed view.
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
        for fn in msg["attachments"]:
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
# Streamlit Interface
# ----------------------------------------
st.title("Email Forensic Tool (CSV-Only)")

uploaded = st.file_uploader("Upload a CSV file", type=["csv"])

if uploaded:
    with st.spinner("Parsing CSV…"):
        try:
            messages, column_map = parse_csv(uploaded)
        except Exception as e:
            st.error(f"Failed to parse CSV: {e}")
            st.stop()

    st.write("**Detected columns:**")
    for key, col in column_map.items():
        st.write(f"- {key.capitalize()}: {col or '—'}")

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

        st.download_button(
            "Download Filtered as CSV",
            data=generate_csv_download(filtered),
            file_name="filtered_emails.csv",
            mime="text/csv"
        )
        st.download_button(
            "Download Filtered as PDF",
            data=generate_pdf_download(filtered),
            file_name="filtered_emails.pdf",
            mime="application/pdf"
        )

        st.write("## Message Details & Attachments")
        idx_list = disp_df.index.tolist()
        selected_index = st.selectbox("Select message by Index", options=idx_list)
        msg = [m for i, m in enumerate(messages) if i == selected_index][0]

        st.write(f"**Date:** {msg['date'].strftime('%Y-%m-%d %H:%M:%S')}")
        st.write(f"**Subject:** {msg['subject']}")
        st.write(f"**Sender:** {msg['sender']}")
        st.write(f"**Recipients:** {msg['recipients']}")

        if msg["attachments"]:
            st.write("**Attachments:**")
            for fn in msg["attachments"]:
                st.write(f"• {fn}")

        st.write("**Body:**")
        st.write(msg["body"])

        pdf_data = generate_single_pdf(msg)
        st.download_button(
            "Download This Communication's Report as PDF",
            data=pdf_data,
            file_name=f"message_{selected_index}.pdf",
            mime="application/pdf"
        )
    else:
        st.info("No messages match the current filters.")
