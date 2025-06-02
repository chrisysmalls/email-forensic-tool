import os
import re
import tempfile
import pandas as pd
import streamlit as st
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# ----------------------------------------
# Helper functions to parse CSV and build downloads
# ----------------------------------------

@st.cache_data
def parse_csv(file_obj):
    """
    Parse a CSV report (e.g., from Microsoft eDiscovery or a CSV you exported yourself).
    Expects columns: DeliveryDate, Subject, From, To, CC, BCC, Body, [Attachments].
    If 'Attachments' column is present, it should be a semicolon-separated list of filenames.
    """
    # Peek at columns to see if DeliveryDate exists
    df_header = pd.read_csv(file_obj, nrows=0)
    parse_dates = ["DeliveryDate"] if "DeliveryDate" in df_header.columns else None

    # Rewind the file buffer after peeking
    file_obj.seek(0)
    df = pd.read_csv(file_obj, encoding="utf-8", errors="ignore", parse_dates=parse_dates)

    messages = []
    # If there is an 'Attachments' column, split by ';' and store as list
    has_attachments_col = "Attachments" in df.columns

    for _, row in df.iterrows():
        # Subject, sender, recipients
        subject = row.get("Subject", "") or ""
        sender = row.get("From", "") or ""
        rec_to = row.get("To", "") or ""
        rec_cc = row.get("CC", "") or ""
        rec_bcc = row.get("BCC", "") or ""
        recipients = ", ".join(filter(None, [rec_to, rec_cc, rec_bcc]))

        # Delivery date (if present)
        if "DeliveryDate" in row and not pd.isna(row.get("DeliveryDate")):
            date = pd.to_datetime(row.get("DeliveryDate")).to_pydatetime()
        else:
            date = datetime.fromtimestamp(0)

        # Body
        body = row.get("Body", "") or ""

        # Extract email addresses & phone numbers from body
        emails_in_body = re.findall(r"\b[\w\.-]+@[\w\.-]+\.\w+\b", body)
        phones_in_body = re.findall(
            r"(\+?\d{1,3}[-\.\s]?)?\(?\d{3}\)?[-\.\s]?\d{3}[-\.\s]?\d{4}", body
        )
        emails_in_body = list(set(emails_in_body))
        # Join phone‐number regex groups back into a single string
        phones_in_body = ["".join(p) for p in phones_in_body]
        phones_in_body = list(set(phones_in_body))

        # Attachments (if provided in CSV)
        attachments = []
        if has_attachments_col:
            raw = row.get("Attachments", "") or ""
            # Assume semicolon-separated filenames
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

    return messages

@st.cache_data
def generate_csv_download(filtered_messages):
    """
    Build a CSV (bytes) from filtered_messages list of dicts.
    """
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
    """
    Create a PDF (bytes) listing all filtered_messages.
    """
    buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    doc = SimpleDocTemplate(buffer.name, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()

    # Title
    title = Paragraph("Filtered Email Report", styles["Title"])
    elements.append(title)
    elements.append(Spacer(1, 12))

    # Table header + rows
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
        date_str = msg.get("date").strftime("%Y-%m-%d %H:%M:%S") if msg.get("date") else ""
        row = [
            date_str,
            msg.get("subject", ""),
            msg.get("sender", ""),
            msg.get("recipients", ""),
            msg.get("emails_in_body", ""),
            msg.get("phones_in_body", ""),
            ";".join(msg.get("attachments", [])),
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

    # Build PDF
    doc.build(elements)

    with open(buffer.name, "rb") as f:
        data = f.read()
    os.unlink(buffer.name)
    return data


# ----------------------------------------
# Streamlit UI
# ----------------------------------------

st.title("Email Forensic Tool (CSV-only)")

uploaded_file = st.file_uploader(
    "Upload a CSV file (from eDiscovery or your local PST→CSV export)",
    type=["csv"]
)

if uploaded_file:
    # Parse CSV
    try:
        with st.spinner("Parsing CSV…"):
            messages = parse_csv(uploaded_file)
    except Exception as e:
        st.error(f"Error parsing CSV: {e}")
        st.stop()

    # Build a DataFrame for display & filtering
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

    # Sidebar filters
    st.sidebar.header("Filters")
    subj_filter = st.sidebar.text_input("Subject contains")
    sender_filter = st.sidebar.text_input("Sender contains")
    rec_filter = st.sidebar.text_input("Recipient contains")
    email_filter = st.sidebar.text_input("Email in body contains")
    phone_filter = st.sidebar.text_input("Phone in body contains")
    start_date = st.sidebar.date_input("Start date", value=datetime(2000, 1, 1).date())
    end_date = st.sidebar.date_input("End date", value=datetime.today().date())

    # Apply filters
    filtered = []
    for msg in messages:
        if subj_filter and subj_filter.lower() not in msg.get("subject", "").lower():
            continue
        if sender_filter and sender_filter.lower() not in msg.get("sender", "").lower():
            continue
        if rec_filter and rec_filter.lower() not in msg.get("recipients", "").lower():
            continue
        if email_filter and email_filter.lower() not in msg.get("emails_in_body", "").lower():
            continue
        if phone_filter and phone_filter.lower() not in msg.get("phones_in_body", "").lower():
            continue
        msg_date = msg.get("date")
        if msg_date:
            if msg_date.date() < start_date or msg_date.date() > end_date:
                continue
        filtered.append(msg)

    if filtered:
        # Show filtered DataFrame
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

        # Format Date column
        disp_df["Date"] = disp_df["Date"].dt.strftime("%Y-%m-%d %H:%M:%S")

        st.dataframe(disp_df.set_index("Index"))

        # Download buttons
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

        # Show individual message details & attachments
        st.write("## Message Details and Attachments")

        idx_list = disp_df.index.tolist()
        if idx_list:
            selected_index = st.selectbox(
                "Select the Index of the message to view details",
                options=idx_list
            )

            msg = [m for i, m in enumerate(messages) if i == selected_index][0]
            st.write(f"**Date:** {msg.get('date').strftime('%Y-%m-%d %H:%M:%S')}")
            st.write(f"**Subject:** {msg.get('subject')}")
            st.write(f"**Sender:** {msg.get('sender')}")
            st.write(f"**Recipients:** {msg.get('recipients')}")

            # Attachments (if any)
            if msg.get("attachments"):
                st.write("**Attachments:**")
                for filename in msg.get("attachments"):
                    # Since we don’t have binary data in CSV, we can only show the filename
                    st.write(f"• {filename}")

            st.write("**Body:**")
            st.write(msg.get("body", ""))

    else:
        st.info("No messages match the current filters.")
