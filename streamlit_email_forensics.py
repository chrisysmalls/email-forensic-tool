import os
import re
import tempfile
import pandas as pd
import streamlit as st
from datetime import datetime
import pypff
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

def parse_pst(file_obj):
    pst_file = pypff.file()
    pst_file.open_file_object(file_obj)
    root = pst_file.get_root_folder()
    messages = []
    attachments_storage = {}

    def traverse_folder(folder, folder_path=""):
        for i in range(folder.get_number_of_sub_folders()):
            sub_folder = folder.get_sub_folder(i)
            traverse_folder(sub_folder, folder_path + "/" + sub_folder.get_name())
        for j in range(folder.get_number_of_sub_messages()):
            msg = folder.get_sub_message(j)
            try:
                subject = msg.subject or ""
            except Exception:
                subject = ""
            try:
                sender_name = msg.sender_name or ""
                sender_email = msg.sender_email_address or ""
                sender = f"{sender_name} <{sender_email}>" if sender_name or sender_email else ""
            except Exception:
                sender = ""
            recipients_list = []
            try:
                num_rec = msg.number_of_recipients
                for k in range(num_rec):
                    rec = msg.get_recipient(k)
                    rec_name = rec.name or ""
                    rec_email = rec.email_address or ""
                    if rec_name or rec_email:
                        recipients_list.append(f"{rec_name} <{rec_email}>")
            except Exception:
                pass
            recipients = ", ".join(recipients_list)
            try:
                dt = msg.delivery_time
                date = dt if isinstance(dt, datetime) else datetime.fromtimestamp(0)
            except Exception:
                date = datetime.fromtimestamp(0)
            try:
                body_bytes = msg.get_plain_text_body() or b""
                body = body_bytes.decode("utf-8", errors="ignore")
            except Exception:
                try:
                    body = msg.get_transport_headers() or ""
                except Exception:
                    body = ""
            emails_in_body = re.findall(r"\b[\w\.-]+@[\w\.-]+\.\w+\b", body)
            phones_in_body = re.findall(r"(\+?\d{1,3}[-\.\s]?)?\(?\d{3}\)?[-\.\s]?\d{3}[-\.\s]?\d{4}", body)
            emails_in_body = list(set(emails_in_body))
            phones_in_body = ["".join(p) for p in phones_in_body]
            phones_in_body = list(set(phones_in_body))
            attachments = []
            try:
                num_att = msg.number_of_attachments
                for a in range(num_att):
                    att = msg.get_attachment(a)
                    filename = att.get_name() or f"attachment_{a}"
                    data = att.read_buffer()
                    key = f"{len(attachments_storage)}_{filename}"
                    attachments_storage[key] = data
                    attachments.append((filename, key))
            except Exception:
                pass
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
    traverse_folder(root)
    pst_file.close()
    return messages, attachments_storage

@st.cache_data
def parse_csv(file_obj):
    df = pd.read_csv(file_obj, encoding="utf-8", errors="ignore", parse_dates=["DeliveryDate"] if "DeliveryDate" in pd.read_csv(file_obj, nrows=0).columns else None)
    messages = []
    attachments_storage = {}
    for _, row in df.iterrows():
        subject = row.get("Subject", "") or ""
        sender = row.get("From", "") or ""
        rec_to = row.get("To", "") or ""
        rec_cc = row.get("CC", "") or ""
        rec_bcc = row.get("BCC", "") or ""
        recipients = ", ".join(filter(None, [rec_to, rec_cc, rec_bcc]))
        if "DeliveryDate" in row and not pd.isna(row.get("DeliveryDate")):
            date = pd.to_datetime(row.get("DeliveryDate")).to_pydatetime()
        else:
            date = datetime.fromtimestamp(0)
        body = row.get("Body", "") or ""
        emails_in_body = re.findall(r"\b[\w\.-]+@[\w\.-]+\.\w+\b", body)
        phones_in_body = re.findall(r"(\+?\d{1,3}[-\.\s]?)?\(?\d{3}\)?[-\.\s]?\d{3}[-\.\s]?\d{4}", body)
        emails_in_body = list(set(emails_in_body))
        phones_in_body = ["".join(p) for p in phones_in_body]
        phones_in_body = list(set(phones_in_body))
        messages.append({
            "date": date,
            "subject": subject,
            "sender": sender,
            "recipients": recipients,
            "body": body,
            "emails_in_body": ", ".join(emails_in_body),
            "phones_in_body": ", ".join(phones_in_body),
            "attachments": [],
        })
    return messages, attachments_storage

@st.cache_data
def generate_csv_download(filtered_messages):
    df = pd.DataFrame([{
        "Date": msg.get("date"),
        "Subject": msg.get("subject"),
        "Sender": msg.get("sender"),
        "Recipients": msg.get("recipients"),
        "EmailsInBody": msg.get("emails_in_body"),
        "PhonesInBody": msg.get("phones_in_body"),
        "Attachments": ";".join([att[0] for att in msg.get("attachments", [])]),
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
    title = Paragraph("Filtered Email Report", styles['Title'])
    elements.append(title)
    elements.append(Spacer(1, 12))
    table_data = [["Date", "Subject", "Sender", "Recipients", "Emails In Body", "Phones In Body", "Attachments"]]
    for msg in filtered_messages:
        date_str = msg.get("date").strftime("%Y-%m-%d %H:%M:%S") if msg.get("date") else ""
        row = [
            date_str,
            msg.get("subject", ""),
            msg.get("sender", ""),
            msg.get("recipients", ""),
            msg.get("emails_in_body", ""),
            msg.get("phones_in_body", ""),
            ";".join([att[0] for att in msg.get("attachments", [])]),
        ]
        table_data.append(row)
    tbl = Table(table_data, repeatRows=1)
    tbl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    elements.append(tbl)
    doc.build(elements)
    with open(buffer.name, 'rb') as f:
        data = f.read()
    os.unlink(buffer.name)
    return data

st.title("Email Forensic Tool - Streamlit Edition")
uploaded_file = st.file_uploader("Upload a .pst or .csv file", type=["pst", "csv"] )
if uploaded_file:
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    if ext == ".pst":
        messages, attachments_storage = parse_pst(uploaded_file)
    else:
        messages, attachments_storage = parse_csv(uploaded_file)
    df = pd.DataFrame([{
        "Date": msg.get("date"),
        "Subject": msg.get("subject"),
        "Sender": msg.get("sender"),
        "Recipients": msg.get("recipients"),
        "EmailsInBody": msg.get("emails_in_body"),
        "PhonesInBody": msg.get("phones_in_body"),
        "AttachmentsCount": len(msg.get("attachments", [])),
        "Index": i,
    } for i, msg in enumerate(messages)])
    st.sidebar.header("Filters")
    subj_filter = st.sidebar.text_input("Subject contains")
    sender_filter = st.sidebar.text_input("Sender contains")
    rec_filter = st.sidebar.text_input("Recipient contains")
    email_filter = st.sidebar.text_input("Email in body contains")
    phone_filter = st.sidebar.text_input("Phone in body contains")
    start_date = st.sidebar.date_input("Start date", value=datetime(2000,1,1).date())
    end_date = st.sidebar.date_input("End date", value=datetime.today().date())
    filtered = []
    for msg in messages:
        if subj_filter and subj_filter.lower() not in msg.get("subject", "").lower(): continue
        if sender_filter and sender_filter.lower() not in msg.get("sender", "").lower(): continue
        if rec_filter and rec_filter.lower() not in msg.get("recipients", "").lower(): continue
        if email_filter and email_filter.lower() not in msg.get("emails_in_body", "").lower(): continue
        if phone_filter and phone_filter.lower() not in msg.get("phones_in_body", "").lower(): continue
        msg_date = msg.get("date")
        if msg_date:
            if msg_date.date() < start_date or msg_date.date() > end_date: continue
        filtered.append(msg)
    if filtered:
        display_df = pd.DataFrame([{
            "Date": msg.get("date"),
            "Subject": msg.get("subject"),
            "Sender": msg.get("sender"),
            "Recipients": msg.get("recipients"),
            "EmailsInBody": msg.get("emails_in_body"),
            "PhonesInBody": msg.get("phones_in_body"),
            "AttachmentsCount": len(msg.get("attachments", [])),
            "Index": i,
        } for i, msg in enumerate(messages) if msg in filtered])
        display_df["Date"] = display_df["Date"].dt.strftime("%Y-%m-%d %H:%M:%S")
        st.dataframe(display_df.set_index("Index"))
        st.download_button("Download Filtered as CSV", data=generate_csv_download(filtered), file_name="filtered_emails.csv", mime="text/csv")
        st.download_button("Download Filtered as PDF", data=generate_pdf_download(filtered), file_name="filtered_emails.pdf", mime="application/pdf")
        st.write("## Message Details and Attachments")
        selected_index = st.number_input("Enter Index of message for details", min_value=min(display_df.index), max_value=max(display_df.index), step=1)
        selected_msgs = [msg for idx, msg in zip(display_df.index, filtered) if idx == selected_index]
        if selected_msgs:
            msg = selected_msgs[0]
            st.write(f"**Date:** {msg.get('date').strftime('%Y-%m-%d %H:%M:%S')}")
            st.write(f"**Subject:** {msg.get('subject')}" )
            st.write(f"**Sender:** {msg.get('sender')}" )
            st.write(f"**Recipients:** {msg.get('recipients')}" )
            if msg.get("attachments"):
                st.write("**Attachments:**")
                for filename, key in msg.get("attachments"):
                    data = attachments_storage.get(key)
                    if data:
                        st.download_button(f"Download {filename}", data=data, file_name=filename)
            st.write("**Body:**")
            st.write(msg.get("body"))
    else:
        st.info("No messages match the current filters.")
