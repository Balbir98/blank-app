import streamlit as st
import zipfile
import os
import tempfile
import shutil
import base64
from email import policy
from email.parser import BytesParser
from fpdf import FPDF
from pathlib import Path

# Helper: Create PDF from email
class EmailPDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, "Email Export", ln=True, align='C')
        self.ln(10)

    def email_to_pdf(self, msg, attachments_dir):
        self.set_auto_page_break(auto=True, margin=15)
        self.add_page()

        self.set_font("Arial", size=12)
        self.multi_cell(0, 10, f"From: {msg.get('From', '')}")
        self.multi_cell(0, 10, f"To: {msg.get('To', '')}")
        self.multi_cell(0, 10, f"Subject: {msg.get('Subject', '')}")
        self.multi_cell(0, 10, f"Date: {msg.get('Date', '')}")
        self.ln(10)

        # Email Body
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                if ctype == 'text/plain' and part.get_content_disposition() != 'attachment':
                    body = part.get_payload(decode=True).decode(errors="ignore")
                    break
        else:
            body = msg.get_payload(decode=True).decode(errors="ignore")

        self.multi_cell(0, 10, "Body:\n\n" + body)
        self.ln(10)

        # Attachments
        for part in msg.walk():
            if part.get_content_disposition() == 'attachment':
                filename = part.get_filename()
                if filename:
                    attachment_path = os.path.join(attachments_dir, filename)
                    with open(attachment_path, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    self.multi_cell(0, 10, f"[Attachment saved: {filename}]")

# Helper: Convert all emails in a folder to PDFs
def convert_emails_to_pdfs(email_dir, output_dir, progress_callback):
    files = list(Path(email_dir).glob("**/*.eml"))
    total = len(files)

    for i, file_path in enumerate(files):
        with open(file_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)

        attachments_dir = os.path.join(output_dir, f"attachments_{i}")
        os.makedirs(attachments_dir, exist_ok=True)

        pdf = EmailPDF()
        pdf.email_to_pdf(msg, attachments_dir)

        pdf_path = os.path.join(output_dir, f"email_{i+1}.pdf")
        pdf.output(pdf_path)

        progress_callback((i + 1) / total)

# Main App
st.title("📧 Email to PDF Converter")
st.markdown("Upload a `.zip` file containing `.eml` email files. The app will convert each email into a PDF including attachments.")

uploaded_file = st.file_uploader("Upload ZIP file with emails", type="zip")

if uploaded_file:
    if st.button("Convert Emails to PDFs"):
        with st.spinner("Processing emails..."):
            temp_dir = tempfile.mkdtemp()
            zip_path = os.path.join(temp_dir, "uploaded.zip")

            with open(zip_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            output_dir = os.path.join(temp_dir, "pdf_output")
            os.makedirs(output_dir, exist_ok=True)

            progress_bar = st.progress(0)
            convert_emails_to_pdfs(temp_dir, output_dir, lambda p: progress_bar.progress(p))

            # Zip PDFs
            zip_output_path = os.path.join(temp_dir, "converted_pdfs.zip")
            with zipfile.ZipFile(zip_output_path, 'w') as zipf:
                for root, _, files in os.walk(output_dir):
                    for file in files:
                        full_path = os.path.join(root, file)
                        arcname = os.path.relpath(full_path, output_dir)
                        zipf.write(full_path, arcname=arcname)

            # Download link
            with open(zip_output_path, "rb") as f:
                zip_bytes = f.read()
                b64 = base64.b64encode(zip_bytes).decode()
                href = f'<a href="data:application/zip;base64,{b64}" download="converted_pdfs.zip">Download Converted PDFs</a>'
                st.markdown(href, unsafe_allow_html=True)

        st.success("Conversion complete! PDFs and attachments saved.")
