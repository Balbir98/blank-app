import streamlit as st
import os
import tempfile
import extract_msg
from fpdf import FPDF
from pathlib import Path
import zipfile

# Helper: Create PDF from .msg email
class EmailPDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, "Email Export", ln=True, align='C')
        self.ln(10)

    def msg_to_pdf(self, msg, attachments_dir):
        self.set_auto_page_break(auto=True, margin=15)
        self.add_page()
        self.set_font("Arial", size=12)

        self.multi_cell(0, 10, f"From: {msg.sender}")
        self.multi_cell(0, 10, f"To: {msg.to}")
        self.multi_cell(0, 10, f"Subject: {msg.subject}")
        self.multi_cell(0, 10, f"Date: {msg.date}")
        self.ln(10)

        body = msg.body if msg.body else "(No message body)"
        self.multi_cell(0, 10, "Body:\n\n" + body)
        self.ln(10)

        for attachment in msg.attachments:
            attachment_path = os.path.join(attachments_dir, attachment.longFilename or attachment.shortFilename)
            with open(attachment_path, 'wb') as f:
                f.write(attachment.data)
            self.multi_cell(0, 10, f"[Attachment saved: {attachment.longFilename or attachment.shortFilename}]")

# Convert individual .msg files to PDFs
def convert_uploaded_msg_files(files, output_dir, progress_callback):
    total = len(files)
    if total == 0:
        return False

    for i, uploaded_file in enumerate(files):
        try:
            temp_msg_path = os.path.join(output_dir, f"email_{i}.msg")
            with open(temp_msg_path, "wb") as f:
                f.write(uploaded_file.read())

            msg = extract_msg.Message(temp_msg_path)
            msg.extract_attachments()

            attachments_dir = os.path.join(output_dir, f"attachments_{i}")
            os.makedirs(attachments_dir, exist_ok=True)

            pdf = EmailPDF()
            pdf.msg_to_pdf(msg, attachments_dir)

            pdf_path = os.path.join(output_dir, f"email_{i+1}.pdf")
            pdf.output(pdf_path)

            progress_callback((i + 1) / total)
        except Exception as e:
            print(f"Error processing file {uploaded_file.name}: {e}")

    return True

# Main App
st.title("📧 Outlook Email (.msg) to PDF Converter")
st.markdown("Upload one or more `.msg` email files. The app will convert each email into a PDF including attachments.")

uploaded_files = st.file_uploader("Upload .msg email files", type="msg", accept_multiple_files=True)

if uploaded_files:
    if st.button("Convert Emails to PDFs"):
        with st.spinner("Processing emails..."):
            temp_dir = tempfile.mkdtemp()
            output_dir = os.path.join(temp_dir, "pdf_output")
            os.makedirs(output_dir, exist_ok=True)

            progress_bar = st.progress(0)
            success = convert_uploaded_msg_files(uploaded_files, output_dir, lambda p: progress_bar.progress(p))

            if not success:
                st.error("No .msg files found or failed to convert emails.")
            else:
                # Zip PDFs
                zip_output_path = os.path.join(temp_dir, "converted_pdfs.zip")
                with zipfile.ZipFile(zip_output_path, 'w') as zipf:
                    for root, _, files in os.walk(output_dir):
                        for file in files:
                            full_path = os.path.join(root, file)
                            arcname = os.path.relpath(full_path, output_dir)
                            zipf.write(full_path, arcname=arcname)

                with open(zip_output_path, "rb") as f:
                    st.download_button(
                        label="📥 Download Converted PDFs",
                        data=f,
                        file_name="converted_pdfs.zip",
                        mime="application/zip"
                    )

                st.success("Conversion complete! PDFs and attachments saved.")
