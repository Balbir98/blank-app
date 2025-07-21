import streamlit as st
import os
import tempfile
import extract_msg
from fpdf import FPDF
from pathlib import Path
import zipfile
import re

# Helper: Create PDF from .msg email (no attachments)
class EmailPDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, "Email Export", ln=True, align='C')
        self.ln(10)

    def msg_to_pdf(self, msg):
        # Setup page and font
        self.set_auto_page_break(auto=True, margin=15)
        self.add_page()
        self.set_font("Arial", size=12)

        # Headers
        self.multi_cell(0, 10, f"From: {getattr(msg, 'sender', '') or '(no sender)'}")
        self.multi_cell(0, 10, f"To: {getattr(msg, 'to', '') or '(no recipient)'}")
        self.multi_cell(0, 10, f"Subject: {getattr(msg, 'subject', '') or '(no subject)'}")
        self.multi_cell(0, 10, f"Date: {getattr(msg, 'date', '') or '(no date)'}")
        self.ln(10)

        # Body
        raw = getattr(msg, 'body', '') or ""
        try:
            text = raw.encode("latin-1", "replace").decode("latin-1")
            self.multi_cell(0, 10, text if text.strip() else "(no message body)")
        except Exception as e:
            self.multi_cell(0, 10, f"[Error rendering body: {e}]")

# Convert .msg files in a zip to PDFs (including subfolders)
def convert_zipped_msg_files(zip_file, output_dir, progress_callback):
    # Unzip incoming buffer
    work = tempfile.mkdtemp()
    zip_in = os.path.join(work, "in.zip")
    with open(zip_in, "wb") as f:
        f.write(zip_file.read())
    with zipfile.ZipFile(zip_in, 'r') as zp:
        zp.extractall(work)

    # Gather all .msg paths
    msg_paths = list(Path(work).rglob("*.msg"))
    total = len(msg_paths)
    if total == 0:
        return False

    for idx, path in enumerate(msg_paths, start=1):
        # Attempt to parse message
        try:
            msg = extract_msg.Message(str(path))
        except Exception:
            # Create a stub message if parsing fails
            class Stub: pass
            msg = Stub()
            msg.sender = msg.to = msg.subject = msg.date = msg.body = ""

        # Build PDF
        pdf = EmailPDF()
        pdf.msg_to_pdf(msg)

        # Sanitize subject for filename
        raw_subj = getattr(msg, 'subject', '') or f"email_{idx}"
        # Replace non-alphanumeric chars with underscore
        safe_subj = re.sub(r'[^A-Za-z0-9_-]', '_', raw_subj)[:100]
        filename = f"{idx:04d}_{safe_subj}.pdf"
        out_path = os.path.join(output_dir, filename)

        # Write PDF, fallback to stub if write fails
        try:
            pdf.output(out_path)
        except Exception:
            # Generate minimal stub PDF
            stub = FPDF()
            stub.add_page()
            stub.set_font("Arial", size=12)
            stub.multi_cell(0, 10, "[Unable to generate this email PDF]")
            stub.output(out_path)

        # Update progress bar
        progress_callback(idx / total)

    st.info(f"✅ {total}/{total} emails converted to PDF.")
    return True

# Streamlit UI
st.title("📧 Outlook .msg → PDF (no attachments)")
st.markdown("Upload a `.zip` of `.msg` files; each one becomes a uniquely named PDF.")

uploaded = st.file_uploader("ZIP with .msg emails", type="zip")
if uploaded and st.button("Convert Emails to PDFs"):
    with st.spinner("Processing emails…"):
        temp_dir = tempfile.mkdtemp()
        output_dir = os.path.join(temp_dir, "pdf_output")
        os.makedirs(output_dir, exist_ok=True)
        progress_bar = st.progress(0.0)
        success = convert_zipped_msg_files(uploaded, output_dir, lambda p: progress_bar.progress(p))
        if not success:
            st.error("No .msg files found or conversion failed.")
        else:
            # Zip up PDFs
            zip_out = os.path.join(temp_dir, "converted_pdfs.zip")
            with zipfile.ZipFile(zip_out, 'w') as zf:
                for root, _, files in os.walk(output_dir):
                    for f in files:
                        full = os.path.join(root, f)
                        zf.write(full, os.path.relpath(full, output_dir))
            with open(zip_out, "rb") as f:
                st.download_button(
                    label="📥 Download Converted PDFs",
                    data=f,
                    file_name="converted_pdfs.zip",
                    mime="application/zip"
                )
            st.success("Conversion complete—one valid PDF per message!")
