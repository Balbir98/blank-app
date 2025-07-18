import streamlit as st
import os
import tempfile
import extract_msg
from fpdf import FPDF
from pathlib import Path
import zipfile
import shutil

# Helper: Create PDF from .msg email (with inline images)
class EmailPDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, "Email Export", ln=True, align='C')
        self.ln(10)

    def msg_to_pdf(self, msg):
        # Render headers & body
        self.set_auto_page_break(auto=True, margin=15)
        self.add_page()
        self.set_font("Arial", size=12)
        self.multi_cell(0, 10, f"From: {msg.sender or '(no sender)'}")
        self.multi_cell(0, 10, f"To: {msg.to or '(no recipient)'}")
        self.multi_cell(0, 10, f"Subject: {msg.subject or '(no subject)'}")
        self.multi_cell(0, 10, f"Date: {msg.date or '(no date)'}")
        self.ln(10)

        body = msg.body or ""
        try:
            rendered = body.encode("latin-1", "replace").decode("latin-1")
            self.multi_cell(0, 10, rendered if rendered.strip() else "(no message body)")
        except Exception as e:
            self.multi_cell(0, 10, f"[Error rendering body: {e}]")
        self.ln(10)

    def embed_attachments(self, msg, attachments_dir):
        # Save all attachments, embed images
        for i, att in enumerate(msg.attachments):
            fname = att.longFilename or att.shortFilename or f"attachment_{i}"
            safe_fname = fname.replace("/", "_").replace("\\", "_")
            path = os.path.join(attachments_dir, safe_fname)
            try:
                with open(path, "wb") as f:
                    f.write(att.data)
            except Exception:
                continue

            # If it's an image, put it in the PDF
            if safe_fname.lower().endswith((".png", ".jpg", ".jpeg", ".gif")):
                try:
                    self.add_page()
                    self.set_font("Arial", "B", 12)
                    self.multi_cell(0, 10, f"Attachment: {safe_fname}")
                    # Fit image to page width minus margins
                    max_w = self.w - 2*self.l_margin
                    self.image(path, w=max_w)
                    self.ln(10)
                except Exception:
                    continue

# Convert .msg files in a zip to PDFs (including subfolders)
def convert_zipped_msg_files(zip_file, output_dir, progress_callback):
    temp_dir = tempfile.mkdtemp()
    zip_path = os.path.join(temp_dir, "emails.zip")
    with open(zip_path, "wb") as f:
        f.write(zip_file.read())
    with zipfile.ZipFile(zip_path, 'r') as zp:
        zp.extractall(temp_dir)

    msg_paths = list(Path(temp_dir).rglob("*.msg"))
    total = len(msg_paths)
    if total == 0:
        return False

    converted = 0
    for idx, msg_path in enumerate(msg_paths, start=1):
        # Load message
        try:
            msg = extract_msg.Message(str(msg_path))
        except Exception as e:
            # if we can’t even parse, make a stub PDF
            msg = type("stub", (), {})()
            msg.sender = msg.subject = msg.body = msg.to = msg.date = ""
            msg.attachments = []

        # Prepare PDF
        pdf = EmailPDF()
        pdf.msg_to_pdf(msg)

        # Save attachments & embed images
        attachments_dir = os.path.join(output_dir, f"attachments_{idx}")
        os.makedirs(attachments_dir, exist_ok=True)
        try:
            pdf.embed_attachments(msg, attachments_dir)
        except Exception:
            pass

        # Write out PDF regardless of errors
        safe_subj = (msg.subject or f"email_{idx}").strip()
        safe_subj = "_".join(safe_subj.split())[:100].replace("/", "_").replace("\\", "_")
        pdf_filename = os.path.join(output_dir, f"{safe_subj}.pdf")
        try:
            pdf.output(pdf_filename)
        except Exception as e:
            st.warning(f"Failed to save PDF for {msg_path.name}: {e}")

        converted += 1
        progress_callback(converted / total)

    st.info(f"✅ {converted}/{total} emails converted to PDF.")
    return True

# Streamlit UI
st.title("📧 Outlook Email (.msg) ZIP → PDF Converter")
st.markdown("Upload a `.zip` of `.msg` files; each will become a PDF with inline images embedded.")

uploaded = st.file_uploader("Upload ZIP with .msg emails", type="zip")
if uploaded and st.button("Convert Emails to PDFs"):
    with st.spinner("Processing…"):
        tmp = tempfile.mkdtemp()
        out_dir = os.path.join(tmp, "pdf_output")
        os.makedirs(out_dir, exist_ok=True)
        bar = st.progress(0.0)
        success = convert_zipped_msg_files(uploaded, out_dir, lambda p: bar.progress(p))
        if not success:
            st.error("No .msg files found or conversion failed.")
        else:
            # Zip up results
            out_zip = os.path.join(tmp, "converted_pdfs.zip")
            with zipfile.ZipFile(out_zip, 'w') as zf:
                for root, _, files in os.walk(out_dir):
                    for f in files:
                        full = os.path.join(root, f)
                        arc = os.path.relpath(full, out_dir)
                        zf.write(full, arc)
            with open(out_zip, "rb") as f:
                st.download_button("📥 Download Converted PDFs", data=f, file_name="converted_pdfs.zip", mime="application/zip")
            st.success("All done! 🎉 PDFs (with images) are ready.")
