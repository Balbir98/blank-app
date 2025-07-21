import streamlit as st
import os
import tempfile
import extract_msg
from fpdf import FPDF
from pathlib import Path
import zipfile

# Helper: Create PDF from .msg email (no attachments)
class EmailPDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, "Email Export", ln=True, align='C')
        self.ln(10)

    def msg_to_pdf(self, msg):
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
    work = tempfile.mkdtemp()
    zpath = os.path.join(work, "in.zip")
    with open(zpath, "wb") as f:
        f.write(zip_file.read())
    with zipfile.ZipFile(zpath, 'r') as zp:
        zp.extractall(work)

    paths = list(Path(work).rglob("*.msg"))
    total = len(paths)
    if total == 0:
        return False

    for idx, p in enumerate(paths, start=1):
        try:
            msg = extract_msg.Message(str(p))
        except Exception:
            # stub for unparsable
            class Stub: pass
            msg = Stub()
            msg.sender = msg.to = msg.subject = msg.date = msg.body = ""

        # Build and save PDF
        pdf = EmailPDF()
        pdf.msg_to_pdf(msg)

        # Unique, index-prefixed filename
        subj = getattr(msg, 'subject', '') or f"email_{idx}"
        safe = "_".join(subj.split())[:100].replace("/", "_").replace("\\", "_")
        filename = f"{idx:04d}_{safe}.pdf"
        out_path = os.path.join(output_dir, filename)

        try:
            pdf.output(out_path)
        except Exception as e:
            st.warning(f"Couldn’t write PDF #{idx}: {e}")

        progress_callback(idx / total)

    st.info(f"✅ {total}/{total} emails converted to PDF.")
    return True

# Streamlit UI
st.title("📧 Outlook .msg → PDF (unique filenames)")
st.markdown("Upload a `.zip` of `.msg` files; each one becomes a PDF named `####_subject.pdf`.")

up = st.file_uploader("ZIP with .msg emails", type="zip")
if up and st.button("Convert"):
    with st.spinner("Working…"):
        td = tempfile.mkdtemp()
        od = os.path.join(td, "pdfs")
        os.makedirs(od, exist_ok=True)
        bar = st.progress(0.0)
        ok = convert_zipped_msg_files(up, od, lambda p: bar.progress(p))
        if not ok:
            st.error("No .msg files found.")
        else:
            # Zip and provide download
            zip_out = os.path.join(td, "out.zip")
            with zipfile.ZipFile(zip_out, 'w') as zf:
                for r, _, files in os.walk(od):
                    for fn in files:
                        full = os.path.join(r, fn)
                        zf.write(full, os.path.relpath(full, od))
            with open(zip_out, "rb") as f:
                st.download_button("📥 Download PDFs", data=f, file_name="converted_pdfs.zip", mime="application/zip")
            st.success("Done—one unique PDF per .msg!")
