import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import numbers, Font
from datetime import datetime
import zipfile
import os
import time
import tempfile
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

st.title("Commission Statement Generator")

st.markdown("""
Upload the **raw commission data** and the **Excel template** to begin.
This tool will generate one statement per AR Firm.
""")

raw_data_file = st.file_uploader("Upload Raw Commission Data (Excel)", type=["xlsx"], key="raw")
template_file = st.file_uploader("Upload Commission Statement Template (Excel)", type=["xlsx"], key="template")

if raw_data_file and template_file:
    if st.button("Run - Generate Statements"):
        start_time = time.time()
        raw_df = pd.read_excel(raw_data_file)

        expected_columns = [
            "Principal/Adviser Email Address", "AR Firm Name", "Adviser Name", "Date of Statement",
            "Lender", "Policy Reference", "Product Type", "Client First Name", "Client Surname",
            "Class", "Commission Payable", "Date Paid to AR"
        ]

        if not all(col in raw_df.columns for col in expected_columns):
            st.error("The raw data is missing one or more required columns.")
        else:
            unique_firms = raw_df["AR Firm Name"].dropna().unique()
            total_firms = len(unique_firms)
            progress_bar = st.progress(0)
            status_text = st.empty()
            zip_buffer = io.BytesIO()
            eml_zip_buffer = io.BytesIO()

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf, \
                 zipfile.ZipFile(eml_zip_buffer, "w", zipfile.ZIP_DEFLATED) as eml_zip:
                for i, firm in enumerate(unique_firms):
                    firm_data = raw_df[raw_df["AR Firm Name"] == firm].sort_values(by="Adviser Name")
                    template_file.seek(0)
                    wb = load_workbook(template_file)
                    ws = wb.active
                    ws["B2"] = firm
                    ws["B3"] = firm_data["Date Paid to AR"].iloc[0].date() if pd.notnull(firm_data["Date Paid to AR"].iloc[0]) else ""
                    start_row = 7
                    for idx, row in firm_data.iterrows():
                        ws.cell(row=start_row, column=1, value=row["Adviser Name"])
                        ws.cell(row=start_row, column=2, value=row["Date of Statement"].date() if pd.notnull(row["Date of Statement"]) else "")
                        ws.cell(row=start_row, column=3, value=row["Lender"])
                        ws.cell(row=start_row, column=4, value=row["Policy Reference"])
                        ws.cell(row=start_row, column=5, value=row["Product Type"])
                        ws.cell(row=start_row, column=6, value=row["Client First Name"])
                        ws.cell(row=start_row, column=7, value=row["Client Surname"])
                        ws.cell(row=start_row, column=8, value=row["Class"])
                        commission_cell = ws.cell(row=start_row, column=9, value=row["Commission Payable"])
                        commission_cell.number_format = u"\u00a3#,##0.00"
                        sample_font = ws.cell(row=7, column=1).font
                        commission_cell.font = Font(name=sample_font.name, size=sample_font.size, bold=sample_font.bold)
                        start_row += 1
                    output_buffer = io.BytesIO()
                    wb.save(output_buffer)
                    output_buffer.seek(0)
                    filename = f"Statement_{firm.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    zipf.writestr(filename, output_buffer.getvalue())

                    recipient = firm_data["Principal/Adviser Email Address"].iloc[0]
                    subject = f"Commission Statement - {firm}"
                    html_body = f"""
                        <html>
                            <body>
                                <p>Dear Adviser,</p>
                                <p>Please find attached the latest commission statement for your firm: <strong>{firm}</strong>.</p>
                                <p>If you have any questions, feel free to get in touch.</p>
                                <p>Best regards!</p>
                            </body>
                        </html>
                    """

                    msg = MIMEMultipart("mixed")
                    msg["To"] = recipient
                    msg["Subject"] = subject
                    msg.add_header("X-Unsent", "1")

                    alt_part = MIMEMultipart("alternative")
                    alt_part.attach(MIMEText(html_body, "html"))
                    msg.attach(alt_part)

                    part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    part.set_payload(output_buffer.getvalue())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename=\"{filename}\"")
                    msg.attach(part)

                    eml_filename = f"Email_{firm.replace(' ', '_')}.eml"
                    eml_io = io.BytesIO()
                    from email.generator import BytesGenerator
                    gen = BytesGenerator(eml_io)
                    gen.flatten(msg)
                    eml_io.seek(0)
                    eml_zip.writestr(eml_filename, eml_io.read())

                    elapsed = time.time() - start_time
                    progress_bar.progress((i + 1) / total_firms)
                    status_text.text(f"Processed {i + 1} of {total_firms} firms in {elapsed:.2f} seconds")

            zip_buffer.seek(0)
            eml_zip_buffer.seek(0)
            st.download_button(
                label="Download All Statements as ZIP",
                data=zip_buffer,
                file_name="All_Commission_Statements.zip",
                mime="application/zip"
            )
            st.download_button(
                label="Download Draft Emails as EML ZIP",
                data=eml_zip_buffer,
                file_name="All_Email_Drafts.zip",
                mime="application/zip"
            )
            total_time = time.time() - start_time
            st.success(f"All statements and email drafts generated successfully in {total_time:.2f} seconds!")
