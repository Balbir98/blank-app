import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import numbers, Font
from datetime import datetime
import zipfile
import os
import time

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

        # Check required columns exist
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

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for i, firm in enumerate(unique_firms):
                    firm_data = raw_df[raw_df["AR Firm Name"] == firm].sort_values(by="Adviser Name")
                    template_file.seek(0)  # Reset pointer to read template for each firm
                    wb = load_workbook(template_file)
                    ws = wb.active

                    # Set header data
                    ws["B2"] = firm  # Firm name in B2
                    ws["B3"] = firm_data["Date Paid to AR"].iloc[0].date() if pd.notnull(firm_data["Date Paid to AR"].iloc[0]) else ""  # Date in B3

                    # Start writing from row 7 (after header)
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
                        commission_cell.number_format = u"\u00a3#,##0.00"  # Format as GBP currency

                        # Match font of first row
                        sample_font = ws.cell(row=7, column=1).font
                        commission_cell.font = Font(name=sample_font.name, size=sample_font.size, bold=sample_font.bold)

                        start_row += 1

                    # Save output to in-memory buffer
                    output_buffer = io.BytesIO()
                    wb.save(output_buffer)
                    output_buffer.seek(0)

                    filename = f"Statement_{firm.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    zipf.writestr(filename, output_buffer.read())

                    elapsed = time.time() - start_time
                    progress_bar.progress((i + 1) / total_firms)
                    status_text.text(f"Processed {i + 1} of {total_firms} firms in {elapsed:.2f} seconds")

            zip_buffer.seek(0)
            st.download_button(
                label="Download All Statements as ZIP",
                data=zip_buffer,
                file_name="All_Commission_Statements.zip",
                mime="application/zip"
            )

            total_time = time.time() - start_time
            st.success(f"All statements generated successfully in {total_time:.2f} seconds!")
