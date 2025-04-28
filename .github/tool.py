import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
import re

st.title("📄 Commission PDF Extractor")
st.markdown("""
Upload a commission statement PDF and select the provider.  
Supported:  
- ✅ Canada Life  
- ✅ MetLife  
- ✅ Aviva Healthcare  
- ✅ CETA  
- ✅ Accord BTL  
- ✅ Medicash
- ✅ Cigna (Once exported, enter in the Firm Name and Broker Reference to the spreadsheet.)
- ✅ DenPlan
- ✅ National Friendly
         
 """)

uploaded_file = st.file_uploader("Upload your PDF", type=["pdf"])
provider = st.selectbox("Select the Provider", [
    "Choose...", "Canada Life", "MetLife", "Aviva Healthcare", "CETA", "Accord BTL", "Medicash", "Cigna"
, "DenPlan", "National Friendly" ])

if st.button("RUN"):
    if uploaded_file is None or provider == "Choose...":
        st.error("Please upload a PDF and select a provider.")
    else:
        try:
            all_rows = []

            with pdfplumber.open(uploaded_file) as pdf:
                if provider == "Canada Life":
                    current_intermediary = None
                    for page in pdf.pages:
                        text = page.extract_text()
                        if not text:
                            continue
                        lines = text.split('\n')
                        for line in lines:
                            line = line.strip()
                            if line.startswith("Intermediary:"):
                                current_intermediary = line.split("Intermediary:")[1].strip()
                            elif ("Policy Name" in line or "Intermediary Total" in line or
                                  "statement of commission" in line.lower()):
                                continue
                            elif (current_intermediary and
                                  len(re.findall(r"\d{2}/\d{2}/\d{4}", line)) >= 2 and
                                  len(re.findall(r"£", line)) >= 2):
                                parts = line.split()
                                policy_idx = next((i for i, part in enumerate(parts) if '/' in part), None)
                                if policy_idx and policy_idx >= 1:
                                    company_name = " ".join(parts[:policy_idx])
                                    policy_code = parts[policy_idx]
                                    remainder = parts[policy_idx + 1:]
                                    if len(remainder) >= 6:
                                        row = [current_intermediary, company_name, policy_code] + remainder[:6]
                                        all_rows.append(row)
                    columns = [
                        "Intermediary", "Policy Name", "Policy Code", "Policy Type",
                        "Received Date", "Due Date", "Payment Received",
                        "Commission Percentage", "Intermediary Commission"
                    ]

                elif provider == "MetLife":
                    current_broker_id = ""
                    current_firm_name = ""
                    for page in pdf.pages:
                        text = page.extract_text()
                        if not text:
                            continue
                        lines = text.split('\n')
                        for i, line in enumerate(lines):
                            line = line.strip()
                            if line.startswith("Broker -"):
                                parts = line.split(" - ")
                                if len(parts) >= 3:
                                    current_broker_id = parts[1].strip()
                                    current_firm_name = parts[2].strip()
                            elif line.startswith("Policy Number") and "Scheme Name" in line:
                                continue
                            elif (
                                "£" in line and "%" in line and
                                re.search(r"\d{1,2} \w+ \d{4}", line)
                            ):
                                row_parts = line.split()
                                if len(row_parts) >= 7:
                                    policy_number = row_parts[0]
                                    date_match = re.search(r"\d{1,2} \w+ \d{4}", line)
                                    if not date_match:
                                        continue
                                    date_str = date_match.group()
                                    try:
                                        scheme_name = line.split(policy_number)[1].split(date_str)[0].strip()
                                    except:
                                        scheme_name = ""
                                    amount_match = re.search(r"£[\d,]+\.\d{2}", line)
                                    rate_match = re.search(r"\d{1,2}\.\d{2}%", line)
                                    commission_due_match = re.findall(r"£[\d,]+\.\d{2}", line)
                                    amount_received = amount_match.group() if amount_match else ""
                                    commission_rate = rate_match.group() if rate_match else ""
                                    commission_due = commission_due_match[-1] if commission_due_match else ""
                                    all_rows.append([
                                        current_broker_id,
                                        current_firm_name,
                                        policy_number,
                                        scheme_name,
                                        date_str,
                                        amount_received,
                                        commission_rate,
                                        commission_due
                                    ])
                    columns = [
                        "Broker ID", "Firm Name", "Policy Number", "Scheme Name",
                        "Date Received", "Amount Received", "Commission Rate", "Commission Due"
                    ]

                elif provider == "Aviva Healthcare":
                    for i, page in enumerate(pdf.pages):
                        table = page.extract_table()
                        if table:
                            data_rows = table[1:] if i == 0 else table
                            all_rows.extend(data_rows)
                    columns = [
                        "Line of Business", "IPT", "Policy No", "Policy Name", "Billing Date",
                        "FRQ", "Premium", "Type", "Comm %", "Comm Paid", "Agency Code"
                    ]

                elif provider == "CETA":
                    for page in pdf.pages:
                        table = page.extract_table()
                        if table:
                            data_rows = table[1:]
                            all_rows.extend(data_rows)
                    columns = [
                        "Master No", "Policy ID", "Client Name", "Type",
                        "Date", "Reason", "Insurer", "Premium", "Code", "Commission"
                    ]

                elif provider == "Accord BTL":
                    for page in pdf.pages:
                        table = page.extract_table()
                        if table:
                            data_rows = table[1:]
                            all_rows.extend(data_rows)
                    columns = [
                        "Broker Name", "Company Name", "FCA Reference", "Customer Surname",
                        "Reference", "Product Transfer Date", "Account Balance", "Proc Fee Amount"
                    ]

                elif provider == "Medicash":
                    first_page_text = pdf.pages[0].extract_text()
                    firm_name = "Unknown Firm"
                    if first_page_text:
                        lines = [line.strip() for line in first_page_text.split('\n') if line.strip()]
                        for line in lines:
                            if (
                                line.lower() not in ["commission statement", "date:", "month ending:"]
                                and not re.match(r"date[:]?|month ending[:]?|commission statement", line, re.IGNORECASE)
                                and len(line) > 2
                            ):
                                firm_name = line
                                break
                    for page in pdf.pages:
                        table = page.extract_table()
                        if table and "Policy/Group number" in table[0][0]:
                            data_rows = table[1:]
                            for row in data_rows:
                                all_rows.append([firm_name] + row)
                            break
                    columns = [
                        "Firm Name",
                        "Policy/Group number", "Policyholder/Group Name",
                        "Premium rec'd", "IPT deduction", "Rate", "Commission due"
                    ]
                elif provider == "Cigna":
                    firm_name = ""
                    broker_ref = ""
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text:
                            for line in text.split('\n'):
                                if "Account Name:" in line:
                                    firm_name = line.split("Account Name:")[1].strip()
                                if "Broker Reference Number:" in line:
                                    broker_ref = line.split("Broker Reference Number:")[1].strip()
                        table = page.extract_table()
                        if table:
                            headers = table[0]
                            for row in table[1:]:
                                # Skip rows that are not table data
                                if not row or len(row) < 8 or "Total Commission Due" in row[0]:
                                    continue
                                all_rows.append([
                                    firm_name,
                                    broker_ref,
                                    row[0],  # Policy Number
                                    row[1],  # Policyholder
                                    row[2],  # Transaction No
                                    row[3],  # Premium Due Date
                                    row[4],  # Premium Paid Date
                                    row[5],  # Premium Paid (Inc Tax)
                                    row[6],  # Commission Percent
                                    row[7],  # Commission Amount
                                ])
                    columns = [
                        "Firm Name", "Broker Reference Number",
                        "Policy Number", "Policyholder", "Transaction No",
                        "Premium Due Date", "Premium Paid Date",
                        "Premium Paid (Inc Tax)", "Commission Percent", "Commission Amount"
                    ]
                elif provider == "DenPlan":
                    broker_ref = ""
                    broker_name = ""

                    # Extract Broker Ref and Broker Name from header
                    first_page_text = pdf.pages[0].extract_text()
                    if first_page_text:
                        for line in first_page_text.split("\n"):
                            if "Broker Ref:" in line:
                                broker_ref = line.split("Broker Ref:")[1].strip().split()[0]  # just number
                            if "Broker Name:" in line:
                                broker_name = line.split("Broker Name:")[1].strip()

                    for page in pdf.pages:
                        text = page.extract_text()
                        if not text:
                            continue
                        lines = text.split("\n")
                        start = False
                        for line in lines:
                            if line.strip().startswith("Group Ref and Name"):
                                start = True
                                continue
                            if not start or not line.strip():
                                continue
                            if "Total Paid" in line or line.strip() == "":
                                break
                            if not line.strip().startswith("GR"):
                                continue

                            parts = line.strip().split()
                            if len(parts) < 8:
                                continue  # not enough data

                            # **ALWAYS take the last 5 columns as rightmost fields**
                            # (commission_amount, commission_rate, payment_received, start_date, type_val)
                            commission_amount = parts[-1]
                            commission_rate = parts[-2]
                            payment_received = parts[-3]
                            start_date = parts[-4]
                            type_val = parts[-5]
                            # Everything before that = group ref and name
                            group_ref_and_name = " ".join(parts[:-5])

                            all_rows.append([
                                broker_ref,
                                broker_name,
                                group_ref_and_name,
                                type_val,
                                start_date,
                                payment_received,
                                commission_rate,
                                commission_amount
                            ])

                    columns = [
                        "Broker Ref", "Broker Name", "Group Ref and Name",
                        "Type", "Start Date", "Payment Received",
                        "Commission Rate", "Commission Amount"
                    ]
                elif provider == "National Friendly":
                    # Define your columns
                    columns = [
                        "Company Name", "Product", "Plan No", "Member No", "Member Name",
                        "FC No", "Agent", "UW", "Annual Premium", "Commission Rate",
                        "Commission £", "1st Premium Received", "Issue Date", "Clawback Period (Mths)"
                    ]
                    all_rows = []
                    for page in pdf.pages:
                        table = page.extract_table()
                        if table:
                            header_row = None
                            for idx, row in enumerate(table):
                                # Find the header row by checking for column names
                                if row and "Company Name" in row[0]:
                                    header_row = idx
                                    break
                            if header_row is not None:
                                for row in table[header_row+1:]:
                                    # Skip blank or summary lines
                                    if not row or not row[0] or "Total Payable" in row[0]:
                                        continue
                                    # Fix common issues with Annual Premium/Commission Rate
                                    # Sometimes the cells get merged or split, so clean up
                                    if len(row) > len(columns):
                                        # Try to merge back split columns
                                        row = row[:8] + [' '.join(row[8:-5])] + row[-5:]
                                    if len(row) < len(columns):
                                        # If still short, pad
                                        row += [''] * (len(columns) - len(row))
                                    # Remove stray letters like "ng" from Commission Rate
                                    row[9] = re.sub(r'[^\d.%]', '', row[9])
                                    # Remove any extra characters from Annual Premium
                                    row[8] = re.sub(r'[^\d.,£]', '', row[8])
                                    all_rows.append(row)
                                break 
                

                                
            if not all_rows:
                st.warning("⚠️ No data rows were found.")
            else:
                df = pd.DataFrame(all_rows, columns=columns)
                st.success(f"✅ {provider} commission data extracted successfully!")
                st.dataframe(df)

                output = BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    df.to_excel(writer, index=False, sheet_name=f"{provider} Data")
                output.seek(0)

                st.download_button(
                    label="📥 Download Excel File",
                    data=output,
                    file_name=f"{provider.replace(' ', '_')}_commission_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"❌ Error while processing: {e}")
