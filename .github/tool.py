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
- ✅ INET  
""")

uploaded_file = st.file_uploader("Upload your PDF", type=["pdf"])
provider = st.selectbox("Select the Provider", ["Choose...", "Canada Life", "MetLife", "Aviva Healthcare", "INET"])

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

                elif provider == "INET":
                    agent_id = ""
                    firm_name = ""
                    rows = []

                    for page in pdf.pages:
                        text = page.extract_text()
                        if text:
                            for line in text.split('\n'):
                                if line.strip().startswith("Agent"):
                                    agent_match = re.search(r"\((.*?)\)", line)
                                    if agent_match:
                                        agent_id = agent_match.group(1).strip()
                                        post_paren = line.split(")", 1)[1] if ")" in line else ""
                                        firm_name_segment = post_paren.split("Account Name")[0].strip()
                                        firm_name_cleaned = re.sub(r"(?<=\d)(?=[A-Za-z])|(?<=[a-z])(?=[A-Z])", " ", firm_name_segment)
                                        firm_name = " ".join(firm_name_cleaned.split())

                        table = page.extract_table()
                        if not table:
                            continue
                        data_rows = table[1:]

                        for row in data_rows:
                            row = [col if col else "" for col in row]

                            # 🔐 Fix: use row[0] if the row is short
                            combined = row[2] if len(row) > 2 else row[0]

                            # Pattern to extract merged content
                            pattern = (
                                r"(?P<first>[A-Z][a-z]+)"
                                r"(?P<last>[A-Z][a-z]+)?"
                                r"(IMD\d+)"
                                r"(?P<uw>[A-Z]{2,3})?"       # Optional UW
                                r"(?P<status>[A-Z]{2,4})"    # Status is always present
                                r"(\d{2}[A-Za-z]{3}\d{4})"
                                r"(\d{1,2}\.\d+%)"
                                r"£([\d,]+\.\d{2})"
                                r"£([\d,]+\.\d{2})"
                            )

                            match = re.search(pattern, combined.replace(" ", ""))
                            if match:
                                client = f"{match.group('first')} {match.group('last') or ''}".strip()
                                policy_number = match.group(3)
                                uw = match.group('uw') or ""
                                status = match.group('status')
                                start = match.group(6)
                                rate = match.group(7)
                                premium = f"£{match.group(8)}"
                                commission = f"£{match.group(9)}"

                                rows.append([
                                    agent_id,
                                    firm_name,
                                    client,
                                    policy_number,
                                    uw,
                                    status,
                                    start,
                                    rate,
                                    premium,
                                    commission
                                ])
                            else:
                                fallback = [agent_id, firm_name] + row + [""] * (10 - len(row))
                                rows.append(fallback[:10])

                    all_rows.extend(rows)
                    columns = [
                        "Agent ID", "Firm Name", "Client", "Policy Number", "UW",
                        "Status", "Start", "Rate", "Premium Exc. IPT", "Commission"
                    ]


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

