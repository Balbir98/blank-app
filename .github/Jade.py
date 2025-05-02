import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from datetime import datetime

st.title("📄 Bank Statement PDF Extractor")
st.markdown(""" Upload a Bank Statement and Select the Provider """)
uploaded_file = st.file_uploader("Upload a bank statement PDF", type=["pdf"])
bank_account = st.selectbox("Select Bank Account", ["-- Select --", "Transfer Wise"])

if uploaded_file and bank_account == "Transfer Wise":
    with pdfplumber.open(uploaded_file) as pdf:
        records = []

        for page in pdf.pages:
            text = page.extract_text()
            lines = text.split('\n')

            for i, line in enumerate(lines):
                # Extract payee
                payee = None
                if "Sent money to" in line:
                    payee = line.split("Sent money to")[-1].strip()
                elif "Received money from" in line:
                    payee = line.split("Received money from")[-1].strip()
                elif "GBP Assets service fee" in line:
                    payee = "GBP Assets service fee"

                # Clean payee
                if payee and payee != "GBP Assets service fee":
                    payee = re.sub(r"( with reference.*| [\-]?[\d,]+\.\d{2}.*$)", "", payee).strip()

                # Look ahead for date and amounts if payee was found
                if payee:
                    date = None
                    amount = None
                    incoming = None
                    outgoing = None

                    # Search for date in next 1-3 lines before 'Transaction:'
                    for j in range(i + 1, min(i + 4, len(lines))):
                        if "Transaction:" in lines[j]:
                            date_match = re.search(r"(\d{1,2} \w+ \d{4})", lines[j])
                            if date_match:
                                date_str = date_match.group(1)
                                try:
                                    date = datetime.strptime(date_str, "%d %B %Y").date()
                                except:
                                    pass

                    # Search for amount in lines near the payee block
                    for k in range(i, min(i + 6, len(lines))):
                        # Grab all monetary values on line
                        amount_match = re.findall(r"([\-]?[\d,]+\.\d{2})", lines[k])
                        if len(amount_match) >= 2:
                            in_val = amount_match[0].replace(",", "")
                            out_val = amount_match[1].replace(",", "")
                            if in_val != "0.00":
                                amount = float(in_val)
                            else:
                                amount = -float(out_val)
                            break
                        elif len(amount_match) == 1 and "GBP Assets service fee" in payee:
                            # Assume fee is negative
                            amount = -float(amount_match[0].replace(",", ""))
                            break

                    if date and amount is not None:
                        records.append({
                            "Date": date,
                            "Amount": amount,
                            "Payee": payee
                        })

        df = pd.DataFrame(records)

        if not df.empty:
            st.subheader("Preview of Extracted Data")
            st.dataframe(df.head(10))  # Show only top 10 rows

            # Download button
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Transactions')
            st.download_button(
                label="📥 Download Excel File",
                data=output.getvalue(),
                file_name="bank_account_1_transactions.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No transaction data found. Check formatting or try another document.")
