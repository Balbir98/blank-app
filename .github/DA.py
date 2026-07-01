import io
import re
import zipfile
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

APP_VERSION = "v2.1 - Steve/Steven recruiter alias"


# -----------------------------
# Helpers
# -----------------------------

SUPPORT_REQUIRED_HEADERS = [
    "Date",
    "Source",
    "Description",
    "Reference",
    "Debit (GBP)",
    "Credit (GBP)",
    "Gross (GBP)",
    "VAT (GBP)",
    "Account",
]

MONTHLY_REQUIRED_HEADERS = [
    "DA Firm Name",
    "Sold By",
    "Total Commission Payable",
]

TEMPLATE_HEADERS = [
    "Recruiter",
    "Date Received",
    "Firm",
    "Receipt Name",
    "Payment to TRDA Club",
    "Payment due",
]


def normalise_text(value):
    if value is None:
        return ""
    return str(value).strip()


RECRUITER_ALIASES = {
    "steve howard": "Steven Howard",
    "steven howard": "Steven Howard",
}


def normalise_recruiter(value):
    name = normalise_text(value)
    name = re.sub(r"\s+", " ", name).strip()
    return RECRUITER_ALIASES.get(name.lower(), name)


def normalise_header(value):
    return re.sub(r"\s+", " ", normalise_text(value)).lower()


def previous_month_end(today=None):
    today = today or date.today()
    first_day_this_month = today.replace(day=1)
    return first_day_this_month - timedelta(days=1)


def clean_money(series):
    return (
        series.astype(str)
        .str.replace("£", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
        .replace({"": "0", "nan": "0", "None": "0"})
        .astype(float)
    )


def find_header_row(rows, required_headers):
    required = {normalise_header(h) for h in required_headers}
    for idx, row in enumerate(rows):
        row_headers = {normalise_header(c) for c in row if normalise_text(c)}
        if required.issubset(row_headers):
            return idx
    return None


def parse_support_cash_income(uploaded_file):
    """
    Parses the DA Support Cash Income Report.

    Assumption:
    - Recruiter/staff member name appears above each transaction table section.
    - Each section has a header row matching SUPPORT_REQUIRED_HEADERS.
    - Data continues until the next section header/recruiter or a blank break.
    """
    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = [[cell.value for cell in row] for row in ws.iter_rows()]

    parsed_rows = []
    current_recruiter = None
    i = 0

    required_norm = [normalise_header(h) for h in SUPPORT_REQUIRED_HEADERS]

    while i < len(rows):
        row = rows[i]
        non_empty = [normalise_text(c) for c in row if normalise_text(c)]

        # Candidate recruiter row: one main text value, not a recognised table header.
        row_norms = [normalise_header(c) for c in row if normalise_text(c)]
        looks_like_header = all(h in row_norms for h in required_norm)
        if non_empty and not looks_like_header:
            # Use the first non-empty cell as recruiter name when the row is mostly a title row.
            if len(non_empty) <= 3:
                current_recruiter = non_empty[0]

        if looks_like_header:
            header_row = row
            header_map = {}
            for col_idx, value in enumerate(header_row):
                h = normalise_header(value)
                if h:
                    header_map[h] = col_idx

            missing = [h for h in SUPPORT_REQUIRED_HEADERS if normalise_header(h) not in header_map]
            if missing:
                raise ValueError(f"Support report is missing columns: {', '.join(missing)}")

            i += 1
            while i < len(rows):
                data_row = rows[i]
                data_non_empty = [normalise_text(c) for c in data_row if normalise_text(c)]
                data_norms = [normalise_header(c) for c in data_row if normalise_text(c)]

                next_header = all(h in data_norms for h in required_norm)
                if next_header:
                    i -= 1
                    break

                # A short standalone text row after data likely means next recruiter.
                if data_non_empty and len(data_non_empty) <= 3:
                    # If it has no date/description/account values, treat as next recruiter.
                    date_val = data_row[header_map[normalise_header("Date")]] if header_map[normalise_header("Date")] < len(data_row) else None
                    desc_val = data_row[header_map[normalise_header("Description")]] if header_map[normalise_header("Description")] < len(data_row) else None
                    credit_val = data_row[header_map[normalise_header("Credit (GBP)")]] if header_map[normalise_header("Credit (GBP)")] < len(data_row) else None
                    if not date_val and not desc_val and not credit_val:
                        current_recruiter = data_non_empty[0]
                        break

                # Skip blank rows.
                if not data_non_empty:
                    i += 1
                    continue

                record = {h: None for h in SUPPORT_REQUIRED_HEADERS}
                for h in SUPPORT_REQUIRED_HEADERS:
                    col = header_map[normalise_header(h)]
                    record[h] = data_row[col] if col < len(data_row) else None

                # Only keep real transaction rows.
                if record["Description"] or record["Account"] or record["Credit (GBP)"]:
                    parsed_rows.append(
                        {
                            "Recruiter": normalise_recruiter(current_recruiter or "Unknown Recruiter"),
                            "Date Received": record["Date"],
                            "Firm": record["Description"],
                            "Receipt Name": record["Account"],
                            "Payment to TRDA Club": record["Credit (GBP)"],
                            "Payment due": None,
                            "Source File": "DA Support Cash Income Report",
                        }
                    )
                i += 1
        i += 1

    df = pd.DataFrame(parsed_rows)
    if df.empty:
        raise ValueError("No transaction rows were found in the support cash income report.")
    return df


def parse_monthly_statement(uploaded_file, statement_date):
    """Parses DA-Monthly Statement."""
    raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    rows = raw.values.tolist()
    header_idx = find_header_row(rows, MONTHLY_REQUIRED_HEADERS)
    if header_idx is None:
        raise ValueError(
            "Could not find the monthly statement header row containing: "
            + ", ".join(MONTHLY_REQUIRED_HEADERS)
        )

    df = pd.read_excel(uploaded_file, sheet_name=0, header=header_idx)
    df.columns = [normalise_text(c) for c in df.columns]

    missing = [c for c in MONTHLY_REQUIRED_HEADERS if c not in df.columns]
    if missing:
        raise ValueError(f"Monthly statement is missing columns: {', '.join(missing)}")

    df = df[MONTHLY_REQUIRED_HEADERS].copy()
    df = df.dropna(how="all")
    df = df[df["Sold By"].notna()]

    out = pd.DataFrame(
        {
            "Recruiter": df["Sold By"].apply(normalise_recruiter),
            "Date Received": statement_date,
            "Firm": df["DA Firm Name"],
            "Receipt Name": "Commission",
            "Payment to TRDA Club": df["Total Commission Payable"],
            "Payment due": None,
            "Source File": "DA-Monthly Statement",
        }
    )
    return out


def write_recruiter_template(template_file, recruiter_df):
    """Writes one recruiter's combined data into a copy of the uploaded template."""
    wb = load_workbook(template_file)
    ws = wb[wb.sheetnames[0]]

    # Ensure headers are in A2:F2.
    for col_idx, header in enumerate(TEMPLATE_HEADERS, start=1):
        ws.cell(row=2, column=col_idx).value = header

    # Clear old data below row 2 in A:F.
    for row in range(3, ws.max_row + 1):
        for col in range(1, 7):
            ws.cell(row=row, column=col).value = None

    start_row = 3
    for r, (_, row) in enumerate(recruiter_df.iterrows(), start=start_row):
        ws.cell(r, 1).value = row["Recruiter"]
        ws.cell(r, 2).value = row["Date Received"]
        ws.cell(r, 3).value = row["Firm"]
        ws.cell(r, 4).value = row["Receipt Name"]
        ws.cell(r, 5).value = row["Payment to TRDA Club"]
        ws.cell(r, 6).value = row["Payment due"]

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def safe_filename(name):
    cleaned = re.sub(r"[^A-Za-z0-9._ -]+", "", str(name)).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned or "Unknown_Recruiter"


def build_zip(template_file, combined_df, only_shared):
    zip_buffer = io.BytesIO()

    support_recruiters = set(
        combined_df.loc[combined_df["Source File"] == "DA Support Cash Income Report", "Recruiter"]
        .dropna()
        .astype(str)
        .str.strip()
    )
    monthly_recruiters = set(
        combined_df.loc[combined_df["Source File"] == "DA-Monthly Statement", "Recruiter"]
        .dropna()
        .astype(str)
        .str.strip()
    )

    if only_shared:
        recruiters = sorted(support_recruiters.intersection(monthly_recruiters))
    else:
        recruiters = sorted(set(combined_df["Recruiter"].dropna().astype(str).str.strip()))

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for recruiter in recruiters:
            recruiter_df = combined_df[combined_df["Recruiter"].astype(str).str.strip() == recruiter].copy()
            recruiter_df = recruiter_df[TEMPLATE_HEADERS + ["Source File"]]
            recruiter_df = recruiter_df.sort_values(["Date Received", "Firm"], kind="stable")

            # Need a fresh template stream for each recruiter.
            template_file.seek(0)
            workbook_bytes = write_recruiter_template(template_file, recruiter_df)
            filename = f"{safe_filename(recruiter)}_combined_statement.xlsx"
            zf.writestr(filename, workbook_bytes.getvalue())

    zip_buffer.seek(0)
    return zip_buffer, recruiters, support_recruiters, monthly_recruiters


# -----------------------------
# Streamlit UI
# -----------------------------

st.set_page_config(page_title="DA Statement Builder", page_icon="📄", layout="wide")
st.title("DA Statement Builder")
st.caption(APP_VERSION)
st.write(
    "Upload the DA Support Cash Income Report, DA-Monthly Statement, and statement template. "
    "The app will create one populated template per recruiter and package them in a ZIP."
)

with st.sidebar:
    st.header("Uploads")
    st.caption("Active alias: Steve Howard -> Steven Howard")
    support_file = st.file_uploader(
        "1. DA Support Cash Income Report",
        type=["xlsx", "xlsm", "xls"],
        key="support_file",
    )
    monthly_file = st.file_uploader(
        "2. DA-Monthly Statement",
        type=["xlsx", "xlsm", "xls"],
        key="monthly_file",
    )
    template_file = st.file_uploader(
        "3. Template",
        type=["xlsx", "xlsm", "xls"],
        key="template_file",
    )

    default_statement_date = previous_month_end()
    statement_date = st.date_input(
        "Date received for monthly statement rows",
        value=default_statement_date,
        help="Defaults to the last day of the previous month.",
    )

    only_shared = st.checkbox(
        "Only create files for recruiters found in both spreadsheets",
        value=True,
    )

run = st.button("Build recruiter statement ZIP", type="primary")

if run:
    if not support_file or not monthly_file or not template_file:
        st.error("Please upload all three files before building the ZIP.")
        st.stop()

    try:
        support_df = parse_support_cash_income(support_file)
        monthly_df = parse_monthly_statement(monthly_file, statement_date)

        combined_df = pd.concat([support_df, monthly_df], ignore_index=True)
        combined_df["Recruiter"] = combined_df["Recruiter"].apply(normalise_recruiter)

        zip_buffer, recruiters, support_recruiters, monthly_recruiters = build_zip(
            template_file=template_file,
            combined_df=combined_df,
            only_shared=only_shared,
        )

        st.success(f"Built {len(recruiters)} recruiter statement file(s).")

        col1, col2, col3 = st.columns(3)
        col1.metric("Support report recruiters", len(support_recruiters))
        col2.metric("Monthly statement recruiters", len(monthly_recruiters))
        col3.metric("Files in ZIP", len(recruiters))

        missing_from_monthly = sorted(support_recruiters - monthly_recruiters)
        missing_from_support = sorted(monthly_recruiters - support_recruiters)

        if missing_from_monthly:
            st.warning(
                "Recruiters found in support report but not monthly statement: "
                + ", ".join(missing_from_monthly)
            )
        if missing_from_support:
            st.warning(
                "Recruiters found in monthly statement but not support report: "
                + ", ".join(missing_from_support)
            )

        st.subheader("Preview of combined data")
        st.dataframe(combined_df[TEMPLATE_HEADERS + ["Source File"]], use_container_width=True)

        st.download_button(
            label="Download ZIP",
            data=zip_buffer,
            file_name=f"DA_recruiter_statements_{date.today().isoformat()}.zip",
            mime="application/zip",
        )

    except Exception as exc:
        st.error(str(exc))
        st.exception(exc)

else:
    st.info("Upload the three files, then click Build recruiter statement ZIP.")
