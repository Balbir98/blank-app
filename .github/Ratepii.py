# streamlit_app.py
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Mortgage Rate Review PII", layout="centered")

st.title("Mortgage Rate Review PII")
st.write(
    "Upload **two** Acre exports: the **Mortgage Rate Review** report and the **Combined Case Report**. "
    "Then click **Generate report** to download a merged CSV that includes PII fields."
)

# --- UI: file uploads ---
col1, col2 = st.columns(2)
with col1:
    rr_file = st.file_uploader(
        "Mortgage Rate Review report",
        type=["csv", "xlsx", "xls"],
        key="rr_file",
        help="Upload the Mortgage Rate Review export (CSV or Excel)."
    )
with col2:
    cc_file = st.file_uploader(
        "Combined Case Report",
        type=["csv", "xlsx", "xls"],
        key="cc_file",
        help="Upload the Combined Case Report export (CSV or Excel)."
    )

st.divider()

# --- constants / config ---
EXTRA_HEADERS = [
    "Full name",
    "Dob","Address1","Address2","Address3","Posttown","Postcode","County","Country",
    "Email address","Mobile phone","Home phone","Work phone",
    "Created year","Created month","Created week","Created at",
    "Case type","Regulated","Case status","Mortgage status",
    "Mortgage amount","Property value","Term","Term unit","N clients","Ltv",
]

CC_REQUIRED_BASE = [
    "First name","Last name","Dob","Address1","Address2","Address3","Posttown",
    "Postcode","County","Country","Email address","Mobile phone","Home phone",
    "Work phone","Created year","Created month","Created week","Created at",
    "Case type","Regulated","Case status","Mortgage status","Mortgage amount",
    "Property value","Term","Term unit","N clients","Ltv"
]

JOIN_KEY = "Case id"

OUTPUT_ORDER = [
    "Advisor name",
    "Case id",
    "Case URL",
    "Mtg completion date",
    "Mortgage id",
    "Lender name",
    "Status",
    "Initial rate",
    "Initial rate end date",
    "Current reminder date",
    "Reminder status",
    "Full name",
    "Dob",
    "Address1",
    "Address2",
    "Address3",
    "Posttown",
    "Postcode",
    "County",
    "Country",
    "Email address",
    "Mobile phone",
    "Home phone",
    "Work phone",
    "Created at",
    "Case type",
    "Case status",
    "Property value",
    "N clients",
]

RR_EXPECTED_COLS = [
    "Advisor name",
    "Case id",
    "Mtg completion date",
    "Mortgage id",
    "Lender name",
    "Status",
    "Initial rate",
    "Initial rate end date",
    "Current reminder date",
    "Reminder status",
]

def read_any(file) -> pd.DataFrame:
    if file is None:
        return None
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(file)
    raise ValueError("Unsupported file type. Please upload CSV or Excel.")

def norm(s: str) -> str:
    return " ".join(str(s).strip().replace("_", " ").split()).lower()

def build_lookup(cols):
    return {norm(c): c for c in cols}

def find_col(lookup: dict, target_name: str, alt_variants=None):
    candidates = [target_name] + (alt_variants or [])
    t = target_name
    if " id" in t.lower():
        candidates += [t.replace(" id", " ID"), t.replace(" id", " Id")]
    for c in candidates:
        hit = lookup.get(norm(c))
        if hit:
            return hit
    return None

def clean_iso_date_to_ddmmyyyy(series: pd.Series) -> pd.Series:
    """
    Accepts strings like '2024-05-04T15:43:03Z' (or other parseables) and returns '04/05/2024'.
    Unparseable values become blank strings.
    """
    s = series.astype(str).str.strip()
    # remove anything from 'T' to the end (keeps 'YYYY-MM-DD' part if present)
    s = s.str.replace(r"T.*$", "", regex=True)
    # parse (supports yyyy-mm-dd and many others; dayfirst handles dd/mm/yyyy inputs too)
    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return dt.dt.strftime("%d/%m/%Y").fillna("")

def safe_merge(rr: pd.DataFrame, cc: pd.DataFrame):
    rr_lk = build_lookup(rr.columns)
    cc_lk = build_lookup(cc.columns)

    rr_key = find_col(rr_lk, JOIN_KEY, alt_variants=["Case ID", "case id", "CaseId"])
    cc_key = find_col(cc_lk, JOIN_KEY, alt_variants=["Case ID", "case id", "CaseId"])
    if rr_key is None or cc_key is None:
        missing_side = []
        if rr_key is None:
            missing_side.append("Mortgage Rate Review")
        if cc_key is None:
            missing_side.append("Combined Case Report")
        raise KeyError(
            f"Couldn't find '{JOIN_KEY}' in: {', '.join(missing_side)}. "
            f"Make sure both reports include a '{JOIN_KEY}' column."
        )

    # Map CC columns we need
    cc_map = {}
    missing_in_cc = []
    for base in CC_REQUIRED_BASE:
        col = find_col(cc_lk, base)
        if col is None:
            missing_in_cc.append(base)
        else:
            cc_map[base] = col

    # Compute Full name
    first_col = cc_map.get("First name")
    last_col  = cc_map.get("Last name")
    cc_subset = cc[[cc_key] + [v for k, v in cc_map.items() if k not in ("First name","Last name")]].copy()
    if first_col is None or last_col is None:
        cc_subset["Full name"] = ""
    else:
        cc_subset["Full name"] = (
            cc[first_col].fillna("").astype(str).str.strip() + " " +
            cc[last_col].fillna("").astype(str).str.strip()
        ).str.strip()

    # Build canonical CC output only for fields needed in final output
    needed_from_cc = list(set(OUTPUT_ORDER) & set(EXTRA_HEADERS))
    source = {k: cc[v] for k, v in cc_map.items() if k not in ("First name", "Last name")}
    source["Full name"] = cc_subset["Full name"]

    cc_out = pd.DataFrame()
    cc_out[cc_key] = cc[cc_key]
    for h in needed_from_cc:
        cc_out[h] = source.get(h, pd.NA)

    # Merge (left join keeps all RR rows; duplicates in CC may expand rows)
    merged = rr.merge(
        cc_out,
        left_on=rr_key,
        right_on=cc_key,
        how="left",
        indicator=True
    )

    # Case URL built from RR join key
    merged["Case URL"] = "https://crm.myac.re/cases/" + merged[rr_key].astype(str).str.strip() + "/overview"

    # Ensure "Case id" column exists canonically
    if "Case id" not in merged.columns:
        merged["Case id"] = merged[rr_key]

    # Build final dataframe in requested order (map RR columns if they differ by case/spacing)
    rr_mapped_cols = {c: find_col(rr_lk, c) for c in RR_EXPECTED_COLS}
    final_df = pd.DataFrame()
    for col in OUTPUT_ORDER:
        if col in merged.columns:
            final_df[col] = merged[col]
        elif col in rr_mapped_cols and rr_mapped_cols[col] in merged.columns:
            final_df[col] = merged[rr_mapped_cols[col]]
        else:
            final_df[col] = pd.NA

    # --- Date cleaning for two columns ---
    for date_col in ["Mtg completion date", "Current reminder date"]:
        if date_col in final_df.columns:
            final_df[date_col] = clean_iso_date_to_ddmmyyyy(final_df[date_col])

    # Metrics
    matched = int((merged["_merge"] == "both").sum()) if "_merge" in merged.columns else None
    total = len(merged)
    if "_merge" in merged.columns:
        merged = merged.drop(columns=["_merge"])

    return final_df, {
        "rr_key": rr_key,
        "cc_key": cc_key,
        "matched": matched,
        "total": total,
        "missing_in_cc": missing_in_cc
    }

def generate_download(df: pd.DataFrame, default_name="mortgage_rate_review_pii.csv"):
    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="⬇️ Download Mortgage Rate Review PII CSV",
        data=csv_bytes,
        file_name=default_name,
        mime="text/csv",
        use_container_width=True
    )

# --- Action button ---
if st.button("Generate report", type="primary", use_container_width=True):
    if rr_file is None or cc_file is None:
        st.error("Please upload **both** files before generating the report.")
    else:
        try:
            rr_df = read_any(rr_file)
            cc_df = read_any(cc_file)

            if rr_df is None or rr_df.empty:
                st.error("The Mortgage Rate Review report appears to be empty.")
            elif cc_df is None or cc_df.empty:
                st.error("The Combined Case Report appears to be empty.")
            else:
                with st.spinner("Merging reports…"):
                    output_df, stats = safe_merge(rr_df, cc_df)

                st.success(
                    f"Report is ready! Matched {stats['matched']} of {stats['total']} rows by **{JOIN_KEY}**."
                )
                if stats["missing_in_cc"]:
                    st.warning(
                        "Some PII fields were missing in the Combined Case Report and were left blank: "
                        + ", ".join(stats["missing_in_cc"])
                    )

                generate_download(output_df)

        except KeyError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Something went wrong while generating the report: {e}")
            