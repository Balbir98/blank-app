# streamlit_app.py
import streamlit as st
import pandas as pd
import io
import csv
import time

st.set_page_config(page_title="Boolean Cleaner", layout="wide")
st.title("CSV Boolean Cleaner")

uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])
if uploaded_file:

    @st.cache_data
    def load_csv(file):
        # force these ID columns to stay exact
        dtype_map = {"Adviser ID": str, "Firm ID": str}

        # Try the default UTF-8/c-engine with our dtype map
        try:
            return pd.read_csv(file, dtype=dtype_map)
        except Exception:
            # Fallback: raw bytes → text (utf-8 or latin-1) → sniff delimiter
            file.seek(0)
            raw_bytes = file.read()
            try:
                raw_text = raw_bytes.decode("utf-8")
            except UnicodeDecodeError:
                raw_text = raw_bytes.decode("latin-1")

            dialect = csv.Sniffer().sniff(raw_text[:10_000])
            sep = dialect.delimiter

            return pd.read_csv(
                io.StringIO(raw_text),
                sep=sep,
                engine="python",
                dtype=dtype_map,
            )

    df = load_csv(uploaded_file)

    # --- Robust date normalization (UK dd/mm/yyyy) ---
    date_cols = [
        "Application Date",
        "Effective Date",
        "Benefit End Date",
        "Created Date",
        "Last Updated",
        "Earliest Version Date",
        "Older Version Date",
    ]

    def format_dates_uk(s: pd.Series) -> pd.Series:
        # Keep original so we can preserve blanks
        orig = s.copy()

        # Try parsing common string dates (dd/mm/yyyy, yyyy-mm-dd, with/without time)
        parsed = pd.to_datetime(
            s, errors="coerce", dayfirst=True, infer_datetime_format=True, utc=False
        )

        # Handle Excel serial numbers (e.g., 45231) where parse failed
        num = pd.to_numeric(s, errors="coerce")
        serial_mask = parsed.isna() & num.notna()
        if serial_mask.any():
            parsed.loc[serial_mask] = pd.to_datetime("1899-12-30") + pd.to_timedelta(
                num[serial_mask], unit="D"
            )

        # Format as dd/mm/yyyy where we have a valid datetime
        out = parsed.dt.strftime("%d/%m/%Y")

        # Preserve blanks: if original is NaN/empty/whitespace, keep empty string
        is_blank = orig.isna() | (orig.astype(str).str.strip() == "")
        out = out.where(~is_blank, "")

        return out

    # Apply date formatting only to columns that exist
    for col in date_cols:
        if col in df.columns:
            df[col] = format_dates_uk(df[col])

    # --- Boolean cleanup (unchanged) ---
    bool_cols = [
        "High Risk", "Whole Of Life", "In Trust", "BTL", "Adverse",
        "Self Cert", "Off Panel", "Introduced?", "Been Checked?",
        "App2 Blank?", "Lending into retirement?", "Second Charge?"
    ]

    def transform_bool(val):
        if val == "t":
            return True
        elif val == "f":
            return False
        else:
            return val

    prog = st.progress(0)
    status = st.empty()

    cols_to_do = [c for c in bool_cols if c in df.columns]
    # +1 for date formatting +1 for CSV generation
    total_steps = len(cols_to_do) + 2
    step = 0

    # Boolean transformation
    for col in cols_to_do:
        df[col] = df[col].apply(transform_bool)
        step += 1
        prog.progress(step / total_steps)
        status.text(f"Transforming “{col}” — ≈ {total_steps-step}s remaining")

    status.text("Generating cleaned CSV…")
    # UTF-8 BOM to preserve £ and other symbols
    csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    step += 1
    prog.progress(step / total_steps)
    status.text("All done!")

    st.download_button(
        label="Download cleaned CSV",
        data=csv_bytes,
        file_name="cleaned_data.csv",
        mime="text/csv",
    )

else:
    st.info("Please upload a CSV file to get started.")
