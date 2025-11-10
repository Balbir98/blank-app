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

    # --- Date columns to output as TEXT "dd/MM/yyyy" ---
    date_cols = [
        "Application Date",
        "Effective Date",
        "Benefit End Date",
        "Created Date",
        "Last Updated",
        "Earliest Version Date",
        "Older Version Date",
    ]

    def normalize_date_to_text_ddmmyyyy(series: pd.Series) -> pd.Series:
        """
        Return TEXT dd/MM/yyyy.
        Handles dd/mm/yyyy (± time), yyyy-mm-dd (± time), and Excel serials.
        Leaves blanks as blanks. Keeps unparseable values as-is.
        """
        s = series.astype(str).str.strip()

        # Treat explicit placeholders as blanks
        placeholders = s.str.lower().isin({"nan", "none", "null", "nat"})
        s = s.mask(placeholders, "")

        # 1) Parse UK (dd/mm/yyyy [hh:mm[:ss]])
        parsed = pd.to_datetime(s, errors="coerce", dayfirst=True)

        # 2) Parse ISO (yyyy-mm-dd [hh:mm[:ss]]) where still NaT and looks dashy
        need_iso = parsed.isna() & s.str.contains("-", na=False)
        if need_iso.any():
            parsed.loc[need_iso] = pd.to_datetime(s[need_iso], errors="coerce")

        # 3) Parse Excel serials (e.g., 45231 or 45231.0)
        still_nat = parsed.isna() & s.ne("")
        if still_nat.any():
            # remove thousands separators if any
            nums = pd.to_numeric(s[still_nat].str.replace(",", ""), errors="coerce")
            has_num = nums.notna()
            if has_num.any():
                parsed.loc[still_nat[still_nat].index[has_num]] = (
                    pd.to_datetime("1899-12-30") + pd.to_timedelta(nums[has_num], unit="D")
                )

        # Format to dd/MM/yyyy where parsed; blanks stay blank; keep originals where still NaT
        out = parsed.dt.strftime("%d/%m/%Y")
        out = out.where(parsed.notna(), s)  # if still NaT, keep original text (including "")

        # Ensure TEXT (not date) in the CSV
        return out.astype(str)

    # Apply to all listed date columns that exist
    for col in date_cols:
        if col in df.columns:
            df[col] = normalize_date_to_text_ddmmyyyy(df[col])

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
    total_steps = len(cols_to_do) + 1
    step = 0

    for col in cols_to_do:
        df[col] = df[col].apply(transform_bool)
        step += 1
        prog.progress(step / total_steps)
        status.text(f"Transforming “{col}” — ≈ {total_steps-step}s remaining")

    status.text("Generating cleaned CSV…")

    # UTF-8 with BOM to preserve £ and other symbols reliably
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
