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

    # --- Application Date -> text "dd/MM/yyyy" (others untouched) ---
    FORCE_TEXT_WITH_FORMULA_WRAPPER = True  # set to False to output plain "dd/MM/yyyy" text without ="" trick

    if "Application Date" in df.columns:
        s = df["Application Date"].astype(str).str.strip()

        # Treat placeholder strings as blanks
        placeholders = s.str.lower().isin({"nan", "none", "null"})
        s = s.mask(placeholders, "")

        # Parse UK first (handles dd/mm/yyyy and any time part if present)
        parsed = pd.to_datetime(s, errors="coerce", dayfirst=True)

        # For still-NaT rows that look ISO, try strict ISO yyyy-mm-dd
        need_iso = parsed.isna() & s.str.contains("-", na=False)
        if need_iso.any():
            parsed.loc[need_iso] = pd.to_datetime(
                s[need_iso], errors="coerce", format="%Y-%m-%d"
            )

        # Format to dd/MM/yyyy; keep blank where unparseable
        as_text = parsed.dt.strftime("%d/%m/%Y")
        as_text = as_text.where(parsed.notna(), "")

        if FORCE_TEXT_WITH_FORMULA_WRAPPER:
            # Display as 01/01/2025 but force text in Excel/Zoho by using ="01/01/2025"
            df["Application Date"] = as_text.where(
                as_text.eq(""),
                '="' + as_text + '"'
            )
        else:
            # Plain text dd/MM/yyyy (may still be parsed by some tools)
            df["Application Date"] = as_text

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

    # UTF-8 with BOM to preserve £ in headers/data
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
