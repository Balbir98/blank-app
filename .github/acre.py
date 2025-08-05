# streamlit_app.py
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Boolean Cleaner", layout="wide")
st.title("CSV Boolean Cleaner")

uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])
if uploaded_file:
    # Read CSV into DataFrame
    df = pd.read_csv(uploaded_file)

    # The columns to transform
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
            return val  # leave blanks or other values unchanged

    # Apply transformation only to the specified columns that exist
    for col in bool_cols:
        if col in df.columns:
            df[col] = df[col].apply(transform_bool)

    st.subheader("Preview of Cleaned Data")
    st.dataframe(df)

    # CSV download
    csv_data = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download cleaned CSV",
        data=csv_data,
        file_name="cleaned_data.csv",
        mime="text/csv",
    )
else:
    st.info("Please upload a CSV file to get started.")
