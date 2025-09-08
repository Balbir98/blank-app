import re
from io import BytesIO
import pandas as pd
import streamlit as st
import os

# ---------------------------
# Utilities: robust readers
# ---------------------------

def _read_any_table(uploaded_file, preferred_sheet_name=None):
    """
    Read CSV or Excel into a DataFrame.
    - If file extension is .csv/.txt → try read_csv with sniffed delimiter (python engine, sep=None)
      and UTF-8 fallback to latin-1.
    - If Excel (.xlsx/.xls) → try preferred sheet name, then first sheet.
    """
    name = uploaded_file.name.lower()
    ext = os.path.splitext(name)[1]
    if ext in [".csv", ".txt"]:
        # Try UTF-8 with sep inference
        try:
            return pd.read_csv(uploaded_file, engine="python", sep=None)
        except Exception:
            uploaded_file.seek(0)
            # Fallback encoding
            return pd.read_csv(uploaded_file, engine="python", sep=None, encoding="latin-1")
    elif ext in [".xlsx", ".xls"]:
        if preferred_sheet_name:
            try:
                return pd.read_excel(uploaded_file, sheet_name=preferred_sheet_name)
            except Exception:
                uploaded_file.seek(0)
        # Fall back to first sheet
        return pd.read_excel(uploaded_file, sheet_name=0)
    else:
        # Attempt CSV as a last resort
        try:
            return pd.read_csv(uploaded_file, engine="python", sep=None)
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, engine="python", sep=None, encoding="latin-1")


# ---------------------------
# Transformation Logic
# ---------------------------

ID_COLS = ['Random ID', 'Name', 'Provider Name', 'Phone', 'Email']

def _is_event_label(x):
    """
    Decide if a subheader cell is an event/date-like label (month, quarter, or date string).
    """
    if pd.isna(x):
        return False
    s = str(x).strip()
    # Month-like (prefix check, tolerant of typos like "Febraury")
    months3 = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec']
    if any(s.lower().startswith(m) for m in months3):
        return True
    # Quarters / cadence
    if s in ['Q1','Q2','Q3','Q4','Monthly','Quarterly']:
        return True
    # Has digits (e.g., "4th February - Midlands", "17th June - Scotland")
    if re.search(r'\d', s):
        return True
    return False

def transform(form_df: pd.DataFrame, costs_df: pd.DataFrame) -> pd.DataFrame:
    """
    Transform a Zoho Forms export (wide sheet with section headers and subheaders in row 0)
    into the normalized "Desired Output" long format, and join "Cost" from a cost sheet.

    Assumptions:
    - Row 0 contains subheaders (months, event dates, or product labels) for the block's columns.
    - A "block" starts at a named column (not starting with "Unnamed"), followed by 0+ "Unnamed:" columns.
    - For unnamed columns:
        * If the subheader looks like a date/month/quarter, that becomes Event Date and the cell value is Product.
        * If the cell value contains "option" (e.g., "NTE Option"), the subheader text becomes the Product.
        * Otherwise the subheader text itself is the Product (Event Date blank).
    - Costs are joined on exact match of (Type, Product) after string trimming.
    """
    if form_df.shape[0] < 2:
        # Expect at least one header row (subheaders) + one data row
        return pd.DataFrame(columns=['Random ID','Provider Name','Phone','Email','Type','Event Date (if applicable)','Product','Cost'])

    # Row 0 is the subheader row; respondents start from row 1.
    subheaders = form_df.iloc[0]
    data_rows = form_df.iloc[1:].reset_index(drop=True)

    records = []
    for ridx, row in data_rows.iterrows():
        current_type = None
        for j, col in enumerate(form_df.columns):
            val = row[col]

            # Track the current Type on named (non-Unnamed) columns that are not identity/admin columns
            if not str(col).startswith('Unnamed'):
                if col not in ID_COLS and col not in ['Added Time', 'Referrer Name', 'Task Owner']:
                    current_type = col

            if pd.isna(val):
                continue

            if not str(col).startswith('Unnamed'):
                # Named column: usually a selection belonging to the Type at this column.
                if col in ID_COLS or col in ['Added Time', 'Referrer Name', 'Task Owner']:
                    continue
                prod = str(val).strip()
                sub = subheaders.iloc[j] if j < len(subheaders) else None
                evt = str(sub).strip() if _is_event_label(sub) else None
                records.append({
                    '_ridx': ridx,
                    'Type': current_type,
                    'Event Date (if applicable)': evt,
                    'Product': prod
                })
            else:
                # Unnamed column in the current block
                sub = subheaders.iloc[j] if j < len(subheaders) else None
                text_val = str(val).strip()

                if re.search(r'option', text_val, flags=re.I):
                    # If the cell contains "... Option", use the subheader text as Product
                    prod = str(sub).strip() if not pd.isna(sub) else text_val
                    evt = None
                else:
                    if _is_event_label(sub):
                        # subheader is an event/month/quarter label
                        evt = str(sub).strip()
                        prod = text_val
                    else:
                        # subheader is a product description; no event date
                        evt = None
                        prod = str(sub).strip() if not pd.isna(sub) else text_val

                if current_type is None:
                    # Safety: if we haven't encountered a named Type yet, skip
                    continue

                records.append({
                    '_ridx': ridx,
                    'Type': current_type,
                    'Event Date (if applicable)': evt,
                    'Product': prod
                })

    out = pd.DataFrame.from_records(records)

    # Attach identity columns for each record from its source row
    for c in ID_COLS:
        if c in data_rows.columns:
            out[c] = out['_ridx'].map(data_rows[c])
        else:
            out[c] = None

    # Arrange final columns
    out = out[['Random ID','Provider Name','Phone','Email','Type','Event Date (if applicable)','Product']]

    # Cost join on exact (Type, Product) after trimming
    def _norm(s):
        return None if pd.isna(s) else str(s).strip()

    # Costs sheet must have columns: Type, Product, Cost
    costs_df2 = costs_df.copy()
    if not set(['Type','Product','Cost']).issubset(set(costs_df2.columns)):
        raise ValueError("Cost sheet must contain columns: Type, Product, Cost")

    costs_df2['Type_norm'] = costs_df2['Type'].apply(_norm)
    costs_df2['Product_norm'] = costs_df2['Product'].apply(_norm)
    out['Type_norm'] = out['Type'].apply(_norm)
    out['Product_norm'] = out['Product'].apply(_norm)

    out = out.merge(
        costs_df2[['Type_norm','Product_norm','Cost']],
        on=['Type_norm','Product_norm'],
        how='left'
    ).drop(columns=['Type_norm','Product_norm'])

    # Optional: tidy sort for readability
    out = out.sort_values(['Random ID','Type','Event Date (if applicable)','Product']).reset_index(drop=True)
    return out


# ---------------------------
# Streamlit App
# ---------------------------

st.set_page_config(page_title="Zoho Forms Cleaner", layout="wide")
st.title("Zoho Forms → Cleaned Output")

st.markdown("""
**How it works**
1. Upload your raw Zoho Forms export (**CSV** or Excel).
2. Upload the MOF Cost Sheet (**CSV** or Excel) with columns: **Type, Product, Cost**.
3. (Optional) Upload a template (we'll wire this in the next phase).
4. Click **Transform** to get the normalized output (matches your Desired Output structure) with Costs filled in.
""")

c1, c2, c3 = st.columns(3)
with c1:
    form_file = st.file_uploader("1) Upload Zoho Forms export (.csv/.xlsx/.xls)", type=["csv","xlsx","xls","txt"])
with c2:
    cost_file = st.file_uploader("2) Upload MOF Cost Sheet (.csv/.xlsx/.xls)", type=["csv","xlsx","xls","txt"])
with c3:
    template_file = st.file_uploader("3) Upload Template (optional, coming later)", type=["xlsx","xls","csv","docx","pptx","pdf"])

if st.button("Transform to Desired Output"):
    if not form_file or not cost_file:
        st.error("Please upload both the Zoho Forms export and the MOF Cost Sheet.")
    else:
        with st.spinner("Transforming..."):
            try:
                form_df = _read_any_table(form_file, preferred_sheet_name="Form")
            except Exception as e:
                st.exception(RuntimeError(f"Failed to read the Zoho Forms export: {e}"))
                st.stop()

            try:
                costs_df = _read_any_table(cost_file)
            except Exception as e:
                st.exception(RuntimeError(f"Failed to read the MOF Cost Sheet: {e}"))
                st.stop()

            try:
                cleaned = transform(form_df, costs_df)
            except Exception as e:
                st.exception(e)
            else:
                st.success(f"Cleaned {len(cleaned)} rows.")
                st.dataframe(cleaned, use_container_width=True)

                # Quick quality signal on unmatched costs
                missing_costs = cleaned['Cost'].isna().sum()
                if missing_costs > 0:
                    st.warning(f"{missing_costs} row(s) have no Cost match. Ensure (Type, Product) exist in the MOF Cost Sheet.")
                    with st.expander("Preview rows missing Cost"):
                        st.dataframe(
                            cleaned[cleaned['Cost'].isna()][
                                ['Type','Event Date (if applicable)','Product']
                            ].head(200)
                        )

                # Download cleaned output
                buf = BytesIO()
                cleaned.to_excel(buf, index=False)
                st.download_button(
                    "Download cleaned_output.xlsx",
                    data=buf.getvalue(),
                    file_name="cleaned_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

st.markdown("---")
st.caption("Template population coming soon. When you're ready, we’ll wire the cleaned dataframe into your template.")
