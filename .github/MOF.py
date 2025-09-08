import re
import os
from io import BytesIO
import zipfile

import pandas as pd
import streamlit as st

# Excel handling
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

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
        try:
            return pd.read_csv(uploaded_file, engine="python", sep=None)
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, engine="python", sep=None, encoding="latin-1")
    elif ext in [".xlsx", ".xls"]:
        if preferred_sheet_name:
            try:
                return pd.read_excel(uploaded_file, sheet_name=preferred_sheet_name)
            except Exception:
                uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, sheet_name=0)
    else:
        try:
            return pd.read_csv(uploaded_file, engine="python", sep=None)
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, engine="python", sep=None, encoding="latin-1")


# ---------------------------
# Transformation Logic
# ---------------------------

ID_COLS = ['Random ID', 'Provider Name', 'Name', 'Phone', 'Email']

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
    - For *named* columns, the same "option" rule now applies (bug fix).
    - Costs are joined on exact match of (Type, Product) after string trimming.
    """
    if form_df.shape[0] < 2:
        return pd.DataFrame(columns=[
            'Random ID','Provider Name','Name','Phone','Email',
            'Type','Event Date (if applicable)','Product','Cost'
        ])

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

            sub = subheaders.iloc[j] if j < len(subheaders) else None
            text_val = str(val).strip()

            if not str(col).startswith('Unnamed'):
                # Named column: apply "Option" rule here too (bug fix).
                if col in ID_COLS or col in ['Added Time', 'Referrer Name', 'Task Owner']:
                    continue

                if re.search(r'\boption(s)?\b', text_val, flags=re.I) and not pd.isna(sub):
                    # If the cell is "Option"/"... Option", use subheader text as Product
                    prod = str(sub).strip()
                    evt = None if not _is_event_label(sub) else str(sub).strip()  # usually sub is product here
                    # If sub looks like date (rare for an 'Option' case), treat it as event and keep prod=text_val
                    if _is_event_label(sub):
                        prod = text_val
                else:
                    prod = text_val
                    evt = str(sub).strip() if _is_event_label(sub) else None

                records.append({
                    '_ridx': ridx,
                    'Type': current_type,
                    'Event Date (if applicable)': evt,
                    'Product': prod
                })

            else:
                # Unnamed column in the current block (existing logic)
                if re.search(r'\boption(s)?\b', text_val, flags=re.I):
                    prod = str(sub).strip() if not pd.isna(sub) else text_val
                    evt = None
                else:
                    if _is_event_label(sub):
                        evt = str(sub).strip()
                        prod = text_val
                    else:
                        evt = None
                        prod = str(sub).strip() if not pd.isna(sub) else text_val

                if current_type is None:
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
    out = out[['Random ID','Provider Name','Name','Phone','Email','Type','Event Date (if applicable)','Product']]

    # Cost join on exact (Type, Product) after trimming
    def _norm(s):
        return None if pd.isna(s) else str(s).strip()

    # Costs sheet must have columns: Type, Product, Cost (and may have 'F2F or Online?')
    costs_df2 = costs_df.copy()
    if not set(['Type','Product','Cost']).issubset(set(costs_df2.columns)):
        raise ValueError("Cost sheet must contain columns: Type, Product, Cost")

    costs_df2['Type_norm'] = costs_df2['Type'].apply(_norm)
    costs_df2['Product_norm'] = costs_df2['Product'].apply(_norm)
    out['Type_norm'] = out['Type'].apply(_norm)
    out['Product_norm'] = out['Product'].apply(_norm)

    # Merge Cost
    out = out.merge(
        costs_df2[['Type_norm','Product_norm','Cost']],
        on=['Type_norm','Product_norm'],
        how='left'
    ).drop(columns=['Type_norm','Product_norm'])

    # Optional: tidy sort
    out = out.sort_values(['Random ID','Type','Event Date (if applicable)','Product']).reset_index(drop=True)
    return out


# ---------------------------
# Template population
# ---------------------------

def _sanitize_name(name: str) -> str:
    """Remove commas and excessive spaces for 'Name' going into Main contact."""
    if pd.isna(name):
        return ""
    s = str(name).replace(",", " ")
    return re.sub(r"\s+", " ", s).strip()

def _find_table_header_row(ws, header_candidates=("Type","Product","Month","Date")):
    """
    Find the row index (1-based) where the first two headers match 'Type' and 'Product',
    and 'Month' and 'Date' appear to the right somewhere on the same row.
    Returns the header row number, first header column index, and header map {header: col_idx}.
    """
    max_row = min(ws.max_row, 200)
    max_col = min(ws.max_column, 40)

    for r in range(1, max_row + 1):
        # Collect row values
        row_vals = [ws.cell(r, c).value for c in range(1, max_col + 1)]
        # try to locate "Type"
        for c in range(1, max_col + 1):
            if (str(row_vals[c-1]).strip().lower() == "type" if row_vals[c-1] is not None else False):
                # Expect next header is Product
                if c + 1 <= max_col and (str(row_vals[c]).strip().lower() == "product" if row_vals[c] is not None else False):
                    # Look for Month and Date somewhere to the right
                    month_ok = any((str(v).strip().lower() == "month") if v is not None else False for v in row_vals[c:])
                    date_ok = any((str(v).strip().lower() == "date") if v is not None else False for v in row_vals[c:])
                    if month_ok and date_ok:
                        # Build header map for all visible headers on this row
                        hmap = {}
                        for c2 in range(c, max_col + 1):
                            v = ws.cell(r, c2).value
                            if v is None:
                                continue
                            key = str(v).strip()
                            hmap[key] = c2
                        return r, c, hmap
    raise RuntimeError("Could not find the table header row with columns like Type | Product | Month | Date ...")

def _populate_template_bytes(template_bytes: bytes, cleaned: pd.DataFrame, costs_df: pd.DataFrame) -> BytesIO:
    """
    Creates a ZIP of one populated template per provider.
    - Fills fixed cells
    - Inserts rows for each product line under the detected table
    - Preserves template formulas by using openpyxl insert_rows which shifts cells down
    """
    # Optional lookup for F2F/Online from costs sheet (matched by Product)
    f2f_map = {}
    if 'F2F or Online?' in costs_df.columns:
        # normalize product string
        def _n(s): return None if pd.isna(s) else str(s).strip()
        tmp = costs_df[['Product','F2F or Online?']].copy()
        tmp['Product_norm'] = tmp['Product'].apply(_n)
        f2f_map = dict(zip(tmp['Product_norm'], tmp['F2F or Online?']))

    # Group cleaned data by Provider Name
    groups = cleaned.groupby('Provider Name', dropna=False)

    zip_buf = BytesIO()
    with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for provider, dfp in groups:
            # Load template fresh each time
            wb = load_workbook(BytesIO(template_bytes))
            ws = wb.active  # assume first sheet

            # Fixed cells
            provider_val = "" if pd.isna(provider) else str(provider)
            name_val = _sanitize_name(dfp['Name'].iloc[0] if 'Name' in dfp.columns and len(dfp) > 0 else "")
            phone_val = "" if 'Phone' not in dfp.columns else ("" if pd.isna(dfp['Phone'].iloc[0]) else str(dfp['Phone'].iloc[0]))
            email_val = "" if 'Email' not in dfp.columns else ("" if pd.isna(dfp['Email'].iloc[0]) else str(dfp['Email'].iloc[0]))

            # Cells: B7 (Provider), B9 (Main contact), G9 (Phone), I9 (Email),
            # and email again at B12, I12
            ws["B7"] = provider_val
            ws["B9"] = name_val
            ws["G9"] = phone_val
            ws["I9"] = email_val
            ws["B12"] = email_val
            ws["I12"] = email_val

            # Locate table header row and columns
            hdr_row, hdr_col, hmap = _find_table_header_row(ws)

            start_data_row = hdr_row + 1
            n_lines = len(dfp)

            # Ensure exactly n_lines rows available under header.
            # If the template has one sample row, we need to insert (n_lines - 1) rows.
            if n_lines > 1:
                ws.insert_rows(start_data_row + 1, amount=n_lines - 1)

            # Column resolution helpers
            def col_of(header_text, default=None):
                # Accept variations for "F2F or Online?"
                aliases = {
                    "F2F or Online?": ["F2F or Online?","F2F/Online","F2F or Online"],
                    "Qty": ["Qty","Quantity"],
                    "Charge": ["Charge","Cost","Price"],
                    "Month": ["Month","Event Month","Month of"],
                }
                keys = [header_text] + aliases.get(header_text, [])
                for k in keys:
                    if k in hmap: 
                        return hmap[k]
                return default

            col_Type   = col_of("Type")
            col_Prod   = col_of("Product")
            col_Month  = col_of("Month")
            col_Date   = col_of("Date")
            col_F2F    = col_of("F2F or Online?")
            col_Qty    = col_of("Qty")
            col_Charge = col_of("Charge")

            # Write rows
            for i, (_, r) in enumerate(dfp.iterrows()):
                rr = start_data_row + i
                # Values
                typ = r.get("Type", "")
                prod = r.get("Product", "")
                month = r.get("Event Date (if applicable)", "")
                qty = 1
                charge = r.get("Cost", "")

                # F2F lookup by Product (exact, trimmed)
                prod_key = None if pd.isna(prod) else str(prod).strip()
                f2f_val = f2f_map.get(prod_key, "")

                if col_Type:   ws.cell(rr, col_Type,   typ)
                if col_Prod:   ws.cell(rr, col_Prod,   prod)
                if col_Month:  ws.cell(rr, col_Month,  month)
                if col_Date:   ws.cell(rr, col_Date,   None)  # left blank
                if col_F2F:    ws.cell(rr, col_F2F,    f2f_val)
                if col_Qty:    ws.cell(rr, col_Qty,    qty)
                if col_Charge: ws.cell(rr, col_Charge, charge)

            # Save one workbook per provider into ZIP
            out_bytes = BytesIO()
            wb.save(out_bytes)
            out_bytes.seek(0)

            # Safe filename
            safe_provider = re.sub(r'[^A-Za-z0-9 _.-]+', '_', provider_val or "Unknown_Provider")
            zf.writestr(f"{safe_provider}.xlsx", out_bytes.getvalue())

    zip_buf.seek(0)
    return zip_buf


# ---------------------------
# Streamlit App
# ---------------------------

st.set_page_config(page_title="Zoho Forms Cleaner + Template Populator", layout="wide")
st.title("Zoho Forms → Cleaned Output + Populated Templates")

st.markdown("""
**How it works**
1. Upload your raw Zoho Forms export (**CSV** or Excel).
2. Upload the MOF Cost Sheet (**CSV** or Excel) with columns: **Type, Product, Cost** (optionally **F2F or Online?**).
3. (Optional) Upload your **template** (Excel).  
4. Click **Transform** to get the normalized output + costs.  
   - If a template is provided, you'll also get a **ZIP** with one populated workbook **per Provider Name**.
""")

c1, c2, c3 = st.columns(3)
with c1:
    form_file = st.file_uploader("1) Upload Zoho Forms export (.csv/.xlsx/.xls)", type=["csv","xlsx","xls","txt"])
with c2:
    cost_file = st.file_uploader("2) Upload MOF Cost Sheet (.csv/.xlsx/.xls)", type=["csv","xlsx","xls","txt"])
with c3:
    template_file = st.file_uploader("3) Upload Template (optional, Excel)", type=["xlsx","xls"])

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

            # Clean the data
            try:
                cleaned = transform(form_df, costs_df)
            except Exception as e:
                st.exception(e)
                st.stop()

            st.success(f"Cleaned {len(cleaned)} rows.")
            st.dataframe(cleaned, use_container_width=True)

            # Quality signal on unmatched costs
            missing_costs = cleaned['Cost'].isna().sum()
            if missing_costs > 0:
                st.warning(f"{missing_costs} row(s) have no Cost match. Ensure (Type, Product) exist in the MOF Cost Sheet.")
                with st.expander("Preview rows missing Cost"):
                    st.dataframe(
                        cleaned[cleaned['Cost'].isna()][
                            ['Provider Name','Type','Event Date (if applicable)','Product']
                        ].head(500)
                    )

            # Download cleaned output
            cleaned_buf = BytesIO()
            cleaned.to_excel(cleaned_buf, index=False)
            st.download_button(
                "Download cleaned_output.xlsx",
                data=cleaned_buf.getvalue(),
                file_name="cleaned_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # If template uploaded → also build ZIP of populated templates
            if template_file is not None:
                try:
                    template_bytes = template_file.read()
                    zip_buf = _populate_template_bytes(template_bytes, cleaned, costs_df)
                except Exception as e:
                    st.exception(RuntimeError(f"Template population failed: {e}"))
                else:
                    st.download_button(
                        "Download populated_templates.zip",
                        data=zip_buf.getvalue(),
                        file_name="populated_templates.zip",
                        mime="application/zip"
                    )

st.markdown("---")
st.caption("If you want custom header cell locations, alternate sheet names, or fuzzy product-to-cost matching, I can add options above.")
