import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="EDI Converter", page_icon="📄", layout="centered")

st.title("EDI Converter")
st.caption("Convert provider EDI files (EDIFACT COMTFR) into a clean CSV.")

provider = st.selectbox(
    "Provider",
    ["L&G", "Aviva"],
    index=0,
    help="L&G is implemented. Aviva is a placeholder for now."
)
uploaded = st.file_uploader(
    "Upload the raw EDI text file",
    type=["txt", "edi", "dat"],
    accept_multiple_files=False
)
convert_clicked = st.button("Convert!", type="primary", disabled=(uploaded is None))


# -------------------- helpers --------------------
def parse_edi_segments(text: str):
    """Split EDIFACT content into segments on apostrophes, with a small normalisation."""
    raw = text.replace("\r\n", "\n").replace("\r", "\n")
    raw = re.sub(r"^End-of-Header:\s*\n", "", raw, flags=re.MULTILINE)
    segs = [s.strip() for s in raw.split("'") if s.strip()]
    return segs


def tokenise(segment: str):
    parts = segment.split("+")
    return parts[0], parts[1:]


def keep_spaces(token: str, width: int = None):
    """Keep EDIFACT spacing (e.g., 'R  '), padding to width if provided."""
    if token is None:
        return None
    if width and len(token) < width:
        return token + (" " * (width - len(token)))
    return token


def parse_chd_fields(fields):
    """
    Parse CHD -> commission/charge line: amount, currency, charge type (CBS/ACH/CCH),
    due date CDD, and premium composite (type/amount/currency).
    """
    d = {
        "chd_amt_qual": None,      # I/R/X with spaces preserved
        "chd_amount": None,
        "chd_currency": None,
        "chd_cur_qual": None,
        "charge_type": None,
        "due_date": None,
        "due_date_fmt": None,
        "premium_type": None,
        "premium_amount": None,
        "premium_currency": None,
    }
    if fields:
        c516 = fields[0]
        c = c516.split(":")
        if len(c) >= 1:
            d["chd_amt_qual"] = keep_spaces(c[0], width=3)
        if len(c) >= 2:
            d["chd_amount"] = c[1]
        if len(c) >= 3:
            d["chd_currency"] = c[2]

    if len(fields) >= 2 and fields[1]:
        c876 = fields[1]
        c = c876.split(":")
        if len(c) >= 1 and c[0]:
            d["chd_cur_qual"] = c[0]
        if len(c) >= 2:
            d["charge_type"] = c[1]

    c507_idx = None
    for idx, f in enumerate(fields[2:], start=2):
        if f.startswith("CDD:"):
            c507_idx = idx
            break

    if c507_idx is not None:
        c507 = fields[c507_idx].split(":")
        if len(c507) > 1:
            d["due_date"] = c507[1]
        if len(c507) > 2:
            d["due_date_fmt"] = c507[2]

        if len(fields) > c507_idx + 1:
            prem = fields[c507_idx + 1].split(":")
            if len(prem) >= 1:
                d["premium_type"] = prem[0]
            if len(prem) >= 2:
                d["premium_amount"] = prem[1]
            if len(prem) >= 3:
                d["premium_currency"] = prem[2]

    return d


def parse_pol_fields(fields):
    """
    POL may contain:
      - basis of sale (two-digit code, e.g., 59/99) as first field
      - party qualifier (PH/MR)
      - name format + name (e.g., U:SMITH A)
      - trailing product code (short alphanumeric)
    """
    d = {
        "party_qual": None,
        "name_fmt": None,
        "name": None,
        "initials": None,
        "product_code": None,
        "basis_of_sale": None
    }
    if fields and re.fullmatch(r"\d{2}", fields[0] or ""):
        d["basis_of_sale"] = fields[0]

    for i, f in enumerate(fields):
        if f and (f in ("PH", "MR")):  # treat '99' as basis-of-sale, not a party qualifier
            d["party_qual"] = f
            if i + 1 < len(fields):
                namefield = fields[i + 1]
                c = namefield.split(":")
                if len(c) >= 2:
                    d["name_fmt"] = c[0]
                    d["name"] = c[1]
            if len(fields) >= i + 3:
                maybe_prod = fields[-1]
                if maybe_prod and len(maybe_prod) <= 4 and re.match(r"[A-Z0-9]+", maybe_prod):
                    d["product_code"] = maybe_prod
            break
    return d


def parse_rff_fields(fields):
    refs = {}
    for f in fields:
        if ":" in f:
            k, v = f.split(":", 1)
            refs[k] = v
    return refs


def parse_pdt_fields(fields):
    pdt1 = fields[0] if len(fields) >= 1 and fields[0] else ""
    pdt2 = fields[1] if len(fields) >= 2 and fields[1] else ""
    return pdt1, pdt2


def parse_cnt_fields(fields):
    out = {}
    for f in fields:
        if ":" in f:
            k, v = f.split(":", 1)
            out[k] = v
    return out


def lstrip_zeros(s: str) -> str:
    if s is None:
        return ""
    t = s.lstrip("0")
    return t if t != "" else "0"


def extract_unb(text: str):
    """Find and split the UNB envelope even if the file has a preamble header line."""
    segments = parse_edi_segments(text)
    for seg in segments:
        if "UNB+" in seg:
            unb = seg[seg.index("UNB+"):]
            return unb.split("+")
    return None


def convert_lg(text: str) -> pd.DataFrame:
    """
    Convert L&G COMTFR EDI into the exact 37-column flat file that matches your "perfect" output.
    One output row per CHD.
    """
    segments = parse_edi_segments(text)

    meta = {
        "unb_sender": None,
        "unb_recipient": None,
        "unb_datetime": None,
        "unb_control_ref": None,  # Column 0
    }
    current = {
        "batch_seq": None,  # UNH index
        "BGM_PYD": None,    # payment date
        "NAD_BO": None,
        "NAD_IN": None,
        "NAD_PA": None,
        "GIS": None,
        "RFF_POL": None,
        "POL": {},
        "PDT": ("", ""),
        "CNT": {},
    }
    rows = []
    rows_by_message = {}
    ifn_by_policy = {}  # reset per message

    # UNB envelope
    f = extract_unb(text)
    if f and len(f) >= 6:
        meta["unb_sender"] = f[2]
        # recipient shown without leading zeros in the perfect output
        meta["unb_recipient"] = lstrip_zeros(f[3])
        meta["unb_datetime"] = f[4]
        meta["unb_control_ref"] = f[5]  # Reference (col 0)

    for seg in segments:
        tag, fields = tokenise(seg)

        if tag == "UNH":
            current["batch_seq"] = str(fields[0]) if fields else ""
            current.update({
                "BGM_PYD": None, "NAD_BO": None, "NAD_IN": None, "NAD_PA": None, "GIS": None,
                "RFF_POL": None, "POL": {}, "PDT": ("", ""), "CNT": {}
            })
            rows_by_message.setdefault(current["batch_seq"], [])
            ifn_by_policy = {}  # clear per message

        elif tag == "BGM":
            for f2 in fields:
                if f2.startswith("PYD:"):
                    parts = f2.split(":")
                    if len(parts) > 1:
                        current["BGM_PYD"] = parts[1]

        elif tag == "NAD":
            if fields and len(fields) >= 2:
                qual, val = fields[0], fields[1]
                if qual == "BO":
                    current["NAD_BO"] = val
                elif qual == "IN":
                    current["NAD_IN"] = lstrip_zeros(val)  # drop leading zeros to match perfect
                elif qual == "PA":
                    current["NAD_PA"] = val

        elif tag == "GIS":
            if fields:
                current["GIS"] = fields[0]

        elif tag == "RFF":
            refs = parse_rff_fields(fields)
            if "POL" in refs:
                current["RFF_POL"] = refs["POL"]
            if "IFN" in refs and current.get("RFF_POL"):
                # IFN is per-policy, not global
                ifn_by_policy[current["RFF_POL"]] = refs["IFN"]

        elif tag == "POL":
            current["POL"] = parse_pol_fields(fields)

        elif tag == "PDT":
            p1, p2 = parse_pdt_fields(fields)
            # normalise PDT2 '01' -> '1'
            p2 = lstrip_zeros(p2) if p2 else p2
            current["PDT"] = (p1, p2)
            # backfill onto the most recent CHD row in this message
            if rows_by_message.get(current["batch_seq"]):
                last_row_idx = rows_by_message[current["batch_seq"]][-1]
                rows[last_row_idx][23] = current["PDT"][0] or ""
                rows[last_row_idx][24] = current["PDT"][1] or ""

        elif tag == "CNT":
            current["CNT"] = parse_cnt_fields(fields)
            # apply to all rows in this message
            for idx in rows_by_message.get(current["batch_seq"], []):
                rows[idx][25] = current["CNT"].get("CTN", "")
                rows[idx][26] = current["CNT"].get("CAM", "")

        elif tag == "CHD":
            chd = parse_chd_fields(fields)
            basis = (current["POL"].get("basis_of_sale") or "")
            policy = current.get("RFF_POL")
            ifn = ifn_by_policy.get(policy, "")

            # build row (37 columns; exact ordering)
            row = [None] * 37
            row[0] = str(meta["unb_control_ref"] or "")     # Reference
            row[1] = str(current["batch_seq"] or "")        # UNH index
            row[2] = str(current["NAD_BO"] or "")
            row[3] = str(current["NAD_IN"] or "")
            row[4] = str(current["NAD_PA"] or "")
            row[5] = str(current["BGM_PYD"] or "")
            row[6] = str(current["GIS"] or "")
            row[7] = "POL"
            row[8] = str(policy or "")
            row[9] = str(basis)                              # basis-of-sale (e.g., 59/99)
            row[10] = str(current["POL"].get("party_qual") or "")
            row[11] = str(current["POL"].get("name_fmt") or "")
            row[12] = str(current["POL"].get("name") or "")
            row[13] = str(current["POL"].get("initials") or "")
            row[14] = str(current["POL"].get("product_code") or "")
            row[15] = str(chd["chd_amt_qual"] or "")         # keep padded spacing
            row[16] = str(chd["chd_amount"] or "")
            row[17] = str(chd["chd_currency"] or "")
            row[18] = str(chd["chd_cur_qual"] or "N")
            row[19] = str(chd["due_date"] or "")
            # Premium type: strip leading zeros to match perfect (e.g., '01' -> '1', '00' -> '')
            row[20] = (str(chd["premium_type"]) or "")
            row[20] = row[20].lstrip("0") if row[20] != "" else ""
            row[21] = str(chd["premium_amount"] or "")
            row[22] = ""                                     # premium currency: forced blank to match perfect
            row[23] = ""                                     # PDT1 backfilled later
            row[24] = ""                                     # PDT2 backfilled later
            row[25] = ""                                     # CNT CTN backfilled later
            row[26] = ""                                     # CNT CAM backfilled later
            row[27] = ""                                     # UNT count backfilled later
            row[28] = ""                                     # blank
            row[29] = str(meta["unb_recipient"] or "")
            row[30] = ""
            row[31] = ""
            row[32] = ""
            row[33] = str(ifn or "")                         # IFN per policy (including 'tc'/'sm' literals)
            row[34] = str(policy or "")
            row[35] = str(chd.get("charge_type") or "CBS")
            row[36] = "EORM"

            rows.append(row)
            rows_by_message[current["batch_seq"]].append(len(rows) - 1)

        elif tag == "UNT":
            # backfill the segment count to all rows in this UNH message
            if fields and fields[0]:
                unt_count = fields[0]
                for idx in rows_by_message.get(current["batch_seq"], []):
                    rows[idx][27] = unt_count

    return pd.DataFrame(rows)


# -------------------- UI action --------------------
if convert_clicked and uploaded is not None:
    text = uploaded.read().decode("utf-8", errors="ignore")
    if provider == "L&G":
        df = convert_lg(text)
        if df.empty:
            st.warning("No CHD records found. Please check that the file is a valid L&G COMTFR EDI export.")
        else:
            st.success(f"Parsed {len(df)} rows from the file.")
            st.dataframe(df.head(50))
            st.download_button(
                "Download CSV",
                data=df.to_csv(index=False, header=False).encode("utf-8"),
                file_name="LG_clean_output.csv",
                mime="text/csv",
            )
    else:
        st.info("Aviva mapping is not implemented yet. Please choose L&G for now.")
