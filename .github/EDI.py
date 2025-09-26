python
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="EDI Converter", page_icon="📄", layout="centered")

st.title("EDI Converter")
st.caption("Convert provider EDI files (EDIFACT COMTFR) into a clean CSV.")

provider = st.selectbox("Provider", ["L&G", "Aviva"], index=0, help="L&G is implemented. Aviva is a placeholder for now.")
uploaded = st.file_uploader("Upload the raw EDI text file", type=["txt","edi","dat"], accept_multiple_files=False)
convert_clicked = st.button("Convert!", type="primary", disabled=(uploaded is None))

# -------------------- helpers --------------------
def parse_edi_segments(text: str):
    raw = text.replace("\\r\\n", "\\n").replace("\\r", "\\n")
    raw = re.sub(r"^End-of-Header:\\s*\\n", "", raw, flags=re.MULTILINE)
    segs = [s.strip() for s in raw.split("'") if s.strip()]
    return segs

def tokenise(segment: str):
    parts = segment.split("+")
    return parts[0], parts[1:]

def parse_chd_fields(fields):
    d = {
        "chd_amt_qual": None,
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
        if len(c) >= 1: d["chd_amt_qual"] = c[0].strip()
        if len(c) >= 2: d["chd_amount"] = c[1].strip()
        if len(c) >= 3: d["chd_currency"] = c[2].strip()
    if len(fields) >= 2 and fields[1]:
        c876 = fields[1]
        c = c876.split(":")
        if len(c) >= 1 and c[0]: d["chd_cur_qual"] = c[0].strip()
        if len(c) >= 2: d["charge_type"] = c[1].strip()
    c507_idx = None
    for idx, f in enumerate(fields[2:], start=2):
        if f.startswith("CDD:"):
            c507_idx = idx
            break
    if c507_idx is not None:
        c507 = fields[c507_idx]
        c = c507.split(":")
        if len(c) > 1: d["due_date"] = c[1].strip()
        if len(c) > 2: d["due_date_fmt"] = c[2].strip()
        if len(fields) > c507_idx + 1:
            prem = fields[c507_idx + 1]
            c = prem.split(":")
            if len(c) >= 1: d["premium_type"] = c[0].strip()
            if len(c) >= 2: d["premium_amount"] = c[1].strip()
            if len(c) >= 3: d["premium_currency"] = c[2].strip()
    return d

def parse_pol_fields(fields):
    d = {"party_qual": None, "name_fmt": None, "name": None, "initials": None, "product_code": None, "basis_of_sale": None}
    if fields and re.fullmatch(r"\\d{2}", fields[0] or ""):
        d["basis_of_sale"] = fields[0]
    for i, f in enumerate(fields):
        if f and (f in ("PH", "MR", "99")):
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
    pdt1 = fields[0].strip() if len(fields) >= 1 and fields[0] else None
    pdt2 = fields[1].strip() if len(fields) >= 2 and fields[1] else None
    return pdt1, pdt2

def parse_cnt_fields(fields):
    out = {}
    for f in fields:
        if ":" in f:
            k, v = f.split(":", 1)
            out[k] = v
    return out

def convert_lg(text: str) -> pd.DataFrame:
    segments = parse_edi_segments(text)
    meta = {"ua_content_id": None, "ua_message_id": None, "unb_sender": None, "unb_recipient": None, "unb_datetime": None}
    current = {"batch_seq": None, "BGM_PYD": None, "NAD_BO": None, "NAD_IN": None, "NAD_PA": None, "GIS": None,
               "RFF_POL": None, "RFF_IFN": None, "POL": {}, "PDT": (None, None), "CNT": {}, "UNZ_count": None}
    rows = []
    m = re.search(r"UA-Content-ID:\\s*(\\S+)", text);        meta["ua_content_id"] = m.group(1).strip() if m else None
    m = re.search(r"UA-Message-ID:\\s*(\\S+)", text);        meta["ua_message_id"] = m.group(1).strip() if m else None
    for seg in segments:
        if seg.startswith("UNB+"):
            f = seg.split("+")
            if len(f) >= 5:
                meta["unb_sender"] = f[2]
                meta["unb_recipient"] = f[3]
                meta["unb_datetime"] = f[4]
            break
    for seg in segments:
        tag, fields = tokenise(seg)
        if tag == "UNH":
            if fields:
                try: current["batch_seq"] = int(fields[0])
                except Exception: current["batch_seq"] = fields[0]
            current.update({"BGM_PYD": None, "NAD_BO": None, "NAD_IN": None, "NAD_PA": None, "GIS": None,
                            "RFF_POL": None, "RFF_IFN": None, "POL": {}, "PDT": (None, None), "CNT": {}})
        elif tag == "BGM":
            for f in fields:
                if f.startswith("PYD:"):
                    parts = f.split(":")
                    if len(parts) > 1: current["BGM_PYD"] = parts[1]
        elif tag == "NAD":
            if fields and len(fields) >= 2:
                qual, val = fields[0], fields[1]
                if   qual == "BO": current["NAD_BO"] = val
                elif qual == "IN": current["NAD_IN"] = val
                elif qual == "PA": current["NAD_PA"] = val
        elif tag == "GIS":
            if fields: current["GIS"] = fields[0]
        elif tag == "RFF":
            refs = parse_rff_fields(fields)
            if "POL" in refs: current["RFF_POL"] = refs["POL"]
            if "IFN" in refs: current["RFF_IFN"] = refs["IFN"]
        elif tag == "POL":
            current["POL"] = parse_pol_fields(fields)
        elif tag == "PDT":
            current["PDT"] = parse_pdt_fields(fields)
        elif tag == "CNT":
            current["CNT"] = parse_cnt_fields(fields)
        elif tag == "CHD":
            chd = parse_chd_fields(fields)
            row = [
                meta["ua_message_id"] or meta["ua_content_id"],
                current["batch_seq"],
                current["NAD_BO"],
                current["NAD_IN"],
                current["NAD_PA"],
                current["BGM_PYD"],
                current["GIS"],
                "POL",
                current["RFF_POL"],
                "",
                current["POL"].get("party_qual"),
                current["POL"].get("name_fmt"),
                current["POL"].get("name"),
                current["POL"].get("initials"),
                current["POL"].get("product_code"),
                chd["chd_amt_qual"],
                chd["chd_amount"],
                chd["chd_currency"],
                chd["chd_cur_qual"] or "N",
                chd["due_date"],
                chd["premium_type"],
                chd["premium_amount"],
                chd.get("premium_currency",""),
                (current["PDT"][0] or ""),
                (current["PDT"][1] or ""),
                (current["CNT"].get("CTN") or ""),
                (current["CNT"].get("CAM") or ""),
                "",
                "",
                meta["unb_recipient"],
                "",
                "",
                "",
                current["RFF_IFN"] or "",
                current["RFF_POL"],
                chd.get("charge_type") or "CBS",
                "EORM",
            ]
            rows.append(row)
        elif tag == "UNZ":
            pass
    maxlen = max((len(r) for r in rows), default=0)
    for r in rows: r += [""] * (maxlen - len(r))
    df = pd.DataFrame(rows)
    return df

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
