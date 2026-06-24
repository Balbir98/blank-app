import re
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Universal EDI to CSV", page_icon="📄", layout="wide")

st.title("Universal EDI to CSV converter")
st.caption("Upload raw OpenText/EDIFACT files, including extensionless files, and convert them to CSV.")


def read_uploaded_file(uploaded_file) -> str:
    data = uploaded_file.read()
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="ignore")


def clean_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"^End-of-Header:\s*\n", "", text, flags=re.IGNORECASE | re.MULTILINE)
    # EDIFACT segments are ended by apostrophe. Newlines are usually transport formatting only.
    return text.replace("\n", "")


def parse_segments(text: str) -> List[str]:
    text = clean_text(text)
    return [seg.strip() for seg in text.split("'") if seg.strip()]


def tokenise(segment: str) -> Tuple[str, List[str]]:
    parts = segment.split("+")
    return parts[0].strip(), parts[1:]


def split_composite(value: Optional[str]) -> List[str]:
    if value is None:
        return []
    return value.split(":")


def strip_leading_zeros(value: Optional[str]) -> str:
    if value is None:
        return ""
    value = str(value)
    stripped = value.lstrip("0")
    return stripped if stripped else ("0" if value else "")


def parse_date(value: Optional[str]) -> str:
    if not value:
        return ""
    value = str(value)
    if re.fullmatch(r"\d{8}", value):
        return f"{value[0:4]}-{value[4:6]}-{value[6:8]}"
    if re.fullmatch(r"\d{6}", value):
        return f"20{value[0:2]}-{value[2:4]}-{value[4:6]}"
    return value


def parse_unb(fields: List[str]) -> Dict[str, str]:
    out = {"syntax": "", "sender": "", "recipient": "", "datetime_raw": "", "control_ref": ""}
    if len(fields) > 0:
        out["syntax"] = fields[0]
    if len(fields) > 1:
        out["sender"] = fields[1]
    if len(fields) > 2:
        out["recipient"] = strip_leading_zeros(fields[2])
    if len(fields) > 3:
        out["datetime_raw"] = fields[3]
    if len(fields) > 4:
        out["control_ref"] = fields[4]
    return out


def universal_raw_csv(text: str, source_file: str) -> pd.DataFrame:
    segments = parse_segments(text)
    rows = []
    unb = {"sender": "", "recipient": "", "datetime_raw": "", "control_ref": ""}
    message_no = ""
    segment_no_in_message = 0
    max_fields = 0

    for absolute_segment_no, segment in enumerate(segments, start=1):
        tag, fields = tokenise(segment)
        max_fields = max(max_fields, len(fields))

        if tag == "UNB":
            unb = parse_unb(fields)
        elif tag == "UNH":
            message_no = fields[0] if fields else ""
            segment_no_in_message = 0

        segment_no_in_message += 1
        row = {
            "source_file": source_file,
            "interchange_ref": unb.get("control_ref", ""),
            "sender": unb.get("sender", ""),
            "recipient": unb.get("recipient", ""),
            "interchange_datetime_raw": unb.get("datetime_raw", ""),
            "message_no": message_no,
            "absolute_segment_no": absolute_segment_no,
            "segment_no_in_message": segment_no_in_message,
            "tag": tag,
            "raw_segment": segment,
        }
        for i, field in enumerate(fields, start=1):
            row[f"field_{i}"] = field
        rows.append(row)

    df = pd.DataFrame(rows)
    for i in range(1, max_fields + 1):
        col = f"field_{i}"
        if col not in df.columns:
            df[col] = ""
    fixed = ["source_file", "interchange_ref", "sender", "recipient", "interchange_datetime_raw", "message_no", "absolute_segment_no", "segment_no_in_message", "tag", "raw_segment"]
    return df[fixed + [f"field_{i}" for i in range(1, max_fields + 1)]].fillna("")


def parse_references(fields: List[str]) -> Dict[str, str]:
    refs = {}
    for field in fields:
        if ":" in field:
            key, val = field.split(":", 1)
            refs[key] = val
    return refs


def parse_name_from_pol(fields: List[str]) -> Dict[str, str]:
    out = {"party_qualifier": "", "name_format": "", "name": "", "surname": "", "forename": "", "basis_of_sale": "", "product_code": ""}
    if fields and re.fullmatch(r"\d{2}", fields[0] or ""):
        out["basis_of_sale"] = fields[0]

    for i, field in enumerate(fields):
        if field in ("PH", "MR", "PA", "IN"):
            out["party_qualifier"] = field
            if i + 1 < len(fields):
                parts = split_composite(fields[i + 1])
                if len(parts) >= 1:
                    out["name_format"] = parts[0]
                if len(parts) >= 2:
                    out["surname"] = parts[1]
                if len(parts) >= 3:
                    out["forename"] = parts[2]
                out["name"] = " ".join([x for x in [out["forename"], out["surname"]] if x]).strip()
                if not out["name"] and len(parts) >= 2:
                    out["name"] = parts[1]
            break

    if fields:
        last = fields[-1]
        if last and re.fullmatch(r"[A-Z0-9]{1,6}", last):
            out["product_code"] = last
    return out


def parse_chd(fields: List[str]) -> Dict[str, str]:
    out = {
        "amount_qualifier": "", "amount": "", "currency": "", "charge_qualifier": "", "charge_type": "",
        "due_date_raw": "", "due_date": "", "due_date_format": "", "premium_type": "", "premium_amount": "", "premium_currency": "",
    }
    if len(fields) >= 1:
        parts = split_composite(fields[0])
        if len(parts) > 0:
            out["amount_qualifier"] = parts[0]
        if len(parts) > 1:
            out["amount"] = parts[1]
        if len(parts) > 2:
            out["currency"] = parts[2]
    if len(fields) >= 2:
        parts = split_composite(fields[1])
        if len(parts) > 0:
            out["charge_qualifier"] = parts[0]
        if len(parts) > 1:
            out["charge_type"] = parts[1]

    for idx, field in enumerate(fields[2:], start=2):
        if field.startswith("CDD:"):
            parts = split_composite(field)
            if len(parts) > 1:
                out["due_date_raw"] = parts[1]
                out["due_date"] = parse_date(parts[1])
            if len(parts) > 2:
                out["due_date_format"] = parts[2]
            if idx + 1 < len(fields):
                prem = split_composite(fields[idx + 1])
                if len(prem) > 0:
                    out["premium_type"] = strip_leading_zeros(prem[0])
                if len(prem) > 1:
                    out["premium_amount"] = prem[1]
                if len(prem) > 2:
                    out["premium_currency"] = prem[2]
            break
    return out


def best_effort_business_csv(text: str, source_file: str) -> pd.DataFrame:
    segments = parse_segments(text)
    rows = []
    unb = {"sender": "", "recipient": "", "datetime_raw": "", "control_ref": ""}
    current = {
        "message_no": "", "payment_date_raw": "", "payment_date": "", "nad_bo": "", "nad_in": "", "nad_pa": "",
        "gis": "", "policy_ref": "", "pol": {}, "cnt_ctn": "", "cnt_cam": "", "unt_count": "",
    }
    rows_in_message = []

    for segment in segments:
        tag, fields = tokenise(segment)
        if tag == "UNB":
            unb = parse_unb(fields)
        elif tag == "UNH":
            current.update({"message_no": fields[0] if fields else "", "payment_date_raw": "", "payment_date": "", "nad_bo": "", "nad_in": "", "nad_pa": "", "gis": "", "policy_ref": "", "pol": {}, "cnt_ctn": "", "cnt_cam": "", "unt_count": ""})
            rows_in_message = []
        elif tag == "BGM":
            for field in fields:
                if field.startswith("PYD:"):
                    parts = split_composite(field)
                    if len(parts) > 1:
                        current["payment_date_raw"] = parts[1]
                        current["payment_date"] = parse_date(parts[1])
        elif tag == "NAD" and len(fields) >= 2:
            if fields[0] == "BO":
                current["nad_bo"] = fields[1]
            elif fields[0] == "IN":
                current["nad_in"] = strip_leading_zeros(fields[1])
            elif fields[0] == "PA":
                current["nad_pa"] = fields[1]
        elif tag == "GIS" and fields:
            current["gis"] = fields[0]
        elif tag == "RFF":
            refs = parse_references(fields)
            if "POL" in refs:
                current["policy_ref"] = refs["POL"]
        elif tag == "POL":
            current["pol"] = parse_name_from_pol(fields)
        elif tag == "CHD":
            chd = parse_chd(fields)
            pol = current.get("pol", {}) or {}
            row = {
                "source_file": source_file,
                "provider_hint": "unknown",
                "interchange_ref": unb.get("control_ref", ""),
                "sender": unb.get("sender", ""),
                "recipient": unb.get("recipient", ""),
                "message_no": current.get("message_no", ""),
                "payment_date": current.get("payment_date", ""),
                "payment_date_raw": current.get("payment_date_raw", ""),
                "nad_bo": current.get("nad_bo", ""),
                "nad_in": current.get("nad_in", ""),
                "nad_pa": current.get("nad_pa", ""),
                "gis": current.get("gis", ""),
                "policy_ref": current.get("policy_ref", ""),
                "basis_of_sale": pol.get("basis_of_sale", ""),
                "party_qualifier": pol.get("party_qualifier", ""),
                "name_format": pol.get("name_format", ""),
                "client_name": pol.get("name", ""),
                "surname": pol.get("surname", ""),
                "forename": pol.get("forename", ""),
                "product_code": pol.get("product_code", ""),
                **chd,
                "cnt_count": "",
                "cnt_amount": "",
                "unt_count": "",
            }
            rows.append(row)
            rows_in_message.append(len(rows) - 1)
        elif tag == "CNT":
            refs = parse_references(fields)
            for idx in rows_in_message:
                rows[idx]["cnt_count"] = refs.get("CTN", "")
                rows[idx]["cnt_amount"] = refs.get("CAM", "")
        elif tag == "UNT" and fields:
            for idx in rows_in_message:
                rows[idx]["unt_count"] = fields[0]

    return pd.DataFrame(rows).fillna("")


output_mode = st.radio(
    "Output mode",
    ["Universal raw parsed CSV", "Best-effort business CSV"],
    help="Universal mode keeps every segment and creates field_1, field_2, etc. Business mode creates one row per CHD line using common COMTFR tags.",
)

uploaded_files = st.file_uploader(
    "Upload one or more raw files",
    type=None,
    accept_multiple_files=True,
    help="Leave file type unrestricted so extensionless OpenText files are accepted.",
)

if uploaded_files:
    frames = []
    for uploaded_file in uploaded_files:
        text = read_uploaded_file(uploaded_file)
        if output_mode == "Universal raw parsed CSV":
            frames.append(universal_raw_csv(text, uploaded_file.name))
        else:
            frames.append(best_effort_business_csv(text, uploaded_file.name))

    df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

    if df.empty:
        st.warning("No data could be parsed from the uploaded file(s).")
    else:
        st.success(f"Converted {len(df)} rows from {len(uploaded_files)} file(s).")
        st.dataframe(df, use_container_width=True)
        output_name = "edi_universal_raw.csv" if output_mode == "Universal raw parsed CSV" else "edi_business_best_effort.csv"
        st.download_button(
            "Download CSV",
            data=df.to_csv(index=False).encode("utf-8"),
            file_name=output_name,
            mime="text/csv",
        )
else:
    st.info("Upload a raw EDI/OpenText file to begin. Extensionless files are accepted.")
