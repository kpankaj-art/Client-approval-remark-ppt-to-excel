
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from pptx import Presentation
from io import BytesIO
import re
from difflib import SequenceMatcher

st.set_page_config(page_title="Client Remark Extractor", page_icon="📋", layout="wide")

st.title("📋 Client Remark Extractor")
st.caption("Upload Excel + Client PPT → extract client remarks → add them to a new 'Client Remark' column.")

def norm(s):
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s

def digits(s):
    return re.sub(r"\D", "", str(s or ""))

def extract_ppt_slides(ppt_bytes):
    prs = Presentation(BytesIO(ppt_bytes))
    records = []

    for slide_no, slide in enumerate(prs.slides, start=1):
        texts = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:
                t = shape.text.strip()
                if t:
                    texts.append(t)

        full = "\n".join(texts)
        records.append({
            "slide": slide_no,
            "texts": texts,
            "full_text": full
        })

    return records

def extract_label_value(text, labels):
    for label in labels:
        pattern = rf"(?im)^\s*{re.escape(label)}\s*[:\-]\s*(.+?)\s*$"
        m = re.search(pattern, text)
        if m:
            return m.group(1).strip()
    return ""

def extract_ppt_record(rec):
    text = rec["full_text"]

    outlet = extract_label_value(text, ["Outlet Name", "Outlet", "Dealer Name", "Dealer"])
    contact = extract_label_value(text, ["Contact No", "Contact", "Mobile", "Mob", "Phone"])
    media = extract_label_value(text, ["Media Type", "Media", "Type"])
    size = extract_label_value(text, ["Size", "Size (Inches)", "Type And Size"])

    # Common remark labels. The detector deliberately keeps the complete value.
    remark = extract_label_value(
        text,
        ["Remarks", "Remark", "Client Remark", "Client Remarks",
         "Client Comment", "Comment", "Comments", "Observation", "Observations"]
    )

    qty = extract_label_value(text, ["Qty", "Quantity"])

    return {
        "slide": rec["slide"],
        "outlet": outlet,
        "contact": contact,
        "media": media,
        "size": size,
        "qty": qty,
        "remark": remark,
        "raw_text": text
    }

def parse_size(value):
    nums = re.findall(r"\d+(?:\.\d+)?", str(value or ""))
    if len(nums) >= 2:
        return nums[0], nums[1]
    return "", ""

def find_col(columns, candidates):
    normalized = {norm(c): c for c in columns}
    for c in candidates:
        if norm(c) in normalized:
            return normalized[norm(c)]
    for col in columns:
        n = norm(col)
        for c in candidates:
            if norm(c) in n or n in norm(c):
                return col
    return None

def score_match(ppt, row):
    score = 0
    reasons = []

    outlet_col = find_col(row.index, ["Outlet Name", "Dealer Name", "Dealer", "Name"])
    contact_col = find_col(row.index, ["Contact No", "Contact", "Mobile", "Phone"])
    media_col = find_col(row.index, ["Media Type", "Media", "Type"])
    size_col = find_col(row.index, ["Size", "Size (Inches)"])
    w_col = find_col(row.index, ["W", "Width"])
    h_col = find_col(row.index, ["H", "Height"])
    qty_col = find_col(row.index, ["Qty", "Quantity"])

    if outlet_col and ppt["outlet"]:
        a, b = norm(ppt["outlet"]), norm(row[outlet_col])
        if a and b:
            ratio = SequenceMatcher(None, a, b).ratio()
            if ratio >= .95:
                score += 45; reasons.append("outlet exact")
            elif ratio >= .75:
                score += 30; reasons.append("outlet similar")

    if contact_col and ppt["contact"]:
        if digits(ppt["contact"]) and digits(ppt["contact"]) == digits(row[contact_col]):
            score += 30; reasons.append("contact")

    if media_col and ppt["media"]:
        if norm(ppt["media"]) == norm(row[media_col]):
            score += 10; reasons.append("media")

    pw, ph = parse_size(ppt["size"])
    if pw and ph:
        if w_col and h_col:
            if norm(pw) == norm(row[w_col]) and norm(ph) == norm(row[h_col]):
                score += 15; reasons.append("size")
        elif size_col:
            rw, rh = parse_size(row[size_col])
            if pw == rw and ph == rh:
                score += 15; reasons.append("size")

    if qty_col and ppt["qty"]:
        if norm(ppt["qty"]) == norm(row[qty_col]):
            score += 5; reasons.append("qty")

    return score, reasons

def process_excel(excel_bytes, ppt_records):
    xls = pd.ExcelFile(BytesIO(excel_bytes))
    results = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(BytesIO(excel_bytes), sheet_name=sheet)
        if df.empty:
            results.append((sheet, df, []))
            continue

        # Always create a new column; do not overwrite existing REMARK.
        if "Client Remark" not in df.columns:
            df["Client Remark"] = ""
        else:
            df["Client Remark"] = df["Client Remark"].fillna("")

        matches = []

        for ppt in ppt_records:
            best_idx = None
            best_score = -1
            best_reasons = []

            for idx, row in df.iterrows():
                score, reasons = score_match(ppt, row)
                if score > best_score:
                    best_score, best_idx, best_reasons = score, idx, reasons

            if best_idx is not None and best_score >= 45:
                # Remark itself is NOT matched. Only the PPT record is matched to a row.
                df.at[best_idx, "Client Remark"] = ppt["remark"]
                matches.append({
                    "slide": ppt["slide"],
                    "row": int(best_idx) + 2,
                    "score": best_score,
                    "remark": ppt["remark"],
                    "status": "Matched"
                })
            else:
                matches.append({
                    "slide": ppt["slide"],
                    "row": None,
                    "score": best_score,
                    "remark": ppt["remark"],
                    "status": "Review"
                })

        results.append((sheet, df, matches))

    return results

def make_xlsx(results):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet, df, _ in results:
            df.to_excel(writer, sheet_name=str(sheet)[:31], index=False)
    output.seek(0)
    return output.getvalue()

excel_file = st.file_uploader("📊 Upload Excel file", type=["xlsx", "xls"])
ppt_file = st.file_uploader("📑 Upload Client PPT", type=["pptx"])

if excel_file and ppt_file:
    if st.button("🚀 Process Remarks", type="primary", use_container_width=True):
        try:
            with st.spinner("Reading PPT and extracting client remarks..."):
                ppt_records_raw = extract_ppt_slides(ppt_file.getvalue())
                ppt_records = [extract_ppt_record(r) for r in ppt_records_raw]

            with st.spinner("Matching PPT records to Excel rows..."):
                results = process_excel(excel_file.getvalue(), ppt_records)

            total = sum(len(x[2]) for x in results)
            matched = sum(1 for x in results for m in x[2] if m["status"] == "Matched")
            review = total - matched

            st.success(f"Done — {total} PPT records processed, {matched} matched, {review} need review.")

            st.subheader("Extracted client remarks")
            preview = pd.DataFrame([
                {
                    "Slide": r["slide"],
                    "Outlet": r["outlet"],
                    "Media": r["media"],
                    "Size": r["size"],
                    "Client Remark": r["remark"]
                } for r in ppt_records
            ])
            st.dataframe(preview, use_container_width=True)

            if review:
                st.warning("Some slides could not be confidently linked to an Excel row. V1 leaves them untouched instead of putting a remark into a potentially wrong row.")
                review_df = pd.DataFrame([
                    m for _, _, ms in results for m in ms if m["status"] == "Review"
                ])
                st.dataframe(review_df, use_container_width=True)

            output_bytes = make_xlsx(results)
            st.download_button(
                "⬇️ Download Updated Excel",
                data=output_bytes,
                file_name="Updated_Client_Remarks.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)
else:
    st.info("Upload both files to start.")
