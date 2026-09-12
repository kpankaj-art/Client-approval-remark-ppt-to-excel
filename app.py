import io
import re
import difflib
import pandas as pd
import streamlit as st
from pptx import Presentation
from openpyxl import load_workbook

st.set_page_config(page_title="PPT → Excel Client Remark Extractor", page_icon="📊", layout="wide")

# -----------------------------
# Normalization / matching
# -----------------------------
def clean_text(v):
    if v is None:
        return ""
    s = str(v).replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def norm_name(v):
    s = clean_text(v)
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def norm_address(v):
    s = clean_text(v)
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def phone_numbers(v):
    if v is None:
        return []
    s = str(v)
    # Split possible multiple contacts, then retain digit sequences.
    nums = []
    for part in re.split(r"[,;/|]+", s):
        digits = re.sub(r"\D", "", part)
        if digits.startswith("91") and len(digits) > 10:
            digits = digits[-10:]
        if len(digits) >= 7:
            nums.append(digits[-10:])
    # Also catch numbers separated only by spaces/hyphens.
    if not nums:
        digits = re.sub(r"\D", "", s)
        if digits.startswith("91") and len(digits) > 10:
            digits = digits[-10:]
        if len(digits) >= 7:
            nums.append(digits[-10:])
    return list(dict.fromkeys(nums))

def parse_size(v):
    if v is None:
        return None
    s = str(v).lower().replace("×", "x").replace("*", "x")
    m = re.search(r"(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)", s)
    if not m:
        return None
    return float(m.group(1)), float(m.group(2))

def number(v):
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return None

def norm_type(v):
    s = clean_text(v)
    s = re.sub(r"[^a-z0-9]", "", s)
    return s

def similarity(a, b):
    a, b = norm_name(a), norm_name(b)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if a in b or b in a:
        return 0.92
    return difflib.SequenceMatcher(None, a, b).ratio()

# -----------------------------
# Semantic Excel header finder
# -----------------------------
ALIASES = {
    "name": ["dealer/name", "dealer name", "outlet name", "outlet", "dealer", "customer name", "shop name", "party name", "name"],
    "address": ["dealer/address", "dealer address", "dealer adderess", "dealer/adderess", "address", "outlet address", "location"],
    "contact": ["mobile no.", "mobile no", "mobile", "contact no", "contact", "dealer/contact", "dealer contact", "dealer/contact no", "phone no", "phone", "mobile number", "contact number"],
    "district": ["district name", "district"],
    "type": ["media type", "media", "type", "media_type"],
    "qty": ["qty", "quantity", "qnty"],
    "width": ["w", "width", "w width", "width (0)"],
    "height": ["h", "height", "h height", "height (0)"],
}

def header_key(h):
    s = clean_text(h)
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def find_col(columns, field):
    cols = list(columns)
    keys = {c: header_key(c) for c in cols}
    aliases = [header_key(x) for x in ALIASES[field]]
    for a in aliases:
        for c, k in keys.items():
            if k == a:
                return c
    # fuzzy fallback
    best, score = None, 0
    for c, k in keys.items():
        if not k:
            continue
        r = difflib.SequenceMatcher(None, k, aliases[0]).ratio()
        if r > score:
            best, score = c, r
    return best if score >= 0.72 else None

def infer_columns(df):
    return {f: find_col(df.columns, f) for f in ALIASES}

# -----------------------------
# PPT extraction
# -----------------------------
FIELD_PATTERNS = {
    "name": r"^\s*(outlet\s*name|dealer\s*name|customer\s*name|shop\s*name)\s*[:\-]\s*(.*)$",
    "address": r"^\s*(address|dealer\s*address|outlet\s*address)\s*[:\-]\s*(.*)$",
    "contact": r"^\s*(contact(?:\s*no|\s*number)?|mobile(?:\s*no|\s*number)?)\s*[:\-]\s*(.*)$",
    "district": r"^\s*(district)\s*[:\-]\s*(.*)$",
    "size": r"^\s*(size|dimensions?)\s*[:\-]\s*(.*)$",
    "type": r"^\s*(media\s*type|media|type)\s*[:\-]\s*(.*)$",
    "qty": r"^\s*(qty|quantity|qnty)\s*[:\-]\s*(.*)$",
    "remark": r"^\s*(remarks?|client\s*remarks?|comments?|observation)\s*[:\-]\s*(.*)$",
}

EXCLUDED_LABELS = {
    "qty", "quantity", "size", "media type", "media", "outlet name", "dealer name",
    "address", "contact", "contact no", "contact number", "mobile", "mobile no",
    "sapcode", "sap code", "district", "s_no", "s no", "s.no", "sr no", "sr.no"
}

def slide_texts(slide):
    """Read text from normal shapes, grouped shapes and tables."""
    texts = []

    def walk(shape):
        # Grouped shapes can contain the actual remark/annotation box.
        if getattr(shape, "shape_type", None) == 6:  # MSO_SHAPE_TYPE.GROUP
            for child in shape.shapes:
                walk(child)
            return

        if getattr(shape, "has_table", False):
            for row in shape.table.rows:
                for cell in row.cells:
                    if cell.text and cell.text.strip():
                        texts.append(cell.text.strip())

        if hasattr(shape, "text") and shape.text and shape.text.strip():
            texts.append(shape.text.strip())

    for shape in slide.shapes:
        walk(shape)

    return texts

def extract_ppt_record(slide):
    texts = slide_texts(slide)
    rec = {k: "" for k in ["name","address","contact","district","size","type","qty"]}
    explicit_remarks = []
    used_lines = set()

    for idx, text in enumerate(texts):
        # A text box may contain multiple lines.
        lines = [x.strip() for x in text.splitlines() if x.strip()]
        for line in lines:
            matched = False
            for field, pattern in FIELD_PATTERNS.items():
                m = re.match(pattern, line, flags=re.I)
                if m:
                    val = m.group(2).strip()
                    if field == "remark":
                        if val:
                            explicit_remarks.append(val)
                    else:
                        rec[field] = val
                    used_lines.add(line)
                    matched = True
                    break
            if not matched:
                # Support "Label :" on one line and value on the next line
                if re.match(r"^\s*(remarks?|client\s*remarks?|comments?|observation)\s*[:\-]?\s*$", line, re.I):
                    if idx + 1 < len(texts):
                        pass

    # Handle labels/values when PowerPoint splits label and value into separate shapes.
    for i, text in enumerate(texts):
        label = clean_text(text).rstrip(":").strip()
        if label in {"outlet name","dealer name","customer name","shop name","address",
                     "dealer address","outlet address","contact no","contact","mobile no",
                     "mobile","district","size","dimensions","media type","media","type",
                     "qty","quantity","remarks","remark","client remark","client remarks",
                     "comments","observation"} and i + 1 < len(texts):
            nxt = texts[i+1].strip()
            if label in {"remarks","remark","client remark","client remarks","comments","observation"}:
                if nxt and not re.match(r"^[A-Za-z ]+\s*:", nxt):
                    explicit_remarks.append(nxt)
            else:
                key = None
                if "name" in label or label in {"outlet","dealer"}: key = "name"
                elif "address" in label: key = "address"
                elif label in {"contact","contact no","mobile","mobile no"}: key = "contact"
                elif label == "district": key = "district"
                elif label in {"size","dimensions"}: key = "size"
                elif label in {"media","media type","type"}: key = "type"
                elif label in {"qty","quantity"}: key = "qty"
                if key and not rec[key]:
                    rec[key] = nxt

    return rec, list(dict.fromkeys(explicit_remarks)), texts


def get_ppt_remarks(explicit):
    """Return only text explicitly written against Remarks/Remark fields."""
    result = []
    for x in explicit:
        x = clean_remark(x)
        if x and x not in result:
            result.append(x)
    return result


# -----------------------------
# Row matching
# -----------------------------
def size_score(ppt_size, row):
    ps = parse_size(ppt_size)
    if not ps:
        return None
    w = number(row.get("_width"))
    h = number(row.get("_height"))
    if w is None or h is None:
        return None
    # 0x0 means unknown; don't penalize a real PPT size.
    if w == 0 and h == 0:
        return None
    return 1.0 if ps == (w, h) else 0.0

def contact_score(ppt_contact, excel_contact):
    a, b = phone_numbers(ppt_contact), phone_numbers(excel_contact)
    if not a or not b:
        return None
    return 1.0 if set(a) & set(b) else 0.0

def match_score(rec, row):
    scores, weights = [], []
    ns = similarity(rec["name"], row["_name"])
    if rec["name"] and row["_name"]:
        scores.append(ns); weights.append(30)

    cs = contact_score(rec["contact"], row["_contact"])
    if cs is not None:
        scores.append(cs); weights.append(20)

    if rec["address"] and row["_address"]:
        scores.append(similarity(rec["address"], row["_address"])); weights.append(15)

    if rec["district"] and row["_district"]:
        scores.append(similarity(rec["district"], row["_district"])); weights.append(10)

    ss = size_score(rec["size"], row)
    if ss is not None:
        scores.append(ss); weights.append(15)

    if rec["type"] and row["_type"]:
        scores.append(1.0 if norm_type(rec["type"]) == norm_type(row["_type"]) else 0.0); weights.append(5)

    pq, eq = number(rec["qty"]), number(row["_qty"])
    if pq is not None and eq is not None:
        scores.append(1.0 if pq == eq else 0.0); weights.append(5)

    if not scores:
        return 0.0
    return sum(s*w for s,w in zip(scores,weights)) / sum(weights)

def prepare_rows(df, cols):
    rows = []
    for idx, r in df.iterrows():
        rows.append({
            "_excel_index": idx,
            "_name": r.get(cols["name"], "") if cols["name"] else "",
            "_address": r.get(cols["address"], "") if cols["address"] else "",
            "_contact": r.get(cols["contact"], "") if cols["contact"] else "",
            "_district": r.get(cols["district"], "") if cols["district"] else "",
            "_width": r.get(cols["width"], "") if cols["width"] else "",
            "_height": r.get(cols["height"], "") if cols["height"] else "",
            "_type": r.get(cols["type"], "") if cols["type"] else "",
            "_qty": r.get(cols["qty"], "") if cols["qty"] else "",
        })
    return rows

# -----------------------------
# Remark extraction
# -----------------------------
def clean_remark(s):
    s = re.sub(r"\s+", " ", str(s)).strip()
    if not s:
        return ""
    return s

def get_remarks(explicit, all_texts):
    """
    Detect additional client remarks/annotations anywhere on the PPT slide.

    Explicit Remarks/Comments/Observation values are stored separately in
    'PPT Remarks'. This function handles other meaningful text.

    IMPORTANT:
    - 's_no' itself is a PPT template field and is ignored.
    - The value beside/after s_no is NOT automatically discarded.
    - Known structural fields are excluded.
    """
    result = []

    def add(value):
        value = clean_remark(value)
        if not value:
            return

        key = clean_text(value).rstrip(":").strip()

        # Ignore structural labels only.
        if key in EXCLUDED_LABELS:
            return

        # Ignore complete structural label lines.
        if re.match(
            r"^(outlet name|dealer name|address|contact(?: no| number)?|"
            r"mobile(?: no| number)?|district|size|media type|qty|quantity|"
            r"remarks?|comments?|observation|s_no|s no|s\.no|sr no|sr\.no)"
            r"\s*[:\-]?\s*$",
            value, re.I
        ):
            return

        if value not in result:
            result.append(value)

    lines = []
    for text in all_texts:
        for x in str(text).splitlines():
            x = x.strip()
            if x:
                lines.append(x)

    field_label_re = re.compile(
        r"^(outlet name|dealer name|address|dealer address|"
        r"contact(?: no| number)?|mobile(?: no| number)?|district|"
        r"size|dimensions?|media type|media|type|qty|quantity|"
        r"remarks?|client remarks?|comments?|observation|"
        r"s_no|s no|s\.no|sr no|sr\.no)\s*[:\-]?",
        re.I
    )

    # Only the s_no label is ignored. Do not ignore the following value.
    for line in lines:
        if clean_text(line).replace(".", "").replace("_", "").replace(" ", "") in {"sno", "srno"}:
            continue
        if field_label_re.match(line):
            continue
        if re.fullmatch(r"[\d\W_]+", line):
            continue
        if len(line) > 180:
            continue
        add(line)

    return result


# -----------------------------
# Excel IO
# -----------------------------
def read_excel(uploaded):
    data = uploaded.getvalue()
    name = uploaded.name.lower()
    if name.endswith(".xlsx"):
        return pd.read_excel(io.BytesIO(data))
    if name.endswith(".xls"):
        # pandas needs xlrd for legacy XLS.
        return pd.read_excel(io.BytesIO(data), engine="xlrd")
    raise ValueError("Please upload .xlsx or .xls")

def build_output(df, matches):
    out = df.copy()
    # Create enough columns for multiple remarks.
    # Dedicated column for the exact PPT 'Remarks:' field.
    if "PPT Remarks" not in out.columns:
        out["PPT Remarks"] = ""

    # Client Remark columns remain separate for additional/unlabelled remarks.
    max_r = max([len(x["remarks"]) for x in matches], default=0)
    ncols = max(1, max_r)
    for i in range(ncols):
        col = "Client Remark" if i == 0 else f"Client Remark {i+1}"
        if col not in out.columns:
            out[col] = ""

    for m in matches:
        idx = m["excel_index"]

        ppt_text = " | ".join(m.get("ppt_remarks", []))
        if ppt_text:
            out.at[idx, "PPT Remarks"] = ppt_text

        for i, remark in enumerate(m["remarks"]):
            col = "Client Remark" if i == 0 else f"Client Remark {i+1}"
            out.at[idx, col] = remark

    return out

# -----------------------------
# UI
# -----------------------------
st.markdown("""
<style>
.main-title {font-size:32px;font-weight:700;margin-bottom:4px}
.subtitle {color:#667085;margin-bottom:22px}
.upload-card {padding:18px;border-radius:14px;border:1px solid #ddd;margin-bottom:12px}
.excel-card {background:#eefbf1;border:2px solid #49a85a}
.ppt-card {background:#fff0f0;border:2px solid #d94b4b}
.small {font-size:13px;color:#667085}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">PPT → Excel Client Remark Extractor</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Upload your Excel master and PowerPoint presentation. The tool automatically maps fields, normalizes different formats, matches records, and adds client remarks.</div>', unsafe_allow_html=True)

c1, c2 = st.columns(2)

with c1:
    st.markdown('<div class="upload-card excel-card"><b>🟢 Excel Master File</b><br><span class="small">Supports .xlsx and .xls</span></div>', unsafe_allow_html=True)
    excel_file = st.file_uploader("Upload Excel", type=["xlsx","xls"], key="excel")

with c2:
    st.markdown('<div class="upload-card ppt-card"><b>🔴 PowerPoint File</b><br><span class="small">Supports .pptx</span></div>', unsafe_allow_html=True)
    ppt_file = st.file_uploader("Upload PowerPoint", type=["pptx"], key="ppt")

if excel_file and ppt_file:
    if st.button("🚀 Analyze & Match", type="primary", use_container_width=True):
        try:
            with st.spinner("Analyzing Excel and PowerPoint..."):
                df = read_excel(excel_file)
                cols = infer_columns(df)
                missing = [k for k in ["name","contact"] if not cols[k]]
                if missing:
                    st.error("Could not identify required Excel columns: " + ", ".join(missing))
                    st.stop()

                prs = Presentation(io.BytesIO(ppt_file.getvalue()))
                rows = prepare_rows(df, cols)
                used = set()
                matches = []
                unmatched = []

                for slide_no, slide in enumerate(prs.slides, start=1):
                    rec, explicit, texts = extract_ppt_record(slide)

                    # Skip cover/title slides that don't have a recognizable record.
                    if not rec["name"] and not rec["contact"] and not rec["size"]:
                        continue

                    candidates = []
                    for row in rows:
                        if row["_excel_index"] in used:
                            continue
                        score = match_score(rec, row)
                        candidates.append((score, row))

                    candidates.sort(key=lambda x: x[0], reverse=True)
                    if not candidates:
                        unmatched.append((slide_no, rec, "No unused Excel row"))
                        continue

                    best_score, best_row = candidates[0]
                    second_score = candidates[1][0] if len(candidates) > 1 else 0

                    # Strong enough, or clearly better than second candidate.
                    confident = best_score >= 0.72 or (best_score >= 0.58 and best_score - second_score >= 0.12)
                    if not confident:
                        unmatched.append((slide_no, rec, f"Low confidence ({best_score:.1f}%)"))
                        continue

                    ppt_remarks = get_ppt_remarks(explicit)
                    remarks = get_remarks(explicit, texts)
                    used.add(best_row["_excel_index"])
                    matches.append({
                        "slide": slide_no,
                        "excel_index": best_row["_excel_index"],
                        "score": best_score,
                        "ppt_remarks": ppt_remarks,
                        "remarks": remarks,
                        "name": rec["name"],
                    })

                output = build_output(df, matches)

                # Preserve Excel as a normal downloadable xlsx.
                buf = io.BytesIO()
                output.to_excel(buf, index=False, engine="openpyxl")
                buf.seek(0)

                st.session_state["result"] = buf.getvalue()
                st.session_state["matches"] = matches
                st.session_state["unmatched"] = unmatched
                st.session_state["cols"] = cols

        except Exception as e:
            st.error(f"Processing failed: {e}")

if "result" in st.session_state:
    matches = st.session_state["matches"]
    unmatched = st.session_state["unmatched"]
    cols = st.session_state["cols"]

    st.success(f"Completed. {len(matches)} PPT record(s) matched successfully.")
    st.info("PPT Remarks is a separate column containing exactly what is written in the PPT Remarks/Remark field. The s_no field and its value are always ignored because they are PPT template content. Additional meaningful unlabelled text may be detected separately as Client Remark.")
    st.download_button(
        "⬇️ Download Updated Excel",
        data=st.session_state["result"],
        file_name="Updated_Client_Remarks.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    with st.expander("🔎 Detected Excel field mapping"):
        st.json(cols)

    with st.expander("📋 Match report"):
        report = pd.DataFrame([
            {"PPT Slide": m["slide"], "Excel Row": m["excel_index"] + 2, "Name": m["name"], "Match Score": f'{m["score"] * 100:.1f}%', "PPT Remarks": " | ".join(m.get("ppt_remarks", [])) or "None", "Other Client Remark": " | ".join(m["remarks"]) or "None"}
            for m in matches
        ])
        if not report.empty:
            st.dataframe(report, use_container_width=True)
        else:
            st.info("No confident matches.")

    with st.expander("⚠️ Review Required"):
        if unmatched:
            for slide_no, rec, reason in unmatched:
                st.write(f"Slide {slide_no}: {rec.get('name') or '(name not detected)'} — {reason}")
        else:
            st.success("No low-confidence/unmatched records.")
