import io
import re
import difflib
import pandas as pd
import streamlit as st
from pptx import Presentation
from openpyxl import load_workbook

st.set_page_config(page_title="PPT → Excel Client Remark Extractor", page_icon="📊", layout="wide")
st.session_state.setdefault("_upload_version", 0)

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
# Automatic Excel column detection
# -----------------------------
# Headings are only a hint. The real detection is based on the values in each
# column, so Excel can use "Name of firm", "Dealer", "Party", "XYZ", etc.
ALIASES = {
    "name": ["dealer/name", "dealer name", "dealer", "outlet name", "outlet",
             "customer name", "customer", "shop name", "shop", "party name",
             "party", "client name", "client", "firm name", "firm",
             "business name", "business", "retailer name", "retailer", "name"],
    "address": ["dealer/address", "dealer address", "dealer adderess",
                "dealer/adderess", "address", "outlet address", "shop address",
                "shop location", "location", "full address", "customer address"],
    "contact": ["mobile no.", "mobile no", "mobile", "contact no", "contact",
                "dealer/contact", "dealer contact", "dealer/contact no",
                "dealer mobile", "phone no", "phone", "mobile number",
                "contact number", "phone number", "telephone"],
    "district": ["district name", "district"],
    "type": ["media type", "media", "type", "media_type"],
    "qty": ["qty", "quantity", "qnty"],
    "width": ["w", "width", "w width", "width (0)", "widht"],
    "height": ["h", "height", "h height", "height (0)", "hight"],
    "sap": ["sap code", "sapcode", "sap", "customer code", "dealer code", "party code"],
}

def header_key(h):
    s = clean_text(h)
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def find_col_by_header(columns, field):
    """Use header only as a secondary tie-breaker."""
    best, score = None, 0.0
    aliases = [header_key(x) for x in ALIASES[field]]
    for c in columns:
        k = header_key(c)
        if not k:
            continue
        for a in aliases:
            r = difflib.SequenceMatcher(None, k, a).ratio()
            if r > score:
                best, score = c, r
    return best if score >= 0.72 else None

def value_samples(series, limit=120):
    vals = []
    for v in series.dropna().head(limit):
        s = str(v).strip()
        if s:
            vals.append(s)
    return vals

def score_column_by_data(series, field):
    """
    Return a 0..1 score based on the actual values in a column.
    This deliberately does not require a particular header.
    """
    vals = value_samples(series)
    if not vals:
        return 0.0

    scores = []

    if field == "contact":
        for v in vals:
            nums = phone_numbers(v)
            scores.append(1.0 if nums else 0.0)
        return sum(scores) / len(scores)

    if field in ("width", "height"):
        nums = [number(v) for v in vals]
        valid = [n for n in nums if n is not None]
        if not valid:
            return 0.0

        # W/H columns are numeric, but so are SR NO, SAP Code and QTY.
        # Distinguish them by realistic dimension ranges and by avoiding
        # columns dominated by 0/1 or very large code numbers.
        numeric_ratio = len(valid) / len(vals)
        medium = sum(1 for n in valid if 2 <= n <= 3000) / len(valid)
        not_code = sum(1 for n in valid if n < 10000) / len(valid)
        varied = 1.0 if len(set(valid)) > 1 else 0.5
        return min(1.0, 0.45 * numeric_ratio + 0.35 * medium + 0.15 * not_code + 0.05 * varied)

    if field == "qty":
        good = 0
        for v in vals:
            n = number(v)
            if n is not None and 0 <= n <= 100 and float(n).is_integer():
                good += 1
        return good / len(vals)

    if field == "type":
        known = {"nl", "nlb", "vsb", "gsb", "sb", "fl", "nonlit",
                 "nonlitboard", "nonlit board", "lit", "glow sign board"}
        for v in vals:
            n = norm_type(v)
            if n in {re.sub(r"[^a-z0-9]", "", x) for x in known}:
                scores.append(1.0)
            elif len(n) <= 12 and any(x in n for x in ("nl", "gsb", "vsb", "sb")):
                scores.append(0.65)
            else:
                scores.append(0.0)
        return sum(scores) / len(scores)

    if field == "district":
        # District values are usually short alphabetic location names.
        for v in vals:
            s = norm_name(v)
            scores.append(1.0 if 2 <= len(s.split()) <= 4 and not any(ch.isdigit() for ch in s) else 0.2)
        return sum(scores) / len(scores)

    if field == "address":
        address_words = {
            "road", "rd", "road.", "near", "village", "vill", "gram", "post",
            "po", "nagar", "market", "main", "mohalla", "street", "st",
            "kanpur", "mau", "lucknow", "district", "dist", "chowk", "modle"
        }
        for v in vals:
            s = norm_address(v)
            tokens = set(s.split())
            hits = len(tokens & address_words)
            # Long text with address-like words is strong evidence.
            scores.append(min(1.0, 0.45 + 0.12 * hits) if (len(tokens) >= 3 and hits >= 1) else 0.0)
        return sum(scores) / len(scores)

    if field == "name":
        for v in vals:
            s = norm_name(v)
            tokens = s.split()
            if not s:
                scores.append(0.0)
                continue
            # Firm/shop names: mostly alphabetic, generally 1-8 words, not long.
            raw = str(v)
            alpha_ratio = sum(ch.isalpha() for ch in raw) / max(1, sum(not ch.isspace() for ch in raw))
            digit_ratio = sum(ch.isdigit() for ch in raw) / max(1, sum(not ch.isspace() for ch in raw))
            bad = len(s) > 90 or len(tokens) > 12 or digit_ratio > 0.35
            scores.append(1.0 if alpha_ratio >= 0.65 and not bad else 0.08)
        return sum(scores) / len(scores)

    return 0.0

def ppt_excel_similarity(field, ppt_value, excel_value):
    """Compare one PPT value against one Excel cell for column discovery."""
    if ppt_value in (None, "") or excel_value in (None, ""):
        return 0.0

    if field == "contact":
        a, b = phone_numbers(ppt_value), phone_numbers(excel_value)
        return 1.0 if set(a) & set(b) else 0.0

    if field == "sap":
        a = re.sub(r"\D", "", str(ppt_value))
        b = re.sub(r"\D", "", str(excel_value))
        return 1.0 if a and b and a == b else 0.0

    if field == "size":
        a, b = parse_size(ppt_value), parse_size(excel_value)
        return 1.0 if a and b and a == b else 0.0

    if field == "width":
        a = parse_size(ppt_value)
        b = number(excel_value)
        return 1.0 if a and b is not None and a[0] == b else 0.0

    if field == "height":
        a = parse_size(ppt_value)
        b = number(excel_value)
        return 1.0 if a and b is not None and a[1] == b else 0.0

    if field == "type":
        return 1.0 if norm_type(ppt_value) == norm_type(excel_value) else 0.0

    if field == "qty":
        a, b = number(ppt_value), number(excel_value)
        return 1.0 if a is not None and b is not None and a == b else 0.0

    if field == "name":
        return similarity(ppt_value, excel_value)

    if field in ("address", "district"):
        a, b = norm_address(ppt_value), norm_address(excel_value)
        if not a or not b:
            return 0.0
        if a == b:
            return 1.0
        return difflib.SequenceMatcher(None, a, b).ratio()

    return 0.0


def column_ppt_score(df, col, field, ppt_records):
    """
    Jointly discover the Excel column by matching its DATA to the actual PPT
    records. This is much safer than guessing from headings alone.
    """
    if not ppt_records:
        return 0.0

    vals = [v for v in df[col].dropna().tolist() if str(v).strip()]
    if not vals:
        return 0.0

    per_record = []
    for rec in ppt_records:
        pv = rec.get(field, "")
        if not pv:
            continue

        best = max(
            (ppt_excel_similarity(field, pv, ev) for ev in vals),
            default=0.0
        )
        per_record.append(best)

    return sum(per_record) / len(per_record) if per_record else 0.0


def infer_dimension_pair(df, ppt_records, used=None):
    """
    Find Width and Height together, not independently.
    This prevents a numeric column such as "Shop Painting allocation Sqfit"
    from being mistaken for Height.
    """
    used = used or set()
    sizes = [parse_size(r.get("size", "")) for r in (ppt_records or [])]
    sizes = [x for x in sizes if x]
    if not sizes:
        return None, None, 0.0

    numeric_cols = []
    for c in df.columns:
        if c in used:
            continue
        vals = [number(v) for v in df[c].dropna().tolist()]
        if vals and sum(v is not None for v in vals) / len(vals) >= 0.70:
            numeric_cols.append(c)

    best = (None, None, 0.0)
    for wc in numeric_cols:
        for hc in numeric_cols:
            if wc == hc:
                continue
            # For each PPT size, check whether ANY Excel row contains both
            # dimensions in the same row.
            hits = 0
            for pw, ph in sizes:
                found = False
                for _, row in df[[wc, hc]].iterrows():
                    ew, eh = number(row[wc]), number(row[hc])
                    if ew is not None and eh is not None and ew == pw and eh == ph:
                        found = True
                        break
                if found:
                    hits += 1
            score = hits / len(sizes)
            if score > best[2]:
                best = (wc, hc, score)

    return best


def infer_columns(df, ppt_records=None):
    """
    Automatically discover Excel fields.

    Priority:
      1) Match the actual Excel DATA against actual PPT records.
      2) Use data-pattern detection as fallback.
      3) Use the Excel heading only as a small tie-breaker.

    Therefore headings such as "Name of firm", "Dealer", "Party", etc. are
    NOT required for the tool to know which data column is the firm name.
    """
    ppt_records = ppt_records or []
    columns = list(df.columns)
    result = {f: None for f in
              ["name", "address", "contact", "district", "type", "qty",
               "width", "height", "sap"]}

    fields = list(result.keys())
    scores = {f: {} for f in fields}

    for field in fields:
        for col in columns:
            joint = column_ppt_score(df, col, field, ppt_records)
            profile = score_column_by_data(df[col], field)
            hcol = find_col_by_header(columns, field)
            header_bonus = 0.025 if col == hcol else 0.0

            # Actual PPT↔Excel agreement is deliberately dominant.
            scores[field][col] = min(
                1.0,
                0.88 * joint + 0.10 * profile + header_bonus
            )

    # If the PPT has SAP codes, SAP is an excellent anchor and should be
    # detected before ordinary text fields.
    order = ["sap", "contact", "name", "address", "width", "height",
             "type", "qty", "district"]

    minimums = {
        "sap": 0.55,
        "contact": 0.55,
        "name": 0.55,
        "address": 0.50,
        "width": 0.55,
        "height": 0.55,
        "type": 0.50,
        "qty": 0.50,
        "district": 0.45,
    }

    used = set()
    for field in order:
        ranked = sorted(
            ((scores[field][c], c) for c in columns if c not in used),
            reverse=True
        )
        if not ranked:
            continue

        best_score, best_col = ranked[0]
        if best_score >= minimums[field]:
            result[field] = best_col
            used.add(best_col)

    # Width + Height must be detected as a PAIR from the same Excel rows.
    # This is stronger than independent numeric-column guessing.
    if any(r.get("size") for r in ppt_records):
        wc, hc, pair_score = infer_dimension_pair(df, ppt_records, used)
        if pair_score >= 0.50:
            if result["width"] is not None:
                used.discard(result["width"])
            if result["height"] is not None:
                used.discard(result["height"])
            result["width"], result["height"] = wc, hc
            used.add(wc)
            used.add(hc)

    # If a field is absent from the PPT, leave it unassigned instead of
    # guessing from unrelated data. It is not needed for matching.
    ppt_field_present = {
        f: any(bool(r.get(f)) for r in ppt_records)
        for f in fields
    }

    # Only use data-pattern fallback for fields that actually exist in PPT.
    for field in fields:
        if result[field] is not None or not ppt_field_present.get(field, False):
            continue

        available = [(score_column_by_data(df[c], field), c)
                     for c in columns if c not in used]
        available.sort(reverse=True)

        if available:
            best_score, best_col = available[0]
            if best_score >= minimums[field]:
                result[field] = best_col
                used.add(best_col)

    return result

# -----------------------------
# PPT extraction
# -----------------------------
FIELD_PATTERNS = {
    "name": r"^\s*(outlet\s*name|dealer\s*name|dealer|customer\s*name|customer|shop\s*name|shop|party\s*name|client\s*name|firm\s*name|business\s*name)\s*[:\-]\s*(.*)$",
    "address": r"^\s*(address|dealer\s*address|outlet\s*address|shop\s*address|location|full\s*address)\s*[:\-]\s*(.*)$",
    "contact": r"^\s*(contact(?:\s*no|\s*number)?|mobile(?:\s*no|\s*number)?|phone(?:\s*no|\s*number)?)\s*[:\-]\s*(.*)$",
    "district": r"^\s*(district)\s*[:\-]\s*(.*)$",
    "size": r"^\s*(size|dimensions?)\s*[:\-]\s*(.*)$",
    "type": r"^\s*(media\s*type|media|type)\s*[:\-]\s*(.*)$",
    "qty": r"^\s*(qty|quantity|qnty)\s*[:\-]\s*(.*)$",
    "sap": r"^\s*(sap\s*code|sapcode|customer\s*code)\s*[:\-]\s*(.*)$",
    "remark": r"^\s*(remarks?|client\s*remarks?|comments?|observation)(?:\s*,?\s*if\s+any)?\s*[:\-]\s*(.*)$",
}

EXCLUDED_LABELS = {
    "qty", "quantity", "size", "media type", "media", "outlet name", "dealer name",
    "address", "contact", "contact no", "contact number", "mobile", "mobile no",
    "sapcode", "sap code", "customer code", "dealer code", "party code", "district", "s_no", "s no", "s.no", "sr no", "sr.no"
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
    rec = {k: "" for k in ["name","address","contact","district","size","type","qty","sap"]}
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
        label = re.sub(r"\s*,?\s*if\s+any$", "", label).strip()
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

    # SAP Code is the strongest available identifier when present in both files.
    if rec.get("sap") and row.get("_sap"):
        a = re.sub(r"\D", "", str(rec["sap"]))
        b = re.sub(r"\D", "", str(row["_sap"]))
        if a and b:
            scores.append(1.0 if a == b else 0.0)
            weights.append(25)

    ns = similarity(rec["name"], row["_name"])
    if rec["name"] and row["_name"]:
        scores.append(ns); weights.append(25)

    cs = contact_score(rec["contact"], row["_contact"])
    if cs is not None:
        scores.append(cs); weights.append(20)

    if rec["address"] and row["_address"]:
        scores.append(similarity(rec["address"], row["_address"])); weights.append(10)

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
            "_sap": r.get(cols["sap"], "") if cols.get("sap") else "",
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

    The PPT contains several pieces of information that are already present in
    the Excel master and therefore MUST NOT become Client Remark, including:
      - SAP Code / SAP CODE Added
      - brand-selection text such as "Bangur, Magna", "SC, Magna", "RC, Magna"
      - all normal record fields
      - s_no / slide-template content

    Actual free-form annotations such as "ok", "Ok GSB", "OK only NLB",
    "NO, Sale is not for GSB" remain candidates for Client Remark.
    """
    result = []

    def add(value):
        value = clean_remark(value)
        if not value:
            return

        low = clean_text(value).lower()
        compact = re.sub(r"[^a-z0-9]+", "", low)

        # Never copy information that is already represented by master fields.
        if re.match(
            r"^\s*(sap\s*code|sapcode|customer\s*code|dealer\s*code|party\s*code)"
            r"\s*[:\-]?\s*.*$",
            value, re.I
        ):
            return

        if re.match(r"^\s*sap\s*code\s*added\s*$", value, re.I):
            return

        # Brand/template selection lines seen in the uploaded PPTs.
        # These are not client remarks and are already represented elsewhere
        # in the master workbook.
        brand_template_patterns = [
            r"^(bangur\s*,?\s*magna)$",
            r"^(sc\s*,?\s*magna)$",
            r"^(rc\s*,?\s*magna)$",
            r"^(sc\s*,?\s*magna\s*)$",
            r"^(bangur\s+magna)$",
            r"^(rc\s+magna)$",
        ]
        if any(re.fullmatch(pat, low, re.I) for pat in brand_template_patterns):
            return

        # Ignore structural labels.
        if compact in {
            "outletname", "dealername", "address", "dealeraddress",
            "contact", "contactno", "contactnumber", "mobile", "mobileno",
            "mobilenumber", "district", "size", "dimensions", "mediatype",
            "media", "type", "qty", "quantity", "remarks", "remark",
            "clientremark", "clientremarks", "comments", "observation",
            "sno", "srno"
        }:
            return

        # Ignore complete structural label/value lines.
        if re.match(
            r"^(outlet name|dealer name|address|dealer address|"
            r"contact(?: no| number)?|mobile(?: no| number)?|district|"
            r"size|dimensions?|media type|media|type|qty|quantity|"
            r"remarks?(?:\s*,?\s*if\s+any)?|client remarks?|comments?|"
            r"observation|s_no|s no|s\.no|sr no|sr\.no)\s*[:\-]",
            value, re.I
        ):
            return

        # GPS/coordinate metadata, slide numbering, and pure numeric strings
        # are not client remarks.
        if re.fullmatch(r"[\d\W_]+", value):
            return

        if len(value) > 180:
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
        r"remarks?(?:\s*,?\s*if\s+any)?|client remarks?|comments?|"
        r"observation|s_no|s no|s\.no|sr no|sr\.no|"
        r"sap\s*code|sapcode|customer\s*code|dealer\s*code|party\s*code)"
        r"\s*[:\-]?",
        re.I
    )

    for line in lines:
        compact_line = clean_text(line).replace(".", "").replace("_", "").replace(" ", "").lower()

        # s_no is always template content.
        if compact_line in {"sno", "srno"}:
            continue

        # All standard fields are not free-form client remarks.
        if field_label_re.match(line):
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
.upload-card {padding:18px;border-radius:14px;border:1px solid #ddd;margin-bottom:12px;color:var(--text-color);}
.excel-card {background:#eefbf1;border:2px solid #49a85a;color:var(--text-color)}
.ppt-card {background:#fff0f0;border:2px solid #d94b4b;color:var(--text-color)}
.small {font-size:13px;color:#667085}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">PPT → Excel Client Remark Extractor</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Upload your Excel master and PowerPoint presentation. The tool automatically maps fields, normalizes different formats, matches records, and adds client remarks.</div>', unsafe_allow_html=True)

c1, c2 = st.columns(2)

with c1:
    st.markdown('<div class="upload-card excel-card"><b>🟢 Excel Master File</b><br><span class="small">Supports .xlsx and .xls</span></div>', unsafe_allow_html=True)
    excel_file = st.file_uploader("Upload Excel", type=["xlsx","xls"], key=f"excel_{st.session_state["_upload_version"]}")

with c2:
    st.markdown('<div class="upload-card ppt-card"><b>🔴 PowerPoint File</b><br><span class="small">Supports .pptx</span></div>', unsafe_allow_html=True)
    ppt_file = st.file_uploader("Upload PowerPoint", type=["pptx"], key=f"ppt_{st.session_state["_upload_version"]}")

if excel_file and ppt_file:
    import hashlib

    current_upload_signature = (
        hashlib.md5(excel_file.getvalue()).hexdigest(),
        hashlib.md5(ppt_file.getvalue()).hexdigest(),
    )

    if st.session_state.get("_upload_signature") != current_upload_signature:
        st.session_state["_upload_signature"] = current_upload_signature
        for _key in ["result", "matches", "unmatched", "cols"]:
            st.session_state.pop(_key, None)

    if st.button("🔄 Replace Files / Start New Analysis", use_container_width=True):
        st.session_state.pop("_upload_signature", None)
        for _key in ["result", "matches", "unmatched", "cols"]:
            st.session_state.pop(_key, None)
        st.session_state["_upload_version"] += 1
        st.rerun()

    if st.button("🚀 Analyze & Match", type="primary", use_container_width=True):
        try:
            with st.spinner("Analyzing Excel and PowerPoint..."):
                df = read_excel(excel_file)

                prs = Presentation(io.BytesIO(ppt_file.getvalue()))
                ppt_records = []
                for _slide_no, _slide in enumerate(prs.slides, start=1):
                    _rec, _explicit, _texts = extract_ppt_record(_slide)
                    if _rec["name"] or _rec["contact"] or _rec["size"] or _rec.get("sap"):
                        ppt_records.append(_rec)

                # Let the actual PPT records teach the Excel column detector.
                cols = infer_columns(df, ppt_records)

                missing = [k for k in ["name","contact"] if not cols[k]]
                if missing:
                    st.error("Could not reliably identify required Excel data columns: " + ", ".join(missing))
                    st.info("The tool now compares the actual PPT data against every Excel column. Check the Detected Excel field mapping below if the file uses an unusual data layout.")
                    st.stop()

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
        st.write("Columns are identified by comparing actual PPT record data with actual Excel column values. Headings are only a small tie-breaker.")
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
