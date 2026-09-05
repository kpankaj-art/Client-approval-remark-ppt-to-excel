import streamlit as st
import pandas as pd
from pptx import Presentation
import io
import re

st.set_page_config(page_title="PPT Tag Extractor & Excel Matcher", page_icon="🏷️", layout="wide")

st.title("🏷️ PPT Extra Tag Extractor & Excel Matcher")
st.write("Extract floating status tags (e.g., 'OK', 'Approved') from PPT slides and match them with your Excel dataset.")

# File Uploaders
uploaded_excel = st.file_uploader(
    "1. Upload Excel / CSV File (.xlsx, .xls, .xlsm, .xlsb, .csv)", 
    type=["xlsx", "xls", "xlsm", "xlsb", "csv"]
)
uploaded_ppt = st.file_uploader("2. Upload PPT File (.pptx)", type=["pptx"])

def load_excel_file(file):
    """Safely loads various Excel and CSV file formats"""
    filename = file.name.lower()
    if filename.endswith('.csv'):
        return pd.read_csv(file)
    elif filename.endswith('.xlsb'):
        return pd.read_excel(file, engine='pyxlsb')
    elif filename.endswith('.xls'):
        try:
            return pd.read_excel(file, engine='xlrd')
        except Exception:
            return pd.read_excel(file)
    else:
        return pd.read_excel(file)

def find_column(df, keywords):
    """Dynamic column finder for Excel sheets"""
    for col in df.columns:
        col_clean = str(col).lower().replace("_", " ").replace(".", " ").strip()
        for kw in keywords:
            if kw in col_clean:
                return col
    return None

def extract_numbers(text):
    """Extracts 10-digit phone number"""
    numbers = re.findall(r'\b\d{10}\b', str(text))
    return numbers[0] if numbers else ""

def clean_str(val):
    return re.sub(r'[^a-zA-Z0-9]', '', str(val)).lower().strip()

def clean_outlet_name(text):
    """Trims Address, Contact No, and District tags from Outlet Name"""
    text = re.sub(r'^(Outlet Name:|Dealer Name:|Shop Name:)\s*', '', text, flags=re.IGNORECASE)
    text = re.split(r'\b(Address:|Contact No:|Contact:|District:|Media Type:)\b', text, flags=re.IGNORECASE)[0]
    return text.strip()

def process_ppt_data(ppt_file):
    prs = Presentation(ppt_file)
    ppt_records = []

    standard_labels = [
        "outlet name", "dealer name", "shop name", "address", 
        "contact no", "contact", "district", "media type", "size", "qty", "remarks", "s_no"
    ]

    for idx, slide in enumerate(prs.slides):
        contact_no, outlet_name, size_w, size_h = "", "", "", ""
        extra_tags = []
        full_slide_text = ""

        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if not text:
                    continue

                full_slide_text += " " + text
                text_lower = text.lower()

                # Extract pure Outlet Name
                if any(lbl in text_lower for lbl in ["outlet name:", "dealer name:", "shop name:", "address:"]):
                    if not outlet_name:
                        outlet_name = clean_outlet_name(text)

                # Extract Contact Number
                if ("contact" in text_lower or "mobile" in text_lower or "phone" in text_lower) and not contact_no:
                    found_num = extract_numbers(text)
                    if found_num:
                        contact_no = found_num

                # Extract Dimensions
                if "size:" in text_lower:
                    size_text = text.split(":")[-1].strip().upper()
                    size_parts = re.findall(r'\d+', size_text)
                    if len(size_parts) >= 2:
                        size_w, size_h = size_parts[0], size_parts[1]

                # Extract Extra Status Tags
                is_standard = any(label in text_lower for label in standard_labels)
                if not is_standard and len(text) <= 30:
                    extra_tags.append(text)

        # Slide-level fallback checks
        if not contact_no:
            found_num = extract_numbers(full_slide_text)
            if found_num:
                contact_no = found_num

        if not outlet_name and full_slide_text:
            lines = [l.strip() for l in full_slide_text.split('\n') if l.strip()]
            if lines:
                outlet_name = clean_outlet_name(lines[0])

        final_status = " | ".join(extra_tags) if extra_tags else "Pending/None"

        ppt_records.append({
            "Slide_No": idx + 1,
            "PPT_Outlet_Name": outlet_name,
            "PPT_Contact": contact_no,
            "PPT_Width": size_w,
            "PPT_Height": size_h,
            "PPT_Status": final_status
        })

    return pd.DataFrame(ppt_records)

# Check if both files are uploaded
if uploaded_excel and uploaded_ppt:
    st.markdown("---")
    # Start Processing Button
    if st.button("▶️ Start Processing & Matching", type="primary", use_container_width=True):
        with st.spinner("Processing PPT and matching data with Excel... Please wait."):
            try:
                df_excel = load_excel_file(uploaded_excel)
                df_ppt = process_ppt_data(uploaded_ppt)

                st.subheader("📊 Extracted PPT Data (Status Tags)")
                st.dataframe(df_ppt, use_container_width=True)

                # Auto-Detect Excel Columns
                name_col = find_column(df_excel, [
                    "dealer name", "shop name", "outlet name", "client name", "party name", "name", "dealer"
                ])
                contact_col = find_column(df_excel, [
                    "dealer contact", "contact no", "contact", "mobile no", "mobile", "phone", "number"
                ])
                width_col = find_column(df_excel, ["width", "w", "size w"])
                height_col = find_column(df_excel, ["height", "h", "size h"])

                if name_col and contact_col:
                    st.info(f"Auto-Detected Excel Columns: **Name** -> '{name_col}' | **Contact** -> '{contact_col}'")

                    # Data Cleaning
                    df_excel['clean_name'] = df_excel[name_col].apply(clean_str)
                    df_excel['clean_contact'] = df_excel[contact_col].astype(str).str.extract(r'(\d{10})').fillna('')
                    df_excel['clean_w'] = df_excel[width_col].astype(str).str.extract(r'(\d+)').fillna('') if width_col else ''
                    df_excel['clean_h'] = df_excel[height_col].astype(str).str.extract(r'(\d+)').fillna('') if height_col else ''

                    # Matching Key Generation
                    df_excel['match_key'] = (
                        df_excel['clean_name'] + "_" + 
                        df_excel['clean_contact'] + "_" + 
                        df_excel['clean_w'] + "_" + 
                        df_excel['clean_h']
                    )

                    df_ppt['clean_name'] = df_ppt['PPT_Outlet_Name'].apply(clean_str)
                    df_ppt['clean_contact'] = df_ppt['PPT_Contact'].astype(str)
                    df_ppt['clean_w'] = df_ppt['PPT_Width'].astype(str)
                    df_ppt['clean_h'] = df_ppt['PPT_Height'].astype(str)

                    df_ppt['match_key'] = (
                        df_ppt['clean_name'] + "_" + 
                        df_ppt['clean_contact'] + "_" + 
                        df_ppt['clean_w'] + "_" + 
                        df_ppt['clean_h']
                    )

                    # Map Status
                    status_dict = dict(zip(df_ppt['match_key'], df_ppt['PPT_Status']))
                    df_excel['PPT_Status'] = df_excel['match_key'].map(status_dict).fillna("Not Found / No Match")

                    # Cleanup Temp Columns
                    df_final = df_excel.drop(columns=['clean_name', 'clean_contact', 'clean_w', 'clean_h', 'match_key'])

                    st.success("✅ Matching Completed Successfully!")
                    st.subheader("📋 Matched Preview")
                    st.dataframe(df_final.head(15), use_container_width=True)

                    # Download Prepared File
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False)
                    processed_data = output.getvalue()

                    st.download_button(
                        label="📥 Download Updated Excel File (.xlsx)",
                        data=processed_data,
                        file_name="PPT_Matched_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.error("Could not auto-detect Name or Contact column in Excel. Please check column headers.")

            except Exception as e:
                st.error(f"Error processing files: {e}")
