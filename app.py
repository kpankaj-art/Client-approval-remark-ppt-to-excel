import streamlit as st
import pandas as pd
from pptx import Presentation
import io
import re

st.set_page_config(page_title="PPT Extractor & Excel Matcher", page_icon="🏷️", layout="wide")

st.title("🏷️ PPT Extra Tag Extractor & Excel Matcher")
st.write("Extract details from PPT and cleanly match/append them into your Excel file.")

# File Uploaders
uploaded_excel = st.file_uploader(
    "1. Upload Excel / CSV File (.xlsx, .xls, .xlsm, .xlsb, .csv)", 
    type=["xlsx", "xls", "xlsm", "xlsb", "csv"]
)
uploaded_ppt = st.file_uploader("2. Upload PPT File (.pptx)", type=["pptx"])

def load_excel_file(file):
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
    for col in df.columns:
        col_clean = str(col).lower().replace("_", " ").replace(".", " ").strip()
        for kw in keywords:
            if kw in col_clean:
                return col
    return None

def extract_numbers(text):
    numbers = re.findall(r'\b\d{10}\b', str(text))
    return numbers[0] if numbers else ""

def clean_str(val):
    return re.sub(r'[^a-zA-Z0-9]', '', str(val)).lower().strip()

def extract_field(pattern, text):
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else ""

def parse_slide_content(full_text):
    text = " ".join(full_text.split())

    # 1. Strict Outlet Name Clean (Cut everything from 'Address:' onwards)
    split_parts = re.split(r'\b(Address\s*:|Contact\s*No\s*:|Contact\s*:|District\s*:|Size\s*:|Media\s*Type\s*:)\b', text, flags=re.IGNORECASE)
    pure_outlet_name = split_parts[0].strip()
    pure_outlet_name = re.sub(r'^(Outlet Name|Dealer Name|Shop Name|Party Name)\s*:\s*', '', pure_outlet_name, flags=re.IGNORECASE).strip()

    # 2. Extract Individual Fields
    contact_no = extract_numbers(text)
    address = extract_field(r'Address\s*:\s*(.*?)(?=\s*(?:Contact|District|Size|Media|Remarks|Qty|s_no|SAP|$))', text)
    district = extract_field(r'District\s*:\s*([A-Za-z0-9\s\-_]+?)(?=\s*(?:Size|Media|Remarks|Qty|s_no|Contact|Address|SAP|$))', text)
    size = extract_field(r'Size\s*:\s*([0-9X\s]+?)(?=\s*(?:Media|Remarks|Qty|s_no|District|Contact|Address|SAP|$))', text)
    media_type = extract_field(r'Media\s*Type\s*:\s*([A-Za-z0-9\s\-_]+?)(?=\s*(?:Remarks|Qty|s_no|District|Size|Contact|Address|SAP|$))', text)
    sap_code = extract_field(r'SAP\s*(?:Code)?\s*:\s*([A-Za-z0-9]+?)(?=\s*(?:Address|Contact|District|Size|Media|Remarks|$))', text)

    return pure_outlet_name, address, contact_no, district, size, media_type, sap_code

def process_ppt_data(ppt_file):
    prs = Presentation(ppt_file)
    ppt_records = []

    standard_keywords = [
        "address", "contact", "district", "size", "media type", "sap",
        "remarks", "qty", "s_no", "outlet name", "dealer name", "shop name"
    ]

    for idx, slide in enumerate(prs.slides):
        extra_tags = []
        text_blocks = []

        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if not text:
                    continue

                text_lower = text.lower()
                text_blocks.append(text)

                is_standard = any(kw in text_lower for kw in standard_keywords)
                if not is_standard and len(text) <= 35:
                    extra_tags.append(text)

        full_text = " ".join(text_blocks)
        name, addr, contact, dist, sz, media, sap = parse_slide_content(full_text)

        ppt_records.append({
            "Slide_No": idx + 1,
            "PPT_Outlet_Name": name,
            "PPT_Address": addr,
            "PPT_Contact": contact,
            "PPT_District": dist,
            "PPT_Size": sz,
            "PPT_Media_Type": media,
            "PPT_SAP_Code": sap,
            "PPT_Status": " | ".join(extra_tags) if extra_tags else "Pending/None"
        })

    return pd.DataFrame(ppt_records)

# Execution Flow
if uploaded_excel and uploaded_ppt:
    try:
        df_excel = load_excel_file(uploaded_excel)
        df_ppt = process_ppt_data(uploaded_ppt)

        st.markdown("---")
        st.subheader("⚙️ Match Options & Column Selection")
        
        # Auto-detect initial columns
        auto_name = find_column(df_excel, ["dealer name", "shop name", "outlet name", "client name", "party name", "name", "dealer"])
        auto_contact = find_column(df_excel, ["dealer contact", "contact no", "contact", "mobile no", "mobile", "phone", "number"])

        excel_cols = list(df_excel.columns)
        
        col1, col2 = st.columns(2)
        with col1:
            selected_name_col = st.selectbox(
                "Select Excel Column for Name / Dealer Name:", 
                options=excel_cols, 
                index=excel_cols.index(auto_name) if auto_name in excel_cols else 0
            )
        with col2:
            selected_contact_col = st.selectbox(
                "Select Excel Column for Contact / Mobile Number:", 
                options=excel_cols, 
                index=excel_cols.index(auto_contact) if auto_contact in excel_cols else 0
            )

        if st.button("▶️ Start Processing & Matching", type="primary", use_container_width=True):
            with st.spinner("Processing PPT and matching Excel dataset..."):

                st.subheader("📊 Extracted PPT Data (Split into Clean Columns)")
                st.dataframe(df_ppt, use_container_width=True)

                # Matching Keys
                df_excel['clean_name'] = df_excel[selected_name_col].apply(clean_str)
                df_excel['clean_contact'] = df_excel[selected_contact_col].astype(str).str.extract(r'(\d{10})').fillna('')

                df_ppt['clean_name'] = df_ppt['PPT_Outlet_Name'].apply(clean_str)
                df_ppt['clean_contact'] = df_ppt['PPT_Contact'].astype(str)

                # Match logic: Try Contact match first, fallback to Name
                status_by_contact = dict(zip(df_ppt['clean_contact'], df_ppt['PPT_Status']))
                status_by_name = dict(zip(df_ppt['clean_name'], df_ppt['PPT_Status']))

                addr_by_contact = dict(zip(df_ppt['clean_contact'], df_ppt['PPT_Address']))
                dist_by_contact = dict(zip(df_ppt['clean_contact'], df_ppt['PPT_District']))
                size_by_contact = dict(zip(df_ppt['clean_contact'], df_ppt['PPT_Size']))

                # Mapping to Excel
                df_excel['PPT_Status'] = df_excel['clean_contact'].map(status_by_contact).fillna(df_excel['clean_name'].map(status_by_name)).fillna("Not Found / No Match")
                df_excel['PPT_Address'] = df_excel['clean_contact'].map(addr_by_contact).fillna("")
                df_excel['PPT_District'] = df_excel['clean_contact'].map(dist_by_contact).fillna("")
                df_excel['PPT_Size'] = df_excel['clean_contact'].map(size_by_contact).fillna("")

                # Drop temporary matching columns
                df_final = df_excel.drop(columns=['clean_name', 'clean_contact'])

                st.success("✅ Matching Completed Successfully!")
                st.subheader("📋 Final Updated Excel Dataset")
                st.dataframe(df_final, use_container_width=True)

                # Export File Download
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

    except Exception as e:
        st.error(f"Error processing files: {e}")
