import streamlit as st
import pandas as pd
from pptx import Presentation
import io
import re

st.set_page_config(page_title="PPT Tag Extractor & Excel Matcher", page_icon="🏷️", layout="wide")

st.title("🏷️ Smart PPT Tag Extractor & Excel Splitter")
st.write("Extract and split PPT details into clean individual columns dynamically and match them with your Excel file.")

# File Uploaders
uploaded_excel = st.file_uploader(
    "1. Upload Excel / CSV File (.xlsx, .xls, .xlsm, .xlsb, .csv)", 
    type=["xlsx", "xls", "xlsm", "xlsb", "csv"]
)
uploaded_ppt = st.file_uploader("2. Upload PPT File (.pptx)", type=["pptx"])

def load_excel_file(file):
    """Safely loads various Excel and CSV formats"""
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

def parse_dynamic_text(raw_text):
    """
    Safely parses unstructured text and extracts Outlet Name, Address, Contact, District, Size, Media Type, SAP Code.
    Missing fields remain empty string.
    """
    text = " ".join(raw_text.split())
    
    # 1. Contact Number
    contact_no = extract_numbers(text)

    # 2. Key-Value Regex Extraction (Handles Address, District, Size, Media Type, SAP Code)
    def extract_field(pattern):
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip() if match else ""

    address = extract_field(r'Address:\s*(.*?)(?=\s*(?:Contact|District|Size|Media|SAP|Remarks|Qty|s_no|$))')
    district = extract_field(r'District:\s*([A-Za-z0-9\s\-_]+?)(?=\s*(?:Size|Media|SAP|Remarks|Qty|s_no|Contact|Address|$))')
    size = extract_field(r'Size:\s*([0-9X\s]+?)(?=\s*(?:Media|SAP|Remarks|Qty|s_no|District|Contact|Address|$))')
    media_type = extract_field(r'Media\s*Type:\s*([A-Za-z0-9\s\-_]+?)(?=\s*(?:Remarks|Qty|s_no|District|Size|Contact|Address|$))')
    sap_code = extract_field(r'SAP\s*(?:Code)?:\s*([A-Za-z0-9]+?)(?=\s*(?:Address|Contact|District|Size|Media|Remarks|$))')

    # 3. Pure Outlet Name Extraction
    split_match = re.split(r'\b(Address:|Contact\s*No:|Contact:|District:|Size:|Media\s*Type:|SAP\s*Code:|SAP:)\b', text, flags=re.IGNORECASE)
    outlet_name = split_match[0].strip()
    outlet_name = re.sub(r'^(Outlet Name:|Dealer Name:|Shop Name:)\s*', '', outlet_name, flags=re.IGNORECASE).strip()

    return {
        "PPT_Outlet_Name": outlet_name,
        "PPT_Address": address,
        "PPT_Contact": contact_no,
        "PPT_District": district,
        "PPT_Size": size,
        "PPT_Media_Type": media_type,
        "PPT_SAP_Code": sap_code
    }

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

                # Capture floating status tags
                is_standard = any(kw in text_lower for kw in standard_keywords)
                if not is_standard and len(text) <= 35:
                    extra_tags.append(text)

        full_text = " ".join(text_blocks)
        parsed_fields = parse_dynamic_text(full_text)
        
        parsed_fields["Slide_No"] = idx + 1
        parsed_fields["PPT_Status"] = " | ".join(extra_tags) if extra_tags else "Pending/None"

        ppt_records.append(parsed_fields)

    return pd.DataFrame(ppt_records)

# Execution Flow
if uploaded_excel and uploaded_ppt:
    st.markdown("---")
    if st.button("▶️ Start Processing & Matching", type="primary", use_container_width=True):
        with st.spinner("Processing PPT slides and matching Excel data... Please wait."):
            try:
                df_excel = load_excel_file(uploaded_excel)
                df_ppt = process_ppt_data(uploaded_ppt)

                st.subheader("📊 Extracted PPT Data")
                st.dataframe(df_ppt, use_container_width=True)

                # Column Detection for Excel
                name_col = find_column(df_excel, [
                    "dealer name", "shop name", "outlet name", "client name", "party name", "name", "dealer", "outlet"
                ])
                contact_col = find_column(df_excel, [
                    "dealer contact", "contact no", "contact", "mobile no", "mobile", "phone", "number"
                ])
                width_col = find_column(df_excel, ["width", "w", "size w"])
                height_col = find_column(df_excel, ["height", "h", "size h"])

                if name_col and contact_col:
                    st.info(f"Auto-Detected Excel Columns: **Name** -> '{name_col}' | **Contact** -> '{contact_col}'")

                    # Matching Keys
                    df_excel['clean_name'] = df_excel[name_col].apply(clean_str)
                    df_excel['clean_contact'] = df_excel[contact_col].astype(str).str.extract(r'(\d{10})').fillna('')
                    df_excel['clean_w'] = df_excel[width_col].astype(str).str.extract(r'(\d+)').fillna('') if width_col else ''
                    df_excel['clean_h'] = df_excel[height_col].astype(str).str.extract(r'(\d+)').fillna('') if height_col else ''

                    df_excel['match_key'] = (
                        df_excel['clean_name'] + "_" + 
                        df_excel['clean_contact'] + "_" + 
                        df_excel['clean_w'] + "_" + 
                        df_excel['clean_h']
                    )

                    df_ppt['clean_name'] = df_ppt['PPT_Outlet_Name'].apply(clean_str)
                    df_ppt['clean_contact'] = df_ppt['PPT_Contact'].astype(str)
                    df_ppt['clean_w'] = df_ppt['PPT_Size'].str.extract(r'(\d+)').fillna('')
                    df_ppt['clean_h'] = df_ppt['PPT_Size'].str.extract(r'x(\d+)', flags=re.IGNORECASE).fillna('')

                    df_ppt['match_key'] = (
                        df_ppt['clean_name'] + "_" + 
                        df_ppt['clean_contact'] + "_" + 
                        df_ppt['clean_w'] + "_" + 
                        df_ppt['clean_h']
                    )

                    # Map Split Data into Excel
                    for target_col in ['PPT_Status', 'PPT_Address', 'PPT_District', 'PPT_Size', 'PPT_Media_Type', 'PPT_SAP_Code']:
                        if target_col in df_ppt.columns:
                            col_dict = dict(zip(df_ppt['match_key'], df_ppt[target_col]))
                            df_excel[target_col] = df_excel['match_key'].map(col_dict).fillna("" if target_col != 'PPT_Status' else "Not Found / No Match")

                    # Cleanup Temp Columns
                    df_final = df_excel.drop(columns=['clean_name', 'clean_contact', 'clean_w', 'clean_h', 'match_key'])

                    st.success("✅ Processing and Matching Completed Successfully!")
                    st.subheader("📋 Final Matched Excel Preview")
                    st.dataframe(df_final.head(15), use_container_width=True)

                    # Export to Excel
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
