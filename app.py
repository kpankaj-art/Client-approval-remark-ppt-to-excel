import streamlit as st
import pandas as pd
from pptx import Presentation
import io
import re

st.set_page_config(page_title="Multi-Key PPT to Excel Matcher", page_icon="🔗", layout="wide")

st.title("🔗 Dynamic Multi-Key PPT & Excel Matcher")
st.write("Analyze PPT and Excel columns dynamically, set up to 4 custom matching criteria, and export an updated Excel file.")

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

def extract_numbers(text):
    numbers = re.findall(r'\b\d{10}\b', str(text))
    return numbers[0] if numbers else ""

def clean_key_val(val):
    """Normalize string for robust matching"""
    if pd.isna(val) or str(val).lower() == 'nan':
        return ""
    val_str = str(val).strip()
    # Check if 10 digit number
    num = extract_numbers(val_str)
    if num:
        return num
    return re.sub(r'[^a-zA-Z0-9]', '', val_str).lower()

def extract_field(pattern, text):
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else ""

def parse_slide_content(full_text):
    text = " ".join(full_text.split())

    split_parts = re.split(r'\b(Address\s*:|Contact\s*No\s*:|Contact\s*:|District\s*:|Size\s*:|Media\s*Type\s*:)\b', text, flags=re.IGNORECASE)
    pure_outlet_name = split_parts[0].strip()
    pure_outlet_name = re.sub(r'^(Outlet Name|Dealer Name|Shop Name|Party Name)\s*:\s*', '', pure_outlet_name, flags=re.IGNORECASE).strip()

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
        st.subheader("🔍 Analyzed Data Preview")
        
        tab1, tab2 = st.tabs(["📄 Uploaded Excel Preview", "🖼️ Extracted PPT Preview"])
        with tab1:
            st.dataframe(df_excel.head(10), use_container_width=True)
        with tab2:
            st.dataframe(df_ppt.head(10), use_container_width=True)

        st.markdown("---")
        st.subheader("🎯 Configure Multi-Key Matching Criteria (Up to 4 Fields)")
        st.caption("Select matching columns between Excel and PPT to create a composite unique matching key.")

        excel_columns = ["-- Ignore --"] + list(df_excel.columns)
        ppt_columns = ["-- Ignore --"] + list(df_ppt.columns)

        col_a, col_b, col_c, col_d = st.columns(4)

        with col_a:
            st.markdown("**Key 1 (Primary)**")
            ex_key1 = st.selectbox("Excel Col 1", options=excel_columns, key="ex1", index=1 if len(excel_columns)>1 else 0)
            ppt_key1 = st.selectbox("PPT Col 1", options=ppt_columns, key="ppt1", index=ppt_columns.index("PPT_Outlet_Name") if "PPT_Outlet_Name" in ppt_columns else 0)

        with col_b:
            st.markdown("**Key 2**")
            ex_key2 = st.selectbox("Excel Col 2", options=excel_columns, key="ex2")
            ppt_key2 = st.selectbox("PPT Col 2", options=ppt_columns, key="ppt2", index=ppt_columns.index("PPT_Contact") if "PPT_Contact" in ppt_columns else 0)

        with col_c:
            st.markdown("**Key 3**")
            ex_key3 = st.selectbox("Excel Col 3", options=excel_columns, key="ex3")
            ppt_key3 = st.selectbox("PPT Col 3", options=ppt_columns, key="ppt3", index=ppt_columns.index("PPT_Size") if "PPT_Size" in ppt_columns else 0)

        with col_d:
            st.markdown("**Key 4**")
            ex_key4 = st.selectbox("Excel Col 4", options=excel_columns, key="ex4")
            ppt_key4 = st.selectbox("PPT Col 4", options=ppt_columns, key="ppt4")

        # Gather Selected Mapping Rules
        mapping_rules = []
        if ex_key1 != "-- Ignore --" and ppt_key1 != "-- Ignore --":
            mapping_rules.append((ex_key1, ppt_key1))
        if ex_key2 != "-- Ignore --" and ppt_key2 != "-- Ignore --":
            mapping_rules.append((ex_key2, ppt_key2))
        if ex_key3 != "-- Ignore --" and ppt_key3 != "-- Ignore --":
            mapping_rules.append((ex_key3, ppt_key3))
        if ex_key4 != "-- Ignore --" and ppt_key4 != "-- Ignore --":
            mapping_rules.append((ex_key4, ppt_key4))

        if not mapping_rules:
            st.warning("⚠️ Please select at least Key 1 to perform matching.")
        else:
            if st.button("▶️ Process & Append PPT Data into Excel", type="primary", use_container_width=True):
                with st.spinner("Creating composite keys and matching records..."):
                    
                    # Generate Composite Excel Key
                    excel_key_series = pd.Series([""] * len(df_excel), index=df_excel.index)
                    for ex_col, _ in mapping_rules:
                        excel_key_series += df_excel[ex_col].apply(clean_key_val) + "_"
                    df_excel['composite_match_key'] = excel_key_series

                    # Generate Composite PPT Key
                    ppt_key_series = pd.Series([""] * len(df_ppt), index=df_ppt.index)
                    for _, ppt_col in mapping_rules:
                        ppt_key_series += df_ppt[ppt_col].apply(clean_key_val) + "_"
                    df_ppt['composite_match_key'] = ppt_key_series

                    # Append PPT fields as NEW columns into Excel
                    ppt_target_cols = [c for c in df_ppt.columns if c not in ['composite_match_key']]
                    
                    for target in ppt_target_cols:
                        mapping_dict = dict(zip(df_ppt['composite_match_key'], df_ppt[target]))
                        new_col_name = f"Matched_{target}" if target in df_excel.columns else target
                        df_excel[new_col_name] = df_excel['composite_match_key'].map(mapping_dict).fillna("Not Found / No Match" if target == "PPT_Status" else "")

                    # Clean up temporary matching key column
                    df_final = df_excel.drop(columns=['composite_match_key'])

                    st.success("✅ Matching successfully completed!")
                    st.subheader("📋 Updated Excel Preview (With New PPT Columns)")
                    st.dataframe(df_final, use_container_width=True)

                    # Export File
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False)
                    processed_data = output.getvalue()

                    st.download_button(
                        label="📥 Download Updated Excel File (.xlsx)",
                        data=processed_data,
                        file_name="Updated_Excel_With_PPT_Data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

    except Exception as e:
        st.error(f"Error processing request: {e}")
