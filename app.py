import streamlit as st
import pandas as pd
from pptx import Presentation
import io
import re

st.set_page_config(page_title="PPT Extra Tag & Excel Matcher", page_icon="🏷️", layout="wide")

st.title("🏷️ Smart PPT Extra Tag Extractor & Excel Matcher")
st.write("Ye tool PPT me image ke pass likhe 'OK', 'Approved' jaise floating tags ko automatic pehchan kar Excel me sahi row me attach kar dega.")

uploaded_excel = st.file_uploader(
    "1. Excel / CSV File Upload Karein (.xlsx, .xls, .xlsm, .xlsb, .csv)", 
    type=["xlsx", "xls", "xlsm", "xlsb", "csv"]
)
uploaded_ppt = st.file_uploader("2. PPT File Upload Karein (.pptx)", type=["pptx"])

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
        col_clean = str(col).lower().replace("_", " ").strip()
        for kw in keywords:
            if kw in col_clean:
                return col
    return None

def extract_numbers(text):
    numbers = re.findall(r'\b\d{10}\b', str(text))
    return numbers[0] if numbers else ""

def clean_str(val):
    return re.sub(r'[^a-zA-Z0-9]', '', str(val)).lower().strip()

def process_ppt_data(ppt_file):
    prs = Presentation(ppt_file)
    ppt_records = []

    standard_labels = [
        "outlet name", "dealer name", "shop name", "address", 
        "contact no", "district", "media type", "size", "qty", "remarks", "s_no"
    ]

    for idx, slide in enumerate(prs.slides):
        contact_no, outlet_name, size_w, size_h = "", "", "", ""
        extra_tags = []

        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if not text:
                    continue

                text_lower = text.lower()

                if "outlet name:" in text_lower or "dealer name:" in text_lower:
                    outlet_name = re.sub(r'^(Outlet Name:|Dealer Name:|Shop Name:)', '', text, flags=re.IGNORECASE).strip()

                elif "contact no:" in text_lower or "contact" in text_lower:
                    found_num = extract_numbers(text)
                    if found_num:
                        contact_no = found_num

                elif "size:" in text_lower:
                    size_text = text.split(":")[-1].strip().upper()
                    size_parts = re.findall(r'\d+', size_text)
                    if len(size_parts) >= 2:
                        size_w, size_h = size_parts[0], size_parts[1]

                else:
                    is_standard = any(label in text_lower for label in standard_labels)
                    if not is_standard and len(text) <= 30:
                        extra_tags.append(text)

        # Fixed: Correct variable name inside generator
        if not contact_no:
            full_text = " ".join([sh.text_frame.text for sh in slide.shapes if sh.has_text_frame])
            found_num = extract_numbers(full_text)
            if found_num:
                contact_no = found_num

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

if uploaded_excel and uploaded_ppt:
    try:
        df_excel = load_excel_file(uploaded_excel)
        df_ppt = process_ppt_data(uploaded_ppt)

        st.subheader("PPT Extracted Data (Extra Status Tags Ke Sath)")
        st.dataframe(df_ppt)

        name_col = find_column(df_excel, ["dealer name", "shop name", "outlet name", "client name", "name"])
        contact_col = find_column(df_excel, ["dealer contact", "contact no", "contact", "mobile no", "phone"])
        width_col = find_column(df_excel, ["width", "w"])
        height_col = find_column(df_excel, ["height", "h"])

        if name_col and contact_col:
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
            df_ppt['clean_w'] = df_ppt['PPT_Width'].astype(str)
            df_ppt['clean_h'] = df_ppt['PPT_Height'].astype(str)

            df_ppt['match_key'] = (
                df_ppt['clean_name'] + "_" + 
                df_ppt['clean_contact'] + "_" + 
                df_ppt['clean_w'] + "_" + 
                df_ppt['clean_h']
            )

            status_dict = dict(zip(df_ppt['match_key'], df_ppt['PPT_Status']))
            df_excel['PPT_Status'] = df_excel['match_key'].map(status_dict).fillna("Not Found / No Match")

            df_final = df_excel.drop(columns=['clean_name', 'clean_contact', 'clean_w', 'clean_h', 'match_key'])

            st.success("✅ Matching Complete! Naya Status Column Update Ho Gaya Hai.")
            st.dataframe(df_final.head(15))

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False)
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Updated Excel Download Karein (.xlsx)",
                data=processed_data,
                file_name="PPT_Matched_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Excel me Name ya Contact wala column nahi mila. Kripya check karein.")

    except Exception as e:
        st.error(f"Error aaya file read karne me: {e}")
