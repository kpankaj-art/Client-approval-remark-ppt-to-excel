import pandas as pd
from pptx import Presentation
import io
import re

def parse_dynamic_slide_text(raw_text):
    """
    Parses unstructured slide text dynamically into a dictionary of key-value pairs.
    Handles missing fields gracefully and captures extra fields like SAP Code, GSTIN, etc.
    """
    # Normalize extra whitespaces
    text = " ".join(raw_text.split())
    
    # 1. Extract 10-digit Phone Number dynamically
    contact_match = re.findall(r'\b\d{10}\b', text)
    contact_no = contact_match[0] if contact_match else ""

    # 2. Extract dimensions/size (e.g. 240x48, 240 X 48)
    size_match = re.search(r'\b(\d+\s*[xX]\s*\d+)\b', text)
    size = size_match.group(1).replace(" ", "").upper() if size_match else ""

    # 3. Dynamic Key-Value extraction using Regex
    # Matches patterns like "Key Name: Value"
    kv_pairs = re.findall(r'([A-Za-z0-9\s\_]+?):\s*([^:]+?)(?=\s+[A-Za-z0-9\s\_]+?:|$)', text)
    
    extracted_data = {}
    for key, val in kv_pairs:
        clean_key = key.strip().lower().replace(" ", "_")
        extracted_data[clean_key] = val.strip()

    # 4. Extract Pure Outlet/Dealer Name (Everything before the first known field colon)
    first_colon_index = text.find(":")
    if first_colon_index != -1:
        # Get text before the first key
        prefix_text = text[:first_colon_index]
        # Remove trailing key label if present
        outlet_name = re.sub(r'([A-Za-z0-9\s\_]+)$', '', prefix_text).strip()
    else:
        outlet_name = text.strip()

    # Clean up standard prefixes from outlet name
    outlet_name = re.sub(r'^(Outlet Name|Dealer Name|Shop Name|Party Name)\s*', '', outlet_name, flags=re.IGNORECASE).strip()

    # Construct clean dictionary with standard fallbacks
    parsed_record = {
        "PPT_Outlet_Name": outlet_name,
        "PPT_Address": extracted_data.get("address", ""),
        "PPT_Contact": contact_no or extracted_data.get("contact_no", extracted_data.get("contact", "")),
        "PPT_District": extracted_data.get("district", ""),
        "PPT_Size": size or extracted_data.get("size", ""),
        "PPT_Media_Type": extracted_data.get("media_type", ""),
        "PPT_SAP_Code": extracted_data.get("sapcode", extracted_data.get("sap_code", "")),
        "PPT_Remarks": extracted_data.get("remarks", "")
    }

    return parsed_record

def process_flexible_ppt(ppt_file_path):
    prs = Presentation(ppt_file_path)
    records = []

    for idx, slide in enumerate(prs.slides):
        slide_text_blocks = []
        status_tags = []

        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if not text:
                    continue
                
                # Identify floating status tags (e.g. approved, pending, rejected)
                if len(text) <= 30 and ":" not in text:
                    status_tags.append(text)
                else:
                    slide_text_blocks.append(text)

        full_text = " ".join(slide_text_blocks)
        parsed = parse_dynamic_slide_text(full_text)
        
        parsed["Slide_No"] = idx + 1
        parsed["PPT_Status"] = " | ".join(status_tags) if status_tags else "Pending/None"
        
        records.append(parsed)

    df = pd.DataFrame(records)
    
    # Reorder key columns nicely
    first_cols = ["Slide_No", "PPT_Outlet_Name", "PPT_Address", "PPT_Contact", "PPT_District", "PPT_Size", "PPT_Media_Type", "PPT_SAP_Code", "PPT_Status"]
    existing_cols = [col for col in first_cols if col in df.columns]
    other_cols = [col for col in df.columns if col not in existing_cols]
    
    return df[existing_cols + other_cols]
