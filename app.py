import streamlit as st
import pandas as pd
from pptx import Presentation
from io import BytesIO
import re
from difflib import SequenceMatcher

st.set_page_config(
    page_title="Client Remark Extractor",
    page_icon="📋",
    layout="wide"
)

st.markdown("""
<style>
.main { background-color: #f8fafc; }
.block-container { max-width: 1200px; padding-top: 35px; }
.title {
    text-align: center;
    font-size: 34px;
    font-weight: 700;
    color: #111827;
    margin-bottom: 5px;
}
.subtitle {
    text-align: center;
    color: #667085;
    font-size: 15px;
    margin-bottom: 30px;
}
.upload-card {
    padding: 22px;
    border-radius: 16px;
    min-height: 190px;
    margin-bottom: 15px;
}
.excel-card {
    background: #ecfdf3;
    border: 2px solid #22c55e;
}
.ppt-card {
    background: #fff1f2;
    border: 2px solid #ef4444;
}
.upload-title {
    font-size: 20px;
    font-weight: 700;
    margin-bottom: 5px;
}
.excel-title { color: #15803d; }
.ppt-title { color: #dc2626; }
.upload-description {
    color: #667085;
    font-size: 13px;
    margin-bottom: 15px;
}
.stButton > button {
    width: 100%;
    height: 48px;
    border-radius: 10px;
    font-size: 16px;
    font-weight: 700;
}
</style>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="title">Client Remark Extractor</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'Extract client remarks from PowerPoint and add them to new Excel columns.'
    '</div>',
    unsafe_allow_html=True
)

EXCLUDED_FIELDS = {
    "qty", "quantity", "size", "media type", "media",
    "outlet name", "outlet", "dealer name", "dealer",
    "address", "contact", "contact no", "contact number",
    "mobile", "phone", "sapcode", "sap code", "sap-code",
    "district", "type", "s no", "s.no", "serial no", "serial number"
}

REMARK_LABELS = {
    "remark", "remarks", "client remark", "client remarks",
    "comment", "comments", "client comment", "client comments",
    "observation", "observations"
}

def clean_text(text):
    if text is None:
        return ""
    text = str(text).replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()

def normalize(text):
    text = clean_text(text).lower()
    text = re.sub(r"[^a-z0-9 ]", " ", text)
    return re.sub(r"\s+", " ", text).strip()

def extract_field(lines, possible_labels):
    possible_labels = {normalize(x) for x in possible_labels}

    for line in lines:
        line = clean_text(line)

        match = re.match(
            r"^\s*([^:\-]{2,50})\s*[:\-]\s*(.*?)\s*$",
            line
        )

        if not match:
            continue

        label = normalize(match.group(1))
        value = clean_text(match.group(2))

        if label in possible_labels:
            return value

    return ""

def extract_client_remarks(text):
    lines = [
        clean_text(x)
        for x in text.splitlines()
        if clean_text(x)
    ]

    remarks = []

    # Explicit labels such as:
    # Remarks: OK
    # Client Comment: Please change the board
    for line in lines:
        match = re.match(
            r"^\s*([^:\-]{2,50})\s*[:\-]\s*(.*?)\s*$",
            line
        )

        if not match:
            continue

        label = normalize(match.group(1))
        value = clean_text(match.group(2))

        if label in REMARK_LABELS and value:
            remarks.append(value)

    # Label on one line, remark on following lines
    for i, line in enumerate(lines):
        if normalize(line) in REMARK_LABELS:
            collected = []

            for next_line in lines[i + 1:]:
                match = re.match(
                    r"^\s*([^:\-]{2,50})\s*[:\-]\s*(.*?)\s*$",
                    next_line
                )

                if match:
                    next_label = normalize(match.group(1))

                    if next_label in EXCLUDED_FIELDS:
                        break

                    if next_label in REMARK_LABELS:
                        break

                collected.append(next_line)

            if collected:
                remarks.append(" ".join(collected))

    # Remove duplicates
    final_remarks = []
    seen = set()

    for remark in remarks:
        remark = clean_text(remark)

        if not remark:
            continue

        key = remark.lower()

        if key not in seen:
            seen.add(key)
            final_remarks.append(remark)

    return final_remarks

def extract_ppt_data(ppt_bytes):
    presentation = Presentation(BytesIO(ppt_bytes))
    records = []

    for slide_number, slide in enumerate(
        presentation.slides,
        start=1
    ):
        slide_texts = []

        for shape in slide.shapes:

            # Read table cells
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        text = clean_text(cell.text)
                        if text:
                            slide_texts.append(text)

            # Read normal text boxes
            elif hasattr(shape, "text"):
                text = clean_text(shape.text)

                if text:
                    slide_texts.append(text)

        full_text = "\n".join(slide_texts)

        lines = [
            clean_text(x)
            for x in full_text.splitlines()
            if clean_text(x)
        ]

        outlet = extract_field(
            lines,
            ["Outlet Name", "Outlet", "Dealer Name", "Dealer"]
        )

        contact = extract_field(
            lines,
            ["Contact", "Contact No", "Contact Number",
             "Mobile", "Phone"]
        )

        media = extract_field(
            lines,
            ["Media Type", "Media", "Type"]
        )

        size = extract_field(
            lines,
            ["Size"]
        )

        qty = extract_field(
            lines,
            ["Qty", "Quantity"]
        )

        remarks = extract_client_remarks(full_text)

        records.append({
            "slide": slide_number,
            "outlet": outlet,
            "contact": contact,
            "media": media,
            "size": size,
            "qty": qty,
            "remarks": remarks,
            "raw_text": full_text
        })

    return records

def parse_size(value):
    numbers = re.findall(
        r"\d+(?:\.\d+)?",
        str(value or "")
    )

    if len(numbers) >= 2:
        return numbers[0], numbers[1]

    return "", ""

def find_column(columns, names):
    normalized_columns = {
        normalize(c): c for c in columns
    }

    for name in names:
        key = normalize(name)

        if key in normalized_columns:
            return normalized_columns[key]

    return None

def similarity(a, b):
    a = clean_text(a).lower()
    b = clean_text(b).lower()

    if not a or not b:
        return 0

    return SequenceMatcher(
        None,
        a,
        b
    ).ratio()

def calculate_row_score(ppt_record, row):
    score = 0

    outlet_column = find_column(
        row.index,
        ["Outlet Name", "Dealer Name", "Dealer", "Name"]
    )

    contact_column = find_column(
        row.index,
        ["Contact", "Contact No", "Mobile", "Phone"]
    )

    media_column = find_column(
        row.index,
        ["Media Type", "Media", "Type"]
    )

    size_column = find_column(
        row.index,
        ["Size"]
    )

    width_column = find_column(
        row.index,
        ["W", "Width"]
    )

    height_column = find_column(
        row.index,
        ["H", "Height"]
    )

    # Outlet matching
    if outlet_column and ppt_record["outlet"]:
        ratio = similarity(
            ppt_record["outlet"],
            row[outlet_column]
        )

        if ratio >= 0.92:
            score += 55
        elif ratio >= 0.75:
            score += 35

    # Contact matching
    if contact_column and ppt_record["contact"]:
        ppt_contact = re.sub(
            r"\D",
            "",
            ppt_record["contact"]
        )

        excel_contact = re.sub(
            r"\D",
            "",
            str(row[contact_column])
        )

        if (
            ppt_contact
            and excel_contact
            and ppt_contact == excel_contact
        ):
            score += 35

    # Media matching
    if media_column and ppt_record["media"]:
        if normalize(
            ppt_record["media"]
        ) == normalize(
            row[media_column]
        ):
            score += 10

    # Size matching
    ppt_width, ppt_height = parse_size(
        ppt_record["size"]
    )

    if ppt_width and ppt_height:

        if width_column and height_column:

            excel_width = clean_text(
                row[width_column]
            )

            excel_height = clean_text(
                row[height_column]
            )

            if (
                ppt_width == excel_width
                and
                ppt_height == excel_height
            ):
                score += 15

        elif size_column:

            excel_width, excel_height = parse_size(
                row[size_column]
            )

            if (
                ppt_width == excel_width
                and
                ppt_height == excel_height
            ):
                score += 15

    return score

def process_excel(excel_bytes, ppt_records):
    excel = pd.ExcelFile(
        BytesIO(excel_bytes)
    )

    results = []

    for sheet_name in excel.sheet_names:

        df = pd.read_excel(
            BytesIO(excel_bytes),
            sheet_name=sheet_name
        )

        if df.empty:
            results.append(
                (sheet_name, df, [])
            )
            continue

        # Find maximum number of remarks in any slide
        max_remarks = max(
            [
                len(record["remarks"])
                for record in ppt_records
            ],
            default=1
        )

        # Always create new columns
        for number in range(
            1,
            max_remarks + 1
        ):

            if number == 1:
                column_name = "Client Remark"
            else:
                column_name = f"Client Remark {number}"

            if column_name not in df.columns:
                df[column_name] = ""

        reports = []

        for record in ppt_records:

            best_row = None
            best_score = -1

            # IMPORTANT:
            # Only PPT RECORD is matched to Excel.
            # The remark itself is NEVER used for matching.
            for index, row in df.iterrows():

                current_score = calculate_row_score(
                    record,
                    row
                )

                if current_score > best_score:
                    best_score = current_score
                    best_row = index

            if (
                best_row is not None
                and
                best_score >= 55
            ):

                for number, remark in enumerate(
                    record["remarks"],
                    start=1
                ):

                    if number == 1:
                        column_name = "Client Remark"
                    else:
                        column_name = f"Client Remark {number}"

                    df.at[
                        best_row,
                        column_name
                    ] = remark

                status = "Matched"
                excel_row = int(best_row) + 2

            else:
                status = "Review"
                excel_row = ""

            reports.append({
                "Slide": record["slide"],
                "Excel Row": excel_row,
                "Match Score": best_score,
                "Status": status,
                "Client Remark": (
                    " | ".join(record["remarks"])
                    if record["remarks"]
                    else ""
                )
            })

        results.append(
            (sheet_name, df, reports)
        )

    return results

def create_output_excel(results):
    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        for (
            sheet_name,
            dataframe,
            reports
        ) in results:

            dataframe.to_excel(
                writer,
                sheet_name=str(sheet_name)[:31],
                index=False
            )

    output.seek(0)

    return output.getvalue()

# =========================================================
# UPLOAD UI
# =========================================================

left, right = st.columns(2)

with left:

    st.markdown(
        """
        <div class="upload-card excel-card">
            <div class="upload-title excel-title">
                🟩 Excel File
            </div>
            <div class="upload-description">
                Upload the source Excel workbook.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    excel_file = st.file_uploader(
        "Choose Excel File",
        type=["xlsx", "xls"],
        key="excel_upload"
    )

with right:

    st.markdown(
        """
        <div class="upload-card ppt-card">
            <div class="upload-title ppt-title">
                🟥 PowerPoint File
            </div>
            <div class="upload-description">
                Upload the client PowerPoint presentation.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    ppt_file = st.file_uploader(
        "Choose PowerPoint File",
        type=["pptx"],
        key="ppt_upload"
    )

st.write("")

if excel_file and ppt_file:

    if st.button(
        "PROCESS FILES",
        type="primary",
        use_container_width=True
    ):

        try:

            with st.spinner(
                "Reading PowerPoint and extracting client remarks..."
            ):
                ppt_records = extract_ppt_data(
                    ppt_file.getvalue()
                )

            with st.spinner(
                "Adding client remarks to Excel..."
            ):
                results = process_excel(
                    excel_file.getvalue(),
                    ppt_records
                )

            all_reports = []

            for (
                sheet,
                dataframe,
                reports
            ) in results:

                all_reports.extend(reports)

            total_slides = len(all_reports)

            remarks_found = sum(
                1
                for report in all_reports
                if report["Client Remark"]
            )

            matched = sum(
                1
                for report in all_reports
                if report["Status"] == "Matched"
            )

            col1, col2, col3 = st.columns(3)

            col1.metric(
                "Slides Processed",
                total_slides
            )

            col2.metric(
                "Remarks Found",
                remarks_found
            )

            col3.metric(
                "Rows Matched",
                matched
            )

            st.success(
                "Processing completed successfully."
            )

            st.subheader(
                "Client Remark Preview"
            )

            preview = pd.DataFrame(
                all_reports
            )

            st.dataframe(
                preview,
                use_container_width=True,
                hide_index=True
            )

            review_items = [
                x
                for x in all_reports
                if x["Status"] == "Review"
            ]

            if review_items:

                st.warning(
                    f"{len(review_items)} slide(s) could not be confidently "
                    "linked to an Excel row. They were not written to avoid "
                    "incorrect data."
                )

            output_excel = create_output_excel(
                results
            )

            st.download_button(
                label="DOWNLOAD UPDATED EXCEL",
                data=output_excel,
                file_name="Updated_Client_Remarks.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True
            )

        except Exception as error:

            st.error(
                "An error occurred while processing the files."
            )

            st.exception(error)

else:

    st.info(
        "Upload both an Excel file and a PowerPoint file to start."
    )
