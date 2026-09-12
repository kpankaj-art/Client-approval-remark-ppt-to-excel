# Client Remark Extractor — V1

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## What V1 does

- Upload one Excel and one PPTX.
- Reads text from every PPT slide.
- Detects common remark labels such as Remarks, Remark, Client Comment, Observation.
- Matches the PPT record to an Excel row using available record identifiers.
- Does NOT match the remark text itself.
- Creates a NEW `Client Remark` column.
- Leaves uncertain row matches for review rather than risking a wrong Excel update.

## Important V1 limitation

If a client writes a remark only inside an image/photo or handwritten scan, this first version will not reliably OCR it yet. OCR/AI is planned for V2.
