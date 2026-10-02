# Invoice data extraction API

A FastAPI application that accepts an invoice PDF or image, uses Gemini to extract structured data, and creates separate metadata and tabular Excel workbooks.

## Setup

1. Create and activate a virtual environment.
2. Install dependencies: `python -m pip install -r requirements.txt`
3. Add `GEMINI_API_KEY=your_key_here` to `.env`.
4. Start the application: `uvicorn app.main:app --reload`
5. Open `http://127.0.0.1:8000`.

## Project layout

- `app/main.py` — application setup and routes
- `app/config.py` — environment-backed settings
- `app/services/` — PDF/image conversion, Gemini, parsing, validation, Excel, and the pipeline
- `app/templates/` and `app/static/` — upload interface styling
- `data/jobs/` — generated files, created automatically and excluded from Git

## API

- `GET /health` — service health status
- `POST /api/extract` — browser multipart request with an `invoice` PDF/PNG/JPG file. It returns a `job_id` and a tabular preview; JSON is not exposed in the browser download controls.
- `GET /api/jobs/{job_id}/download/metadata` — download the existing comprehensive metadata Excel workbook, including traceability.
- `GET /api/jobs/{job_id}/download/extracted-excel` — download the clean, one-sheet extracted Excel workbook.
- `POST /api/sap/extract` — authenticated SAP multipart upload endpoint. Send `X-API-Key` and the `invoice` file.
- `GET /api/sap/jobs/{job_id}/extracted-data` — authenticated SAP JSON endpoint. It returns one `records` array: every record contains invoice, seller, buyer, total, and one product-line values. Call the SAP upload endpoint first, then use its returned `job_id` here.

Every job retains two internal JSON documents: `metadata.json` and `extracted_data.json`. The extracted workbook is a business-readable Excel document with these logical sheets: `Invoice Details`, `Seller Details`, `Buyer Bill-To`, `Ship-To Details`, `Goods Line Items`, `Tax & Totals`, and `Other Details`. Master data is stored as `Field | Value`; goods are stored as one product per row.

Set `SAP_API_KEY` to a long random secret in the deployed environment and configure the same `X-API-Key` header in the SAP destination. Set `ALLOWED_ORIGINS` only to the Fiori launchpad origin(s), for example `ALLOWED_ORIGINS=["https://fiori.example.com"]`. The local default is `http://127.0.0.1:8000`.

Example SAP request sequence:

```text
POST /api/sap/extract
X-API-Key: <SAP_API_KEY>
Content-Type: multipart/form-data (invoice=<PDF or image>)

GET /api/sap/jobs/<job_id>/extracted-data
X-API-Key: <SAP_API_KEY>
```

For production, set `DEBUG=false`, use a process manager/reverse proxy, and arrange scheduled cleanup of `data/jobs/`.
