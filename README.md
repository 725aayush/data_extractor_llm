# Invoice data extraction API

A FastAPI application that accepts an invoice PDF or image, uses Gemini to extract structured data, and returns JSON and Excel downloads.

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
- `POST /api/extract` — multipart request with an `invoice` PDF/PNG/JPG file
- `GET /api/jobs/{job_id}/download/{artifact}` — download `json`, `excel`, or `raw`

For production, set `DEBUG=false`, use a process manager/reverse proxy, and arrange scheduled cleanup of `data/jobs/`.
