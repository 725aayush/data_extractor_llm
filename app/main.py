import tempfile
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app.config import get_settings
from app.services.document import IMAGE_EXTENSIONS
from app.services.pipeline import process_document


BASE_DIR = Path(__file__).resolve().parent
settings = get_settings()
app = FastAPI(title="Invoice Extractor", version="1.0.0")
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
templates = Jinja2Templates(directory=BASE_DIR / "templates")
ALLOWED_EXTENSIONS = {".pdf", *IMAGE_EXTENSIONS}
ARTIFACTS = {
    "json": ("extraction.json", "application/json"),
    "excel": ("extraction.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    "raw": ("raw_response.txt", "text/plain"),
}


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse(request, "index.html", {"max_upload_mb": settings.max_upload_mb})


@app.get("/health")
def health():
    return {"status": "ok", "gemini_configured": bool(settings.gemini_api_key.strip())}


@app.post("/api/extract")
async def extract_invoice(invoice: UploadFile = File(...)):
    suffix = Path(invoice.filename or "").suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise HTTPException(415, "Upload a PDF, PNG, JPG, JPEG, or WEBP document.")
    if not settings.gemini_api_key.strip():
        raise HTTPException(503, "GEMINI_API_KEY is not configured on the server.")

    max_bytes = settings.max_upload_mb * 1024 * 1024
    size = 0
    temporary = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    temp_path = Path(temporary.name)
    try:
        while chunk := await invoice.read(1024 * 1024):
            size += len(chunk)
            if size > max_bytes:
                raise HTTPException(413, f"Document exceeds the {settings.max_upload_mb} MB upload limit.")
            temporary.write(chunk)
        temporary.close()
        result = process_document(temp_path, settings)
    except HTTPException:
        raise
    except (ValueError, RuntimeError) as error:
        raise HTTPException(422, str(error)) from error
    finally:
        temporary.close()
        temp_path.unlink(missing_ok=True)
        await invoice.close()

    return {
        "job_id": result.job_id,
        "pages_processed": result.page_count,
        "model": result.model,
        "data": result.data,
        "downloads": {name: f"/api/jobs/{result.job_id}/download/{name}" for name in ARTIFACTS},
    }


@app.get("/api/jobs/{job_id}/download/{artifact}")
def download_artifact(job_id: str, artifact: str):
    if artifact not in ARTIFACTS or not job_id.isalnum() or len(job_id) != 32:
        raise HTTPException(404, "File not found.")
    filename, media_type = ARTIFACTS[artifact]
    path = settings.data_dir / job_id / filename
    if not path.is_file():
        raise HTTPException(404, "File not found.")
    return FileResponse(path, media_type=media_type, filename=filename)
