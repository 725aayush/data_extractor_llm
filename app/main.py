import json
import tempfile
from pathlib import Path

from fastapi import FastAPI, File, Header, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app.config import get_settings
from app.services.document import IMAGE_EXTENSIONS
from app.services.pipeline import process_document


BASE_DIR = Path(__file__).resolve().parent
settings = get_settings()
app = FastAPI(title="Invoice Extractor", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.allowed_origins,
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
templates = Jinja2Templates(directory=BASE_DIR / "templates")
ALLOWED_EXTENSIONS = {".pdf", *IMAGE_EXTENSIONS}
ARTIFACTS = {
    "metadata": ("metadata.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    "extracted-excel": ("extracted_data.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    "raw": ("raw_response.txt", "text/plain"),
}


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse(request, "index.html", {"max_upload_mb": settings.max_upload_mb})


@app.get("/health")
def health():
    return {"status": "ok", "gemini_configured": bool(settings.gemini_api_key.strip())}


async def _extract_invoice(invoice: UploadFile) -> dict:
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
        "tabular_data": _read_tabular_data(result.job_id),
        "downloads": {name: f"/api/jobs/{result.job_id}/download/{name}" for name in ARTIFACTS},
    }


@app.post("/api/extract")
async def extract_invoice(invoice: UploadFile = File(...)):
    """Browser upload endpoint."""
    return await _extract_invoice(invoice)


def _read_tabular_data(job_id: str) -> dict:
    path = settings.data_dir / job_id / "extracted_data.json"
    return json.loads(path.read_text(encoding="utf-8"))


def _require_sap_api_key(x_api_key: str | None) -> None:
    if not settings.sap_api_key.strip():
        raise HTTPException(503, "SAP_API_KEY is not configured on the server.")
    if x_api_key != settings.sap_api_key:
        raise HTTPException(401, "Invalid SAP API key.")


@app.post("/api/sap/extract")
async def sap_extract_invoice(
    invoice: UploadFile = File(...), x_api_key: str | None = Header(default=None)
):
    """Authenticated upload endpoint for a SAP destination or backend service."""
    _require_sap_api_key(x_api_key)
    return await _extract_invoice(invoice)


@app.get("/api/jobs/{job_id}/download/{artifact}")
def download_artifact(job_id: str, artifact: str):
    if artifact not in ARTIFACTS or not job_id.isalnum() or len(job_id) != 32:
        raise HTTPException(404, "File not found.")
    filename, media_type = ARTIFACTS[artifact]
    path = settings.data_dir / job_id / filename
    if not path.is_file():
        raise HTTPException(404, "File not found.")
    return FileResponse(path, media_type=media_type, filename=filename)


def _get_extracted_data(job_id: str) -> dict:
    """Load the internal flat extracted-data JSON after validating the job id."""
    if not job_id.isalnum() or len(job_id) != 32:
        raise HTTPException(404, "Extraction not found.")
    try:
        return _read_tabular_data(job_id)
    except FileNotFoundError as error:
        raise HTTPException(404, "Extraction not found.") from error


@app.get("/api/sap/jobs/{job_id}/extracted-data")
def get_sap_extracted_data(job_id: str, x_api_key: str | None = Header(default=None)):
    """Authenticated SAP endpoint returning flat invoice and line-item records."""
    _require_sap_api_key(x_api_key)
    return _get_extracted_data(job_id)
