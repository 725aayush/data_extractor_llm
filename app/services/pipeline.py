import json
from dataclasses import dataclass
from pathlib import Path
from uuid import uuid4

from app.config import Settings
from app.services.document import document_to_images
from app.services.excel import create_excel
from app.services.gemini import extract_document
from app.services.parser import parse_json
from app.services.validation import validate_invoice


@dataclass
class ExtractionResult:
    job_id: str
    data: dict
    page_count: int
    model: str


def process_document(source: Path, settings: Settings) -> ExtractionResult:
    job_id = uuid4().hex
    output_dir = settings.data_dir / job_id
    output_dir.mkdir(parents=True, exist_ok=False)
    pages = document_to_images(source, settings.pdf_dpi)
    raw_response, model = extract_document(pages, settings)
    raw_path = output_dir / "raw_response.txt"
    raw_path.write_text(raw_response, encoding="utf-8")
    data = validate_invoice(parse_json(raw_response, raw_path))
    (output_dir / "extraction.json").write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    create_excel(data, output_dir / "extraction.xlsx")
    return ExtractionResult(job_id=job_id, data=data, page_count=len(pages), model=model)
