import json
from dataclasses import dataclass
from pathlib import Path
from uuid import uuid4

from app.config import Settings
from app.services.document import document_to_images
from app.services.artifacts import metadata_data, tabular_data
from app.services.excel import create_excel
from app.services.gemini import extract_document
from app.services.parser import parse_json
from app.services.tabular_excel import create_tabular_excel
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
    metadata = metadata_data(data)
    extracted_data = tabular_data(data)
    (output_dir / "metadata.json").write_text(json.dumps(metadata, indent=2, ensure_ascii=False), encoding="utf-8")
    (output_dir / "extracted_data.json").write_text(json.dumps(extracted_data, indent=2, ensure_ascii=False), encoding="utf-8")
    # Preserve the existing comprehensive workbook as the metadata download.
    # It includes source-page/confidence traceability alongside document facts.
    create_excel(data, output_dir / "metadata.xlsx")
    create_tabular_excel(data, output_dir / "extracted_data.xlsx")
    return ExtractionResult(job_id=job_id, data=data, page_count=len(pages), model=model)
