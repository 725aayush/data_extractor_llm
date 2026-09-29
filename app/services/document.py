import io
from pathlib import Path

import pymupdf
from PIL import Image


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp"}


def document_to_images(source: Path, dpi: int) -> list[dict]:
    """Convert a supported PDF or image into Gemini-ready RGB page images."""
    suffix = source.suffix.lower()
    if suffix in IMAGE_EXTENSIONS:
        with Image.open(source) as opened:
            return [{"page": 1, "image": opened.convert("RGB").copy()}]
    if suffix != ".pdf":
        raise ValueError("Only PDF, PNG, JPG, JPEG, and WEBP files are supported.")

    scale = dpi / 72.0
    pages: list[dict] = []
    with pymupdf.open(source) as document:
        for page_number, page in enumerate(document, start=1):
            pixmap = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
            image = Image.open(io.BytesIO(pixmap.tobytes("png"))).convert("RGB")
            pages.append({"page": page_number, "image": image.copy()})
    if not pages:
        raise ValueError("The document has no pages.")
    return pages
