import time

from google import genai
from google.genai import types

from app.config import Settings
from app.services.prompts import JSON_INSTRUCTION, SYSTEM_PROMPT


TEMPORARY_ERROR_MARKERS = ("503", "429", "500", "502", "504", "UNAVAILABLE", "RESOURCE_EXHAUSTED", "DEADLINE_EXCEEDED", "INTERNAL", "TIMEOUT")


def _is_temporary_error(error: Exception) -> bool:
    return any(marker in str(error).upper() for marker in TEMPORARY_ERROR_MARKERS)


def extract_document(pages: list[dict], settings: Settings) -> tuple[str, str]:
    """Ask Gemini to extract all pages, retrying transient failures across models."""
    if not settings.gemini_api_key.strip():
        raise RuntimeError("GEMINI_API_KEY is not configured. Add it to the .env file.")

    contents: list = [SYSTEM_PROMPT, JSON_INSTRUCTION, "Analyze all attached pages together and return the complete JSON object."]
    for page in pages:
        contents.extend([f"===== PAGE {page['page']} =====", page["image"]])

    client = genai.Client(api_key=settings.gemini_api_key)
    last_error: Exception | None = None
    for model in settings.models:
        for attempt in range(1, settings.max_retries_per_model + 1):
            try:
                response = client.models.generate_content(
                    model=model,
                    contents=contents,
                    config=types.GenerateContentConfig(response_mime_type="application/json"),
                )
                if not response or not response.text:
                    raise RuntimeError("Gemini returned an empty response.")
                return response.text, model
            except Exception as error:
                last_error = error
                if not _is_temporary_error(error):
                    raise RuntimeError(f"Gemini request failed: {error}") from error
                if attempt < settings.max_retries_per_model:
                    time.sleep(settings.initial_retry_delay * 2 ** (attempt - 1))
    detail = str(last_error) if last_error else "No error details were returned."
    raise RuntimeError(
        "Gemini was unavailable after trying all configured models. "
        f"Last Gemini error: {detail}"
    ) from last_error
