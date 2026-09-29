import json
import re
from pathlib import Path


def parse_json(text: str, raw_response_path: Path) -> dict:
    if not isinstance(text, str):
        raise RuntimeError("Gemini response is not text.")
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip(), flags=re.IGNORECASE).strip()
    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        start, end = cleaned.find("{"), cleaned.rfind("}")
        if start < 0 or end < 0:
            raw_response_path.write_text(text, encoding="utf-8")
            raise RuntimeError("Gemini response does not contain JSON.")
        try:
            data = json.loads(cleaned[start : end + 1])
        except json.JSONDecodeError as error:
            raw_response_path.write_text(text, encoding="utf-8")
            raise RuntimeError("Gemini response was not valid JSON.") from error
    if not isinstance(data, dict):
        raise RuntimeError("Gemini JSON must be a JSON object.")
    return data
