from functools import lru_cache
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", extra="ignore")

    gemini_api_key: str = ""
    gemini_models: str = "gemini-3.8-flash,gemini-3.7-flash,gemini-3.6-flash,gemini-3.5-flash"
    max_retries_per_model: int = 3
    initial_retry_delay: int = 5
    pdf_dpi: int = 180
    max_upload_mb: int = 20
    data_dir: Path = Path("data/jobs")
    allowed_origins: list[str] = ["http://127.0.0.1:8000"]
    sap_api_key: str = ""

    @property
    def models(self) -> list[str]:
        return [model.strip() for model in self.gemini_models.split(",") if model.strip()]


@lru_cache
def get_settings() -> Settings:
    return Settings()
