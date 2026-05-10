from pathlib import Path
from pydantic import ConfigDict
from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    app_name: str = "Dify Excel Tool Server"
    api_prefix: str = "/api/v1"
    storage_dir: Path = Path("storage")
    max_upload_mb: int = 20

    model_config = ConfigDict(env_file=".env", env_file_encoding="utf-8")


settings = Settings()
settings.storage_dir.mkdir(parents=True, exist_ok=True)
