from __future__ import annotations

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict

AVAILABLE_MODELS = [
    "Magic",
    "Gpt_5_5_Chat",
    "Gpt_5_4_Reasoning",
    "Gpt_5_6_Chat",
    "Gpt_5_6_Reasoning",
    # "Claude_Sonnet",
    # "Claude_Opus",
]


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    access_token: str = Field(default="", alias="M365_ACCESS_TOKEN")
    model_alias: str = Field(default="m365-copilot-auto", alias="M365_MODEL_ALIAS")
