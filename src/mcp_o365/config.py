"""T1 — Configuração e segredos.

Carrega e valida a configuração a partir do ambiente / `.env` com `pydantic-settings`.
Falha-rápido no arranque se faltar um segredo crítico. Os segredos são `SecretStr`,
o que garante que nunca aparecem em `repr`/logs por descuido.
"""

from __future__ import annotations

import base64

from pydantic import Field, SecretStr, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Configuração imutável do servidor, derivada do ambiente."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
        frozen=True,
    )

    # --- Entra ID (Plano B) ---
    entra_tenant_id: str = Field(..., alias="ENTRA_TENANT_ID")
    entra_client_id: str = Field(..., alias="ENTRA_CLIENT_ID")
    entra_client_secret: SecretStr = Field(..., alias="ENTRA_CLIENT_SECRET")
    entra_authority: str = Field(..., alias="ENTRA_AUTHORITY")
    oauth_redirect_uri: str = Field(..., alias="OAUTH_REDIRECT_URI")
    graph_scopes_raw: str = Field(
        "User.Read Mail.Read Mail.Send Mail.ReadWrite Calendars.ReadWrite "
        "People.Read Contacts.Read offline_access openid profile",
        alias="GRAPH_SCOPES",
    )

    # --- Plano A (MCP <-> Claude) ---
    mcp_issuer_url: str = Field(..., alias="MCP_ISSUER_URL")
    mcp_public_base_url: str = Field(..., alias="MCP_PUBLIC_BASE_URL")

    # --- Token store / cifra ---
    token_store_path: str = Field("./tokens.db", alias="TOKEN_STORE_PATH")
    token_encryption_key: SecretStr = Field(..., alias="TOKEN_ENCRYPTION_KEY")

    # --- Aprovação em duas fases ---
    approval_ttl_seconds: int = Field(300, alias="APPROVAL_TTL_SECONDS")

    # --- Runtime ---
    log_level: str = Field("INFO", alias="LOG_LEVEL")
    bind_host: str = Field("127.0.0.1", alias="BIND_HOST")
    bind_port: int = Field(8000, alias="BIND_PORT")

    # --- Fase Aprendizagem (US-L.x) — opt-in, DESLIGADA por defeito (RGPD) ---
    learning_enabled: bool = Field(False, alias="LEARNING_ENABLED")
    learning_min_confidence: float = Field(0.5, alias="LEARNING_MIN_CONFIDENCE")
    learning_top_n: int = Field(3, alias="LEARNING_TOP_N")
    learning_retention_days: int = Field(180, alias="LEARNING_RETENTION_DAYS")
    learning_halflife_days: float = Field(90.0, alias="LEARNING_HALFLIFE_DAYS")

    @property
    def graph_scopes(self) -> list[str]:
        """Scopes Graph como lista (a partir da string separada por espaços)."""
        return [s for s in self.graph_scopes_raw.split() if s]

    def encryption_key_bytes(self) -> bytes:
        """Devolve a chave de cifra decodificada (32 bytes para AES-256)."""
        raw = base64.b64decode(self.token_encryption_key.get_secret_value())
        if len(raw) != 32:
            raise ValueError(
                "TOKEN_ENCRYPTION_KEY deve decodificar para 32 bytes (AES-256)."
            )
        return raw

    @field_validator("entra_authority")
    @classmethod
    def _authority_must_be_https(cls, v: str) -> str:
        if not v.startswith("https://"):
            raise ValueError("ENTRA_AUTHORITY deve ser uma URL https://")
        return v


def load_settings() -> Settings:
    """Ponto único de carregamento. Lança `ValidationError` se faltar segredo crítico."""
    return Settings()  # type: ignore[call-arg]
