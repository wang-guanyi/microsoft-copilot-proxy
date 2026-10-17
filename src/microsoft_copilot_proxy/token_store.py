from __future__ import annotations

import base64
import json
import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

SUBSTRATE_AUDIENCE_PREFIX = "https://substrate.office.com/"


def token_serialization_type(token: str) -> str | None:
    if not token or any(character.isspace() for character in token):
        return None
    count = len(token.split("."))
    return "jwt" if count == 3 else "jwe" if count == 5 else None


def decode_jwt_payload(token: str) -> dict[str, Any]:
    if token_serialization_type(token) != "jwt":
        raise ValueError("Token is not a three-part compact JWT.")
    payload = token.split(".")[1]
    payload += "=" * (-len(payload) % 4)
    value = json.loads(base64.urlsafe_b64decode(payload))
    if not isinstance(value, dict):
        raise ValueError("JWT payload is not a JSON object.")
    return value


def is_substrate_token_claims(claims: dict[str, Any]) -> bool:
    audience = claims.get("aud", "")
    if isinstance(audience, str):
        return audience.startswith(SUBSTRATE_AUDIENCE_PREFIX)
    if isinstance(audience, list):
        return any(isinstance(item, str) and item.startswith(SUBSTRATE_AUDIENCE_PREFIX)
                   for item in audience)
    return False


def read_env_value(path: Path | str, name: str) -> str | None:
    try:
        text = Path(path).read_text(encoding="utf-8")
    except FileNotFoundError:
        return None
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        key, separator, value = stripped.partition("=")
        if separator and key.strip() == name:
            return _clean_env_value(value)
    return None


def _parse_iso_datetime(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc)


class AccessTokenStore:
    def __init__(self, token: str, env_path: Path | str = ".env") -> None:
        self._token = token
        self._env_path = Path(env_path)
        self._mtime_ns = self._read_mtime()
        self._lock = threading.RLock()

    def get(self) -> str:
        with self._lock:
            self._reload_if_changed()
            return self._token

    def status(self) -> dict[str, Any]:
        token = self.get()
        token_type = token_serialization_type(token)
        if token_type is None:
            return {
                "valid": False, "error": "Access token is not a compact JWT or JWE.",
                "token_type": None, "expires_at": None, "seconds_remaining": None,
            }
        if token_type == "jwe":
            captured_at = _parse_iso_datetime(
                read_env_value(self._env_path, "M365_TOKEN_CAPTURED_AT")
            )
            refresh_due_at = _parse_iso_datetime(
                read_env_value(self._env_path, "M365_TOKEN_REFRESH_DUE_AT")
            )
            now = datetime.now(timezone.utc)
            refresh_remaining = (
                max(0, int((refresh_due_at - now).total_seconds()))
                if refresh_due_at else None
            )
            return {
                "valid": True,
                "error": None,
                "token_type": "jwe",
                "captured_at": captured_at.isoformat() if captured_at else None,
                "refresh_due_at": refresh_due_at.isoformat() if refresh_due_at else None,
                "refresh_seconds_remaining": refresh_remaining,
            }
        try:
            claims = decode_jwt_payload(token)
            if not is_substrate_token_claims(claims):
                raise ValueError("Access token is not a substrate.office.com token.")
            expires_at = int(claims["exp"])
        except Exception as exc:
            return {
                "valid": False, "error": f"Cannot decode access token: {exc}",
                "token_type": "jwt", "expires_at": None, "seconds_remaining": 0,
            }
        remaining = max(0, expires_at - int(time.time()))
        return {
            "valid": remaining > 0,
            "error": None,
            "token_type": "jwt",
            "expires_at": datetime.fromtimestamp(expires_at, tz=timezone.utc).isoformat(),
            "seconds_remaining": remaining,
            "captured_at": None,
            "refresh_due_at": None,
            "refresh_seconds_remaining": None,
        }

    def _reload_if_changed(self) -> None:
        mtime = self._read_mtime()
        if mtime is None or mtime == self._mtime_ns:
            return
        token = read_env_value(self._env_path, "M365_ACCESS_TOKEN")
        if token:
            self._token = token
        self._mtime_ns = mtime

    def _read_mtime(self) -> int | None:
        try:
            return self._env_path.stat().st_mtime_ns
        except FileNotFoundError:
            return None


def _read_env_token(path: Path) -> str | None:
    return read_env_value(path, "M365_ACCESS_TOKEN")


def _clean_env_value(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
        return value[1:-1]
    return value
