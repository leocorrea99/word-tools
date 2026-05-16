"""Verificação de licença via Supabase."""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Optional, Tuple

import requests

_SUPABASE_URL = "https://pketxqxhahriqpyxmklf.supabase.co"
_SUPABASE_KEY = (
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9"
    ".eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBrZXR4cXhoYWhyaXFweXhta2xmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg5NTA1ODIsImV4cCI6MjA5NDUyNjU4Mn0"
    ".Ims9ZPNZibCju5H2abb6e13qRCM5HMLI15jjAMD2MnE"
)
_HEADERS = {
    "apikey": _SUPABASE_KEY,
    "Authorization": f"Bearer {_SUPABASE_KEY}",
    "Content-Type": "application/json",
}

if sys.platform == "win32":
    import os
    _LICENSE_FILE = Path(os.environ.get("APPDATA", Path.home())) / "WordTools" / "license.json"
else:
    _LICENSE_FILE = Path.home() / ".wordtools" / "license.json"


def get_stored_key() -> Optional[str]:
    try:
        data = json.loads(_LICENSE_FILE.read_text(encoding="utf-8"))
        return data.get("license_key")
    except Exception:
        return None


def save_key(key: str) -> None:
    _LICENSE_FILE.parent.mkdir(parents=True, exist_ok=True)
    _LICENSE_FILE.write_text(
        json.dumps({"license_key": key}, ensure_ascii=False),
        encoding="utf-8",
    )


def validate_key(key: str) -> Tuple[bool, str]:
    """
    Verifica no Supabase se a chave existe e está ativa.
    Retorna (ok, mensagem).
    """
    try:
        resp = requests.get(
            f"{_SUPABASE_URL}/rest/v1/licenses",
            headers=_HEADERS,
            params={"license_key": f"eq.{key}", "select": "active,plan"},
            timeout=10,
        )
        resp.raise_for_status()
        rows = resp.json()
        if not rows:
            return False, "Chave de licença não encontrada."
        row = rows[0]
        if not row.get("active", False):
            return False, "Licença desativada. Entre em contato com o suporte."
        return True, row.get("plan", "free")
    except requests.exceptions.ConnectionError:
        # Sem internet → usa chave local em modo offline
        stored = get_stored_key()
        if stored and stored == key:
            return True, "offline"
        return False, "Sem conexão com a internet e chave não armazenada."
    except Exception as exc:
        return False, f"Erro ao validar licença: {exc}"
