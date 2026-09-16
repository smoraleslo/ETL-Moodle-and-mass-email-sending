"""
Claves de aplicación SMTP (Gmail) guardadas localmente.

- Caché: %LOCALAPPDATA%/PerfeccionaTEC/smtp_keys.bin, cifrada con DPAPI de Windows
  (solo la puede descifrar el mismo usuario de Windows en el mismo equipo).
- Variables de entorno: se leen al iniciar (ENV_VAR_NAMES), no se escriben.
"""
import os
import json
import ctypes
from ctypes import wintypes
from datetime import datetime
from pathlib import Path

ENV_VAR_NAMES = ("PERFECCIONATEC_SMTP_PASSWORD", "GMAIL_APP_PASSWORD", "SMTP_APP_PASSWORD")

CACHE_PATH = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "PerfeccionaTEC" / "smtp_keys.bin"

SOURCE_CACHE = "guardada"


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _dpapi(data: bytes, protect: bool) -> bytes:
    if os.name != "nt":
        raise OSError("El guardado cifrado de claves solo está disponible en Windows.")
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]

    buf = ctypes.create_string_buffer(data, len(data))
    blob_in = _DataBlob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    blob_out = _DataBlob()
    fn = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
    CRYPTPROTECT_UI_FORBIDDEN = 0x1
    if not fn(ctypes.byref(blob_in), None, None, None, None, CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(blob_out)):
        raise ctypes.WinError()
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(blob_out.pbData, ctypes.c_void_p))


def _same_key(a: str, b: str) -> bool:
    # Gmail muestra la clave con espacios ("abcd efgh ..."); con o sin espacios es la misma
    return a.replace(" ", "") == b.replace(" ", "")


def _read_cache() -> list:
    if not CACHE_PATH.is_file():
        return []
    return json.loads(_dpapi(CACHE_PATH.read_bytes(), protect=False).decode("utf-8"))


def _write_cache(entries: list):
    CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    CACHE_PATH.write_bytes(_dpapi(json.dumps(entries).encode("utf-8"), protect=True))


def find_keys():
    """
    Devuelve (claves, errores). Cada clave: {"password", "source"}.
    Primero las guardadas (la más reciente primero), luego las de variables de entorno.
    """
    keys, errors = [], []
    try:
        for entry in _read_cache():
            keys.append({"password": entry["password"], "source": SOURCE_CACHE})
    except Exception as e:
        errors.append(f"No se pudo leer la caché de claves ({CACHE_PATH}): {e}")

    for name in ENV_VAR_NAMES:
        value = (os.environ.get(name) or "").strip()
        if value and not any(_same_key(value, k["password"]) for k in keys):
            keys.append({"password": value, "source": f"variable {name}"})
    return keys, errors


def save_key(password: str):
    """Guarda (o sube al primer lugar) una clave en la caché cifrada."""
    password = password.strip()
    entries = [e for e in _read_cache() if not _same_key(e["password"], password)]
    entries.insert(0, {"password": password, "saved_at": datetime.now().isoformat(timespec="seconds")})
    _write_cache(entries)


def delete_key(password: str):
    entries = [e for e in _read_cache() if not _same_key(e["password"], password)]
    if entries:
        _write_cache(entries)
    else:
        CACHE_PATH.unlink(missing_ok=True)


def mask(password: str) -> str:
    compact = password.replace(" ", "")
    return f"•••• •••• •••• {compact[-4:]}" if len(compact) > 4 else "••••"
