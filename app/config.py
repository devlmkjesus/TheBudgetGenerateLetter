import os


def _parse_bool(value: str, default: bool) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


API_HOST = os.getenv("API_HOST", "0.0.0.0")

try:
    API_PORT = int(os.getenv("API_PORT", os.getenv("PORT", "8000")))
except ValueError:
    API_PORT = 8000

API_RELOAD = _parse_bool(os.getenv("API_RELOAD"), False)

_cors = os.getenv("CORS_ORIGINS", "*").strip()
if _cors == "*":
    CORS_ORIGINS = ["*"]
else:
    CORS_ORIGINS = [o.strip() for o in _cors.split(",") if o.strip()]
