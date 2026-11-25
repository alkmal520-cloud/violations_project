import os
import socket
from urllib.parse import urlparse
from dotenv import load_dotenv

load_dotenv()


def _can_reach_host(db_url: str, timeout: float = 2.0) -> bool:
    try:
        parsed = urlparse(db_url)
        host = parsed.hostname
        port = parsed.port or 5432
        if not host:
            return False
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def get_database_url() -> str:
    mode = os.getenv("DB_MODE", "auto").strip().lower()
    cloud_url = os.getenv("SUPABASE_DATABASE_URL", "").strip()
    local_url = os.getenv(
        "LOCAL_DATABASE_URL",
        "postgresql://postgres:YOUR_LOCAL_PASSWORD@localhost:5432/violations_db",
    ).strip()

    if mode == "local":
        return local_url

    if mode == "cloud":
        if not cloud_url:
            raise RuntimeError("SUPABASE_DATABASE_URL is not set while DB_MODE=cloud")
        return cloud_url

    if cloud_url and _can_reach_host(cloud_url, timeout=2.0):
        return cloud_url

    return local_url


DATABASE_URL = get_database_url()