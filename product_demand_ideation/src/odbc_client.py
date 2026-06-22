from __future__ import annotations

import os
import re
from decimal import Decimal
from pathlib import Path
from typing import Any


_LOCAL_ENV_LOADED = False


def _normalize_value(value: Any) -> Any:
    if isinstance(value, Decimal):
        return float(value)
    return value


def _load_local_env() -> None:
    """Load optional local credentials without requiring python-dotenv."""
    global _LOCAL_ENV_LOADED
    if _LOCAL_ENV_LOADED:
        return
    _LOCAL_ENV_LOADED = True

    paths = [
        Path(os.environ["SUNCO_IDEATION_ENV"])
        for key in ["SUNCO_IDEATION_ENV"]
        if os.environ.get(key)
    ]
    paths.append(Path.home() / ".sunco_ideation_development" / ".env")

    for path in paths:
        if not path.exists():
            continue
        for raw_line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value


def _odbc_value(value: str) -> str:
    if any(char in value for char in [";", "{", "}"]):
        return "{" + value.replace("}", "}}") + "}"
    return value


def redshift_connection_string(default: str = "DSN=Redshift") -> str:
    _load_local_env()

    full = os.environ.get("REDSHIFT_DSN") or os.environ.get("REDSHIFT_CONNECTION_STRING")
    if full:
        return full

    dsn = os.environ.get("REDSHIFT_ODBC_DSN") or os.environ.get("ODBC_DSN")
    host = os.environ.get("REDSHIFT_HOST") or os.environ.get("REDSHIFT_SERVER")
    port = os.environ.get("REDSHIFT_PORT") or "5439"
    database = os.environ.get("REDSHIFT_DATABASE") or os.environ.get("REDSHIFT_DB") or "dev"
    user = os.environ.get("REDSHIFT_USER") or os.environ.get("REDSHIFT_UID")
    password = os.environ.get("REDSHIFT_PASSWORD") or os.environ.get("REDSHIFT_PWD")

    if dsn:
        parts = [f"DSN={_odbc_value(dsn)}"]
        if database:
            parts.append(f"Database={_odbc_value(database)}")
        if user:
            parts.append(f"UID={_odbc_value(user)}")
        if password:
            parts.append(f"PWD={_odbc_value(password)}")
        return ";".join(parts)

    if host and user and password:
        driver = os.environ.get("REDSHIFT_DRIVER") or "Amazon Redshift (x64)"
        parts = [
            f"Driver={{{driver}}}",
            f"Server={_odbc_value(host)}",
            f"Port={_odbc_value(port)}",
            f"Database={_odbc_value(database)}",
            f"UID={_odbc_value(user)}",
            f"PWD={_odbc_value(password)}",
        ]
        return ";".join(parts)

    return default


def redshift_connection_source(default: str = "DSN=Redshift") -> str:
    _load_local_env()
    if os.environ.get("REDSHIFT_DSN") or os.environ.get("REDSHIFT_CONNECTION_STRING"):
        return "local REDSHIFT_DSN / REDSHIFT_CONNECTION_STRING"
    if os.environ.get("REDSHIFT_ODBC_DSN") or os.environ.get("ODBC_DSN"):
        return f"local ODBC DSN {os.environ.get('REDSHIFT_ODBC_DSN') or os.environ.get('ODBC_DSN')}"
    if (os.environ.get("REDSHIFT_HOST") or os.environ.get("REDSHIFT_SERVER")) and (
        os.environ.get("REDSHIFT_PASSWORD") or os.environ.get("REDSHIFT_PWD")
    ):
        return "local REDSHIFT_HOST/REDSHIFT_USER/REDSHIFT_PASSWORD"
    return default


def sanitize_connection_error(exc: Exception) -> str:
    text = str(exc)
    text = re.sub(r"(PWD|Password|password)=([^;)'\\s]+)", r"\1=***", text)
    return text


def execute_odbc_sql(connection_string: str, sql: str, timeout_seconds: int = 240) -> list[dict[str, Any]]:
    try:
        import pyodbc
    except ImportError as exc:
        raise RuntimeError("pyodbc is not installed for this Python interpreter. Run `python -m pip install pyodbc`.") from exc

    connection = pyodbc.connect(connection_string, timeout=timeout_seconds, autocommit=True)
    try:
        cursor = connection.cursor()
        try:
            if hasattr(cursor, "timeout"):
                cursor.timeout = timeout_seconds
            cursor.execute(sql)
            columns = [column[0] for column in cursor.description or []]
            return [
                {columns[index]: _normalize_value(value) for index, value in enumerate(row)}
                for row in cursor.fetchall()
            ]
        finally:
            cursor.close()
    finally:
        connection.close()
