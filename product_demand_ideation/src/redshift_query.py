from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any

from odbc_client import execute_odbc_sql, redshift_connection_source, redshift_connection_string, sanitize_connection_error


DEFAULT_REDSHIFT_ODBC_CONNECTION = "DSN=Redshift"
FORCE_ODBC_ENV = "SUNCO_REDSHIFT_FORCE_ODBC"
DISABLE_MCP_ENV = "SUNCO_DISABLE_REDSHIFT_MCP"
DISABLE_ODBC_FALLBACK_ENV = "SUNCO_REDSHIFT_DISABLE_ODBC_FALLBACK"


def _env_flag(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in {"1", "true", "yes", "y"}


def is_redshift_mcp_source(source: str) -> bool:
    return source.strip().lower() == "redshift mcp"


def _ensure_backend_app_on_path() -> None:
    current = Path(__file__).resolve()
    for parent in current.parents:
        candidate = parent / "backend" / "app"
        if candidate.exists():
            if str(candidate) not in sys.path:
                sys.path.insert(0, str(candidate))
            return


def sanitize_redshift_error(exc: Exception) -> str:
    try:
        _ensure_backend_app_on_path()
        from opportunity_engine.db_mcp import sanitize_mcp_error

        return sanitize_connection_error(Exception(sanitize_mcp_error(exc)))
    except Exception:
        return sanitize_connection_error(exc)


def _redshift_mcp_client(timeout_seconds: int, client_name: str):
    _ensure_backend_app_on_path()
    from opportunity_engine.db_mcp import REDSHIFT_MCP_URL, McpRemoteClient

    return McpRemoteClient(
        server_url=os.environ.get("REDSHIFT_MCP_URL") or REDSHIFT_MCP_URL,
        timeout_seconds=timeout_seconds,
        client_name=client_name,
    )


class RedshiftQueryClient:
    """Execute Redshift SQL through MCP first, with local ODBC as the fallback."""

    def __init__(
        self,
        timeout_seconds: int = 240,
        default_odbc: str = DEFAULT_REDSHIFT_ODBC_CONNECTION,
        client_name: str = "sunco-ideation-redshift",
    ) -> None:
        self.timeout_seconds = timeout_seconds
        self.default_odbc = default_odbc
        self.client_name = client_name
        self.connection_source = ""
        self._mcp_client: Any | None = None
        self._mcp_start_error: str | None = None

    def __enter__(self) -> "RedshiftQueryClient":
        if not _env_flag(FORCE_ODBC_ENV) and not _env_flag(DISABLE_MCP_ENV):
            try:
                self._mcp_client = _redshift_mcp_client(self.timeout_seconds, self.client_name)
                self._mcp_client.__enter__()
                self.connection_source = "Redshift MCP"
                return self
            except Exception as exc:
                self._mcp_client = None
                self._mcp_start_error = sanitize_redshift_error(exc)
                if _env_flag(DISABLE_ODBC_FALLBACK_ENV):
                    raise RuntimeError(f"Redshift MCP unavailable and ODBC fallback is disabled: {self._mcp_start_error}") from exc

        self.connection_source = redshift_connection_source(self.default_odbc)
        if self._mcp_start_error:
            self.connection_source = f"{self.connection_source} (ODBC fallback)"
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self._mcp_client is not None:
            self._mcp_client.__exit__(exc_type, exc, tb)

    @property
    def using_mcp(self) -> bool:
        return self._mcp_client is not None

    @property
    def mcp_start_error(self) -> str | None:
        return self._mcp_start_error

    def execute_sql(self, sql: str, timeout_seconds: int | None = None) -> list[dict[str, Any]]:
        timeout = timeout_seconds or self.timeout_seconds
        if self._mcp_client is not None:
            return self._mcp_client.execute_sql(sql, timeout_seconds=timeout)
        return execute_odbc_sql(redshift_connection_string(self.default_odbc), sql, timeout_seconds=timeout)


def execute_redshift_sql(
    sql: str,
    timeout_seconds: int = 240,
    default_odbc: str = DEFAULT_REDSHIFT_ODBC_CONNECTION,
    client_name: str = "sunco-ideation-redshift",
) -> tuple[list[dict[str, Any]], str]:
    with RedshiftQueryClient(timeout_seconds=timeout_seconds, default_odbc=default_odbc, client_name=client_name) as client:
        return client.execute_sql(sql, timeout_seconds=timeout_seconds), client.connection_source
