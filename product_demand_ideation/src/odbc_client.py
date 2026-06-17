from __future__ import annotations

from decimal import Decimal
from typing import Any


def _normalize_value(value: Any) -> Any:
    if isinstance(value, Decimal):
        return float(value)
    return value


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
