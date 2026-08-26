from __future__ import annotations

import hashlib
import ipaddress
import json
import socket
import shutil
import tempfile
import threading
from contextlib import contextmanager
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterator
from urllib.parse import unquote, urlparse

import duckdb
import httpx
import sqlglot
from openpyxl import load_workbook
from sqlglot import expressions as exp


def _json_value(value: Any) -> Any:
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, (bytes, bytearray)):
        return bytes(value).hex()
    if isinstance(value, Decimal):
        return str(value)
    return value


def _quote(identifier: str) -> str:
    return '"' + identifier.replace('"', '""') + '"'


def _safe_headers(values: tuple[Any, ...]) -> list[str]:
    result: list[str] = []
    used: set[str] = set()
    for index, value in enumerate(values, 1):
        base = str(value).strip() if value is not None else ""
        base = base or f"column_{index}"
        name, suffix = base, 2
        while name.casefold() in used:
            name = f"{base}_{suffix}"
            suffix += 1
        used.add(name.casefold())
        result.append(name)
    return result


def _validate_url(url: str) -> str:
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        raise ValueError("file_url must be an HTTP(S) URL")
    if parsed.username or parsed.password:
        raise ValueError("file_url must not contain embedded credentials")
    try:
        addresses = {item[4][0] for item in socket.getaddrinfo(parsed.hostname, parsed.port or 443)}
    except socket.gaierror as exc:
        raise ValueError("file_url hostname could not be resolved") from exc
    for value in addresses:
        address = ipaddress.ip_address(value)
        if address.is_private or address.is_loopback or address.is_link_local or address.is_reserved:
            raise ValueError("file_url must resolve to a public address")
    filename = Path(unquote(parsed.path)).name or "workbook"
    return filename


class WorkbookAnalyzer:
    def __init__(self, max_bytes: int = 50 * 1024 * 1024, max_rows: int = 1000,
                 timeout_seconds: int = 30, memory_limit: str = "512MB"):
        self.max_bytes = max_bytes
        self.max_rows = max_rows
        self.timeout_seconds = timeout_seconds
        self.memory_limit = memory_limit

    def inspect(self, file_url: str) -> dict[str, Any]:
        with self._database(file_url) as (connection, sha256, filename, file_format):
            sheets = []
            for table in self._tables(connection):
                columns = [
                    {"name": row[0], "type": row[1]}
                    for row in connection.execute(f"DESCRIBE {_quote(table)}").fetchall()
                ]
                row_count = connection.execute(f"SELECT COUNT(*) FROM {_quote(table)}").fetchone()[0]
                sample = connection.execute(
                    f"SELECT * FROM {_quote(table)} LIMIT 10"
                ).fetchall()
                sheets.append({
                    "name": table,
                    "row_count": row_count,
                    "columns": columns,
                    "sample_rows": [[_json_value(value) for value in row] for row in sample],
                })
            return {
                "file_url": file_url,
                "sha256": sha256,
                "filename": filename,
                "format": file_format,
                "sheets": sheets,
            }

    def query(self, file_url: str, sql: str, expected_sha256: str | None = None,
              limit: int = 1000) -> dict[str, Any]:
        statements = sqlglot.parse(sql, read="duckdb")
        if len(statements) != 1 or not isinstance(statements[0], exp.Query):
            raise ValueError("only one read-only SELECT or CTE statement is allowed")
        forbidden = {"Insert", "Update", "Delete", "Create", "Drop", "Alter", "Command", "Copy", "Attach", "Merge"}
        if any(node.__class__.__name__ in forbidden for node in statements[0].walk()):
            raise ValueError("SQL statement is not read-only")
        effective_limit = self._limit(limit)
        with self._database(file_url) as (connection, sha256, _filename, _format):
            if expected_sha256 and expected_sha256 != sha256:
                raise ValueError("source file changed; call get_workbook_info again")
            normalized = statements[0].sql(dialect="duckdb")
            timer = threading.Timer(self.timeout_seconds, connection.interrupt)
            try:
                timer.start()
                cursor = connection.execute(
                    f"SELECT * FROM ({normalized}) AS _result LIMIT ?",
                    [effective_limit + 1],
                )
                columns = [{"name": item[0], "type": str(item[1])} for item in cursor.description]
                rows = [[_json_value(value) for value in row] for row in cursor.fetchall()]
            except duckdb.InterruptException as exc:
                raise TimeoutError(f"query exceeded {self.timeout_seconds} seconds") from exc
            finally:
                timer.cancel()
            truncated = len(rows) > effective_limit
            if truncated:
                rows = rows[:effective_limit]
            return {
                "sha256": sha256,
                "columns": columns,
                "rows": rows,
                "row_count": len(rows),
                "truncated": truncated,
            }

    @contextmanager
    def _database(self, file_url: str) -> Iterator[tuple[duckdb.DuckDBPyConnection, str, str, str]]:
        filename = _validate_url(file_url)
        with tempfile.TemporaryDirectory(prefix="dify-excel-") as directory:
            source = Path(directory) / "source"
            self._download(file_url, source)
            data = source.read_bytes()
            sha256 = hashlib.sha256(data).hexdigest()
            connection = None
            selected = None
            error: Exception | None = None
            # The Dify preview URL often has no real filename suffix. Detect the
            # format by importing the downloaded bytes, not by inspecting its URL.
            for suffix in (".xlsx", ".csv"):
                database = Path(directory) / "workbook.duckdb"
                try:
                    connection = duckdb.connect(
                        str(database),
                        config={"memory_limit": self.memory_limit},
                    )
                    if suffix == ".csv":
                        connection.execute('CREATE TABLE "data" AS SELECT * FROM read_csv_auto(?)', [str(source)])
                    else:
                        self._import_xlsx(connection, source)
                    # CSV/XLSX import needs local file access; user SQL must not.
                    connection.execute("SET enable_external_access=false")
                    selected = suffix
                    break
                except Exception as exc:
                    error = exc
                    if connection:
                        connection.close()
                    connection = None
                    database.unlink(missing_ok=True)
            if connection is None or selected is None:
                raise ValueError("file cannot be read as .xlsx or .csv") from error
            try:
                yield connection, sha256, filename, selected[1:]
            finally:
                connection.close()

    def _download(self, url: str, destination: Path) -> None:
        current = url
        with httpx.Client(timeout=httpx.Timeout(30), follow_redirects=False) as client:
            for _ in range(6):
                _validate_url(current)
                with client.stream("GET", current, headers={"User-Agent": "dify-excel-plugin/1.0"}) as response:
                    if response.is_redirect:
                        location = response.headers.get("location")
                        if not location:
                            raise ValueError("download redirect has no location")
                        current = str(response.url.join(location))
                        continue
                    response.raise_for_status()
                    content_length = response.headers.get("content-length")
                    if content_length and int(content_length) > self.max_bytes:
                        raise ValueError("workbook exceeds the configured size limit")
                    size = 0
                    with destination.open("wb") as output:
                        for chunk in response.iter_bytes():
                            size += len(chunk)
                            if size > self.max_bytes:
                                raise ValueError("workbook exceeds the configured size limit")
                            output.write(chunk)
                    return
            raise ValueError("too many download redirects")

    def _import_xlsx(self, connection: duckdb.DuckDBPyConnection, source: Path) -> None:
        # openpyxl validates the filename suffix before inspecting the file bytes.
        workbook_source = source.with_suffix(".xlsx")
        shutil.copyfile(source, workbook_source)
        workbook = load_workbook(workbook_source, read_only=True, data_only=True)
        try:
            for sheet in workbook.worksheets:
                rows = sheet.iter_rows(values_only=True)
                try:
                    headers = _safe_headers(next(rows))
                except StopIteration:
                    headers = ["column_1"]
                    rows = iter(())
                connection.execute(
                    f"CREATE TABLE {_quote(sheet.title)} ({', '.join(f'{_quote(header)} VARCHAR' for header in headers)})"
                )
                values = [tuple(row[:len(headers)]) for row in rows]
                if values:
                    connection.executemany(
                        f"INSERT INTO {_quote(sheet.title)} VALUES ({', '.join('?' for _ in headers)})",
                        values,
                    )
        finally:
            workbook.close()
            workbook_source.unlink(missing_ok=True)

    @staticmethod
    def _tables(connection: duckdb.DuckDBPyConnection) -> list[str]:
        return [row[0] for row in connection.execute("SHOW TABLES").fetchall()]

    def _limit(self, value: int) -> int:
        if value < 1:
            raise ValueError("limit must be positive")
        return min(value, self.max_rows)


def analyzer_from_env() -> WorkbookAnalyzer:
    import os
    return WorkbookAnalyzer(
        max_bytes=int(os.getenv("EXCEL_MAX_BYTES", str(50 * 1024 * 1024))),
        max_rows=int(os.getenv("EXCEL_SQL_MAX_ROWS", "1000")),
        timeout_seconds=int(os.getenv("EXCEL_SQL_TIMEOUT_SECONDS", "30")),
        memory_limit=os.getenv("EXCEL_SQL_MEMORY_LIMIT", "512MB"),
    )
