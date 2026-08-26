# Excel Analysis Dify Plugin

This plugin provides two stateless tools for analyzing Excel and CSV files:

- `get_workbook_info`: inspect sheets, columns, types, row counts, and sample rows.
- `query_workbook`: execute one read-only DuckDB `SELECT` or `CTE` query.

Each tool downloads the Dify-provided HTTP(S) file URL, builds a temporary DuckDB database, performs its operation, and removes the temporary files. The plugin does not persist workbooks and does not depend on the repository's FastMCP service.

The URL must remain accessible between inspection and query calls. The query must use the exact sheet and column names returned by `get_workbook_info`. Queries are limited to one read-only DuckDB `SELECT` or `CTE` statement, with a maximum result limit of 1,000 rows.

## Usage

1. Call `get_workbook_info` with the Dify file URL.
2. Use the returned sheet names, column names, and `sha256` when calling `query_workbook`.
3. Pass one read-only `SELECT` or `CTE` statement and an optional result limit.

The plugin requires Python 3.12 at runtime. See `PRIVACY.md` for data handling details and `readme/README_zh_Hans.md` for the Chinese documentation.

Source repository: https://github.com/Logan19940824/excel-analysis-dify-plugin
