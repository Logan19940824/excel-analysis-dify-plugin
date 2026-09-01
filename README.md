# Excel Analysis Dify Plugin

This plugin provides two stateless tools for analyzing Excel and CSV files:

- `get_workbook_info`: inspect sheets, columns, types, row counts, and sample rows.
- `query_workbook`: execute one or more read-only DuckDB `SELECT` or `CTE` queries.

Each tool downloads the Dify-provided HTTP(S) file URL, builds a temporary DuckDB database, performs its operation, and removes the temporary files. The plugin does not persist workbooks and does not depend on the repository's FastMCP service.

The query must use the exact sheet and column names returned by `get_workbook_info`. Multiple statements are separated by semicolons; each returns at most 1,000 rows.

## Usage

1. Call `get_workbook_info` with the Dify file URL.
2. Use the returned `cache_id`, sheet names, and column names when calling `query_workbook`; do not pass the original URL again.
3. Pass one or more read-only `SELECT` or `CTE` statements separated by semicolons and an optional result limit.

The plugin requires Python 3.12 at runtime. See `PRIVACY.md` for data handling details and `readme/README_zh_Hans.md` for the Chinese documentation.

Source repository: https://github.com/Logan19940824/excel-analysis-dify-plugin
