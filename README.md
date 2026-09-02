# Excel Analysis Dify Plugin

This plugin provides two tools for analyzing Excel and CSV files:

- `get_workbook_info`: inspect sheets, columns, types, row counts, and sample rows.
- `query_workbook`: execute one or more read-only DuckDB `SELECT` or `CTE` queries.

`get_workbook_info` downloads the Dify-provided HTTP(S) file URL once, converts each worksheet to Parquet, and stores the Parquet data in Dify plugin storage under a returned `cache_id`. `query_workbook` restores the cached Parquet data and builds temporary DuckDB tables, so later queries do not depend on the preview URL remaining valid.

The query must use the exact sheet and column names returned by `get_workbook_info`. Multiple statements are separated by semicolons; each returns at most 1,000 rows.

## Usage

1. Call `get_workbook_info` with the Dify file URL.
2. Use the returned `cache_id`, sheet names, and column names when calling `query_workbook`; do not pass the original URL again.
3. Pass one or more read-only `SELECT` or `CTE` statements separated by English semicolons. For clarity, put each statement on a new line.

The plugin requires Python 3.12 at runtime. See `PRIVACY.md` for data handling details and `readme/README_zh_Hans.md` for the Chinese documentation.

Source repository: https://github.com/Logan19940824/excel-analysis-dify-plugin
