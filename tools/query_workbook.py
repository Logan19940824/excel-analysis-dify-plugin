from dify_plugin import Tool

from services.workbook_analyzer import analyzer_from_env


class QueryWorkbookTool(Tool):
    def _invoke(self, tool_parameters):
        file_url = (tool_parameters.get("file_url") or "").strip()
        sql = (tool_parameters.get("sql") or "").strip()
        expected_sha256 = (tool_parameters.get("sha256") or "").strip() or None
        limit = int(tool_parameters.get("limit") or 1000)
        if not file_url:
            raise ValueError("file_url is required")
        if not sql:
            raise ValueError("sql is required")
        yield self.create_json_message(analyzer_from_env().query(file_url, sql, expected_sha256, limit))
