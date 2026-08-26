from dify_plugin import Tool

from services.workbook_analyzer import analyzer_from_env


class GetWorkbookInfoTool(Tool):
    def _invoke(self, tool_parameters):
        file_url = (tool_parameters.get("file_url") or "").strip()
        if not file_url:
            raise ValueError("file_url is required")
        yield self.create_json_message(analyzer_from_env().inspect(file_url))
