import logging

from dify_plugin import Tool

from services.workbook_analyzer import analyzer_from_env

logger = logging.getLogger("excel_analysis")

class GetWorkbookInfoTool(Tool):
    def _invoke(self, tool_parameters):
        logger.info("get_workbook_info started")
        file_url = (tool_parameters.get("file_url") or "").strip()
        if not file_url:
            raise ValueError("file_url is required")
        try:
            result = analyzer_from_env(storage=self.session.storage).inspect(file_url)
        except Exception:
            logger.exception("get_workbook_info failed")
            raise
        logger.info("get_workbook_info completed cache_id=%s sheets=%d", result["cache_id"], len(result["sheets"]))
        yield self.create_json_message(result)
