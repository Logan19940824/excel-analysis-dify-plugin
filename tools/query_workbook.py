import logging

from dify_plugin import Tool

from services.workbook_analyzer import analyzer_from_env

logger = logging.getLogger("excel_analysis")

class QueryWorkbookTool(Tool):
    def _invoke(self, tool_parameters):
        logger.info("query_workbook started")
        cache_id = (tool_parameters.get("cache_id") or "").strip()
        sql = (tool_parameters.get("sql") or "").strip()
        limit = int(tool_parameters.get("limit") or 1000)
        if not cache_id:
            raise ValueError("cache_id is required; call get_workbook_info first")
        if not sql:
            raise ValueError("sql is required")
        logger.info("query_workbook parameters accepted cache_id=%s limit=%d", cache_id, limit)
        try:
            result = analyzer_from_env(storage=self.session.storage).query(cache_id, sql, limit)
        except Exception:
            logger.exception("query_workbook failed cache_id=%s", cache_id)
            raise
        logger.info("query_workbook completed cache_id=%s result_sets=%d", cache_id, len(result["results"]))
        yield self.create_json_message(result)
