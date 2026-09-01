import logging
import os
import sys

from dify_plugin import Plugin, DifyPluginEnv
from dify_plugin.config.logger_format import plugin_logger_handler

log_level = getattr(logging, os.getenv("EXCEL_LOG_LEVEL", "INFO").upper(), logging.INFO)
plugin_logger_handler.setLevel(log_level)
logging.basicConfig(
    level=log_level,
    format="%(asctime)s %(levelname)s %(name)s - %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("excel_analysis")
logger.setLevel(log_level)
logger.addHandler(plugin_logger_handler)
logger.propagate = False

plugin = Plugin(DifyPluginEnv(MAX_REQUEST_TIMEOUT=120))

if __name__ == '__main__':
    logger.info("starting excel_analysis plugin")
    plugin.run()
