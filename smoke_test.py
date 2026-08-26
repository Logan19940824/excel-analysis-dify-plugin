from __future__ import annotations

import csv
import threading
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tempfile import TemporaryDirectory

from services.workbook_analyzer import WorkbookAnalyzer


def main() -> None:
    with TemporaryDirectory() as directory:
        root = Path(directory)
        with (root / "sales.csv").open("w", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            writer.writerows([["region", "amount"], ["east", 10], ["east", 20], ["west", 5]])
        (root / "preview_url").write_bytes((root / "sales.csv").read_bytes())
        server = ThreadingHTTPServer(("127.0.0.1", 0), partial(SimpleHTTPRequestHandler, directory=directory))
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            # The analyzer rejects private addresses in production; use the test-only override below.
            import services.workbook_analyzer as module
            original = module._validate_url
            module._validate_url = lambda url: "preview"
            try:
                analyzer = WorkbookAnalyzer()
                url = f"http://127.0.0.1:{server.server_port}/preview_url"
                info = analyzer.inspect(url)
                assert info["sheets"][0]["columns"][0]["name"] == "region"
                result = analyzer.query(url, 'SELECT region, SUM(amount) AS total FROM "data" GROUP BY region', info["sha256"])
                assert result["row_count"] == 2
            finally:
                module._validate_url = original
        finally:
            server.shutdown()
            thread.join()
    print("dify_excel_plugin smoke test passed")


if __name__ == "__main__":
    main()
