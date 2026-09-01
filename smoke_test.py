from __future__ import annotations

import csv
import threading
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tempfile import TemporaryDirectory

from services.workbook_analyzer import WorkbookAnalyzer


class _FakeStorage:
    def __init__(self):
        self.data = {}

    def get(self, key):
        if key not in self.data:
            raise KeyError(key)
        return self.data[key]

    def set(self, key, value):
        self.data[key] = value


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
                analyzer = WorkbookAnalyzer(storage=_FakeStorage())
                url = f"http://127.0.0.1:{server.server_port}/preview_url"
                info = analyzer.inspect(url)
                assert info["sheets"][0]["columns"][0]["name"] == "region"
                result = analyzer.query(info["cache_id"], 'SELECT region, SUM(amount) AS total FROM "data" GROUP BY region; SELECT COUNT(*) AS cnt FROM "data"')
                assert len(result["results"]) == 2
                assert result["results"][0]["row_count"] == 2
                assert result["results"][1]["row_count"] == 1
                assert result["cache_id"] == info["cache_id"]
            finally:
                module._validate_url = original
        finally:
            server.shutdown()
            thread.join()
    print("dify_excel_plugin smoke test passed")


if __name__ == "__main__":
    main()
