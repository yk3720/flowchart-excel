"""preview_host のユニットテスト（WebView 起動引数 · Excel 非依存）。"""
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch


class PreviewHostTests(unittest.TestCase):
    def test_create_window_uses_on_top(self) -> None:
        from app.preview_host import run_preview_host

        mock_webview = MagicMock()
        window = MagicMock()
        mock_webview.create_window.return_value = window
        mock_webview.windows = [window]

        with tempfile.TemporaryDirectory() as tmp, patch.dict(sys.modules, {"webview": mock_webview}):
            tmp_path = Path(tmp)
            payload_path = tmp_path / "payload.json"
            result_path = tmp_path / "result.json"
            dist_dir = tmp_path / "dist"
            dist_dir.mkdir()
            (dist_dir / "index.html").write_text("<html></html>", encoding="utf-8")
            payload_path.write_text(
                json.dumps(
                    {
                        "title": "T",
                        "schema": "table-10col-v2",
                        "table": [["10", "端子"]],
                        "layout": {
                            "width": 1,
                            "heightMin": 1,
                            "gapV": 1,
                            "gapH": 1,
                            "baseLeft": 0,
                            "baseTop": 0,
                        },
                        "meta": {"watch": {"workbookName": "Book1", "sheetName": "Sheet1"}},
                    }
                ),
                encoding="utf-8",
            )

            code = run_preview_host(payload_path, result_path, dist_dir)
            self.assertEqual(code, 0)
            _, kwargs = mock_webview.create_window.call_args
            self.assertTrue(kwargs.get("on_top"))


if __name__ == "__main__":
    unittest.main()
