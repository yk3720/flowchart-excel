"""プレビュー確定スナップショットのユニットテスト（Excel 非依存）。"""
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from app.core.live_preview import table_fingerprint
from app.core.preview_payload import table_list_to_com_tuple


class PreviewSnapshotTests(unittest.TestCase):
    def test_table_list_to_com_tuple_roundtrip_shape(self) -> None:
        table = [["10", "端子", None], ["20", "処理", "10"]]
        data = table_list_to_com_tuple(table)
        self.assertEqual(len(data), 2)
        self.assertEqual(data[0][0], "10")

    def test_fingerprint_stable_for_same_payload(self) -> None:
        payload = {
            "title": "T",
            "schema": "table-10col-v2",
            "table": [["10", "端子"]],
            "layout": {"width": 1, "heightMin": 1, "gapV": 1, "gapH": 1},
        }
        self.assertEqual(table_fingerprint(payload), table_fingerprint(payload))

    @patch("app.core.live_preview.try_refresh_studio_payload")
    def test_confirm_writes_snapshot(self, mock_refresh: MagicMock) -> None:
        from app.preview_host import run_preview_host

        mock_webview = MagicMock()
        window = MagicMock()
        mock_webview.create_window.return_value = window
        mock_webview.windows = [window]
        mock_refresh.return_value = None

        def fake_start(func=None, **kwargs):
            api = mock_webview.create_window.call_args.kwargs["js_api"]
            api.confirm()

        mock_webview.start.side_effect = fake_start

        payload = {
            "title": "T",
            "isFullMode": True,
            "schema": "table-10col-v2",
            "table": [["10", "端子"]],
            "layout": {
                "width": 160,
                "heightMin": 60,
                "gapV": 30,
                "gapH": 100,
                "baseLeft": 40,
                "baseTop": 40,
            },
            "meta": {
                "watch": {
                    "workbookName": "Book1",
                    "sheetName": "Sheet1",
                    "anchorAddress": "$A$1",
                    "isFullMode": True,
                }
            },
        }

        with tempfile.TemporaryDirectory() as tmp, patch.dict(sys.modules, {"webview": mock_webview}):
            tmp_path = Path(tmp)
            payload_path = tmp_path / "payload.json"
            result_path = tmp_path / "result.json"
            dist_dir = tmp_path / "dist"
            dist_dir.mkdir()
            (dist_dir / "index.html").write_text("<html></html>", encoding="utf-8")
            payload_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

            code = run_preview_host(payload_path, result_path, dist_dir)
            self.assertEqual(code, 0)
            result = json.loads(result_path.read_text(encoding="utf-8"))
            self.assertEqual(result["action"], "confirm")
            self.assertIn("fingerprint", result)
            self.assertIn("payload", result)
            self.assertEqual(result["payload"]["title"], "T")


if __name__ == "__main__":
    unittest.main()
