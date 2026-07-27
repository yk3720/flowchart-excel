"""preview_inject のユニットテスト。"""
import json
import unittest

from app.core.preview_inject import build_payload_inject_js


class PreviewInjectTests(unittest.TestCase):
    def test_inject_js_contains_payload(self) -> None:
        payload = {"title": "T", "table": [["10", "端子"]]}
        js = build_payload_inject_js(payload)
        self.assertIn("setPreviewPayload", js)
        self.assertIn(json.dumps(payload, ensure_ascii=False), js)


if __name__ == "__main__":
    unittest.main()
