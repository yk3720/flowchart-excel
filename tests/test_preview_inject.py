"""preview_inject のユニットテスト。"""
import json
import unittest

from app.core.preview_inject import build_payload_inject_js, build_set_global_js


class PreviewInjectTests(unittest.TestCase):
    def test_inject_js_contains_payload(self) -> None:
        payload = {"title": "T", "table": [["10", "端子"]]}
        js = build_payload_inject_js(payload)
        self.assertIn("setPreviewPayload", js)
        self.assertIn(json.dumps(payload, ensure_ascii=False), js)


class BuildSetGlobalJsTests(unittest.TestCase):
    def test_calls_named_window_function_with_payload(self) -> None:
        payload = {"requestId": "abc", "proposals": []}
        js = build_set_global_js("setProposalResult", payload)
        self.assertIn("window.setProposalResult", js)
        self.assertIn(json.dumps(payload, ensure_ascii=False), js)

    def test_rejects_non_identifier_fn_name(self) -> None:
        with self.assertRaises(ValueError):
            build_set_global_js("not a fn; alert(1)", {})


if __name__ == "__main__":
    unittest.main()
