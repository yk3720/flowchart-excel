"""preview_payload のユニットテスト。"""
import unittest

from app.core.preview_payload import build_studio_preview_payload, resolve_schema


class PreviewPayloadTests(unittest.TestCase):
    def test_builds_v2_payload(self) -> None:
        data = (
            ("10", "端子", "", "20", "", 0, 0, "開始", "", ""),
            ("20", "処理", "", "", "", 1, 0, "手順", "", ""),
        )
        payload = build_studio_preview_payload(
            data,
            title="T",
            is_full_mode=True,
            config={"height": 60, "width": 160, "gap_v": 30, "gap_h": 100},
        )
        self.assertEqual(payload["schema"], "table-10col-v2")
        self.assertEqual(len(payload["table"]), 2)
        self.assertEqual(payload["layout"]["heightMin"], 60)

    def test_resolve_schema_8col(self) -> None:
        table = [[10, "端子", "20", "", 0, "開始", "", ""]]
        self.assertIsNone(resolve_schema(table))


if __name__ == "__main__":
    unittest.main()
