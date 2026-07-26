"""live_preview のユニットテスト（Excel 非依存）。"""
import unittest

from app.core.live_preview import layout_to_config, table_fingerprint


class LivePreviewTests(unittest.TestCase):
    def test_layout_to_config(self) -> None:
        cfg = layout_to_config(
            {"width": 160, "heightMin": 60, "gapV": 30, "gapH": 100}
        )
        self.assertEqual(cfg["height"], 60)
        self.assertEqual(cfg["gap_v"], 30)

    def test_fingerprint_changes_with_table(self) -> None:
        base = {
            "title": "T",
            "schema": "table-10col-v2",
            "table": [["10", "端子"]],
            "layout": {"width": 1, "heightMin": 1, "gapV": 1, "gapH": 1},
        }
        a = table_fingerprint(base)
        b = table_fingerprint({**base, "table": [["10", "処理"]]})
        self.assertNotEqual(a, b)


if __name__ == "__main__":
    unittest.main()
