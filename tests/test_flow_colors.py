"""flow_colors / 〇 レイアウトのユニットテスト。"""
from __future__ import annotations

import unittest

from app.core.flow_colors import (
    DEFAULT_FILL_HEX,
    fill_hex_for_hint,
    fill_vba_rgb_for_hint,
    hex_to_vba_rgb,
)
from app.core.layout_preview import build_preview_model
from app.core.parse_table import normalize_shape_type, parse_table_rows


class FlowColorsTests(unittest.TestCase):
    def test_hex_to_vba_rgb_matches_studio_fills(self) -> None:
        self.assertEqual(hex_to_vba_rgb("#ffffff"), 0xFFFFFF)
        # #fef9c3 → R=FE G=F9 B=C3 → FE + F9*256 + C3*65536
        self.assertEqual(hex_to_vba_rgb("#fef9c3"), 0xC3F9FE)
        self.assertEqual(hex_to_vba_rgb("#ffedd5"), 0xD5EDFF)
        self.assertEqual(hex_to_vba_rgb("#dbeafe"), 0xFEEADB)

    def test_fill_hex_for_hint(self) -> None:
        self.assertEqual(fill_hex_for_hint(None), DEFAULT_FILL_HEX)
        self.assertEqual(fill_hex_for_hint(""), DEFAULT_FILL_HEX)
        self.assertEqual(fill_hex_for_hint("黄"), "#fef9c3")
        self.assertEqual(fill_hex_for_hint("橙"), "#ffedd5")
        self.assertEqual(fill_hex_for_hint("青"), "#dbeafe")
        self.assertEqual(fill_hex_for_hint("赤"), DEFAULT_FILL_HEX)

    def test_fill_vba_rgb_for_hint(self) -> None:
        self.assertEqual(fill_vba_rgb_for_hint("黄"), 0xC3F9FE)


class OvalShapeTests(unittest.TestCase):
    def test_normalize_oval_aliases(self) -> None:
        self.assertEqual(normalize_shape_type("〇"), "〇")
        self.assertEqual(normalize_shape_type("○"), "〇")
        self.assertEqual(normalize_shape_type("省略記号"), "〇")

    def test_preview_places_oval_and_color(self) -> None:
        data = (
            ("10", "処理", "黄", "20", "", 0, 0, "手前", "", ""),
            ("20", "〇", "", "30", "", 1, 0, "1", "", ""),
            ("30", "処理", "橙", "", "", 2, 0, "先", "", ""),
        )
        nodes, row_map, _ = parse_table_rows(data)
        model = build_preview_model(
            nodes_raw=nodes,
            row_map=row_map,
            config={"height": 60.0, "width": 160.0, "gap_v": 30.0, "gap_h": 100.0},
            title="oval",
            is_full_mode=True,
        )
        by_id = {n.id: n for n in model.nodes}
        self.assertEqual(by_id["10"].color_hint, "黄")
        self.assertEqual(by_id["20"].shape_kind, "oval")
        self.assertEqual(by_id["20"].width, 60.0)
        self.assertEqual(by_id["20"].height, 60.0)
        self.assertAlmostEqual(by_id["20"].left, (160.0 - 60.0) / 2)
        self.assertEqual(by_id["30"].color_hint, "橙")


if __name__ == "__main__":
    unittest.main()
