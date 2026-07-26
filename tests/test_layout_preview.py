"""layout_preview のユニットテスト。"""
import unittest

from app.core.layout_preview import build_preview_model
from app.core.parse_table import parse_table_rows


class LayoutPreviewTests(unittest.TestCase):
    def test_preview_places_tiers_and_edges(self) -> None:
        data = (
            ("10", "端子", "", "20", "", 0, 0, "開始", "", ""),
            ("20", "判断", "", "40", "30", 1, 0, "条件?", "", ""),
            ("30", "処理", "", "40", "", 2, 1, "No側", "", ""),
            ("40", "端子", "", "", "", 3, 0, "終了", "", ""),
        )
        nodes, row_map, _ = parse_table_rows(data)
        model = build_preview_model(
            nodes_raw=nodes,
            row_map=row_map,
            config={"height": 60.0, "width": 160.0, "gap_v": 30.0, "gap_h": 100.0},
            title="テスト",
            is_full_mode=True,
        )
        self.assertEqual(model.node_count, 4)
        self.assertEqual(model.edge_count, 4)  # 20→40, 20→30, 10→20, 30→40
        by_id = {n.id: n for n in model.nodes}
        self.assertEqual(by_id["10"].shape_kind, "roundrect")
        self.assertEqual(by_id["20"].shape_kind, "diamond")
        self.assertGreater(by_id["30"].left, by_id["20"].left)
        self.assertGreater(by_id["40"].top, by_id["10"].top)

    def test_missing_dest_warns(self) -> None:
        data = (("10", "処理", "", "999", "", 0, 0, "孤立", "", ""),)
        nodes, row_map, _ = parse_table_rows(data)
        model = build_preview_model(
            nodes_raw=nodes,
            row_map=row_map,
            config={"height": 60.0, "width": 160.0, "gap_v": 30.0, "gap_h": 100.0},
            title="警告",
            is_full_mode=False,
        )
        self.assertEqual(model.edge_count, 0)
        self.assertTrue(any("接続先" in w for w in model.warnings))


if __name__ == "__main__":
    unittest.main()
