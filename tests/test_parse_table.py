"""parse_table のユニットテスト。"""
import unittest

from app.core.parse_table import parse_table_rows


class ParseTableTests(unittest.TestCase):
    def test_parses_10col_v2(self) -> None:
        data = (
            ("10", "端子", "", "20", "", 0, 0, "開始", "", ""),
            ("20", "処理", "", "30", "", 1, 0, "手順A", "", ""),
        )
        nodes, _, col_count = parse_table_rows(data)
        self.assertEqual(col_count, 10)
        self.assertEqual(len(nodes), 2)
        self.assertEqual(nodes[0]["tier"], 0)
        self.assertEqual(nodes[0]["level"], 0)
        self.assertEqual(nodes[1]["dests_down"], ["30"])

    def test_parses_8col_legacy(self) -> None:
        data = (
            (10, "端子", "20", "", 0, "開始", "", ""),
            (20, "処理", "30", "", 0, "手順", "", ""),
        )
        nodes, _, col_count = parse_table_rows(data)
        self.assertEqual(col_count, 8)
        self.assertNotIn("tier", nodes[0])
        self.assertEqual(nodes[0]["level"], 0)
        self.assertEqual(nodes[0]["dests_down"], ["20"])

    def test_skips_header_row(self) -> None:
        data = (
            ("ID", "図形種別", "色", "接続先(下)", "接続先(右)", "段", "列", "Text1", "Text2", "Text3"),
            ("10", "端子", "", "20", "", 0, 0, "開始", "", ""),
        )
        nodes, _, _ = parse_table_rows(data)
        self.assertEqual(len(nodes), 1)
        self.assertEqual(nodes[0]["id"], "10")

    def test_skips_row_with_blank_shape_type(self) -> None:
        data = (
            ("10", "端子", "", "20", "", 0, 0, "開始", "", ""),
            ("15", "", "", "20", "", 0, 0, "空欄種別", "", ""),
            ("20", "処理", "", "30", "", 1, 0, "手順A", "", ""),
        )
        nodes, _, _ = parse_table_rows(data)
        self.assertEqual([n["id"] for n in nodes], ["10", "20"])

    def test_skips_duplicate_id_and_keeps_first(self) -> None:
        data = (
            ("10", "端子", "", "20", "", 0, 0, "最初", "", ""),
            ("10", "処理", "", "30", "", 1, 0, "重複", "", ""),
            ("20", "処理", "", "", "", 1, 0, "手順A", "", ""),
        )
        nodes, _, _ = parse_table_rows(data)
        self.assertEqual([n["id"] for n in nodes], ["10", "20"])
        self.assertEqual(nodes[0]["full_text"], "最初")

    def test_v1_v2_detection_uses_majority_across_rows(self) -> None:
        # 先頭行は接続先(下)が空(判断ノードで右方向のみ使用)のため単独では
        # v2 と誤判定されうるが、後続行の多数決で v1(9-10列) と正しく判定されること。
        data = (
            ("40", "判断", "", "50", 3, 0, "分岐?", "", "", "黄"),
            ("50", "処理", "60", "", 4, 1, "手順A", "", "", ""),
            ("60", "処理", "70", "", 5, 1, "手順B", "", "", ""),
            ("70", "処理", "80", "", 6, 1, "手順C", "", "", ""),
            ("80", "端子", "", "", 7, 1, "終了", "", "", ""),
        )
        nodes, _, col_count = parse_table_rows(data)
        self.assertEqual(col_count, 10)
        self.assertEqual(nodes[0]["tier"], 3)
        self.assertEqual(nodes[0]["dests_right"], ["50"])
        self.assertEqual(nodes[0]["dests_down"], [])

    def test_parses_judgment_branch(self) -> None:
        data = (
            ("40", "判断", "", "60", "50", 3, 0, "条件?", "", ""),
            ("50", "処理", "", "30", "", 4, 1, "再試行", "", ""),
        )
        nodes, _, _ = parse_table_rows(data, force_v2=True)
        self.assertEqual(nodes[1]["tier"], 4)
        self.assertEqual(nodes[1]["level"], 1)
        self.assertEqual(nodes[1]["dests_down"], ["30"])


if __name__ == "__main__":
    unittest.main()
