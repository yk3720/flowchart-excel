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
