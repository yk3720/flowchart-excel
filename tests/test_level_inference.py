"""level_inference のユニットテスト（C-2 固定点反復ロジック）。"""
import threading
import unittest

from app.constants import _LEGACY_TEMPLATE_8, _legacy_8_to_10_v2
from app.core.level_inference import compute_level_proposals
from app.core.parse_table import parse_level_optional, parse_table_rows


def _nodes_with_raw_levels(rows_10col, blank_ids=()):
    """10列 v2 データから nodes と raw_levels(空欄反映済み) を作る。"""
    data = tuple(tuple(row) for row in rows_10col)
    nodes, _, _ = parse_table_rows(data)
    raw_levels = {}
    for node, row in zip(nodes, rows_10col):
        level_cell = None if node["id"] in blank_ids else row[6]
        raw_levels[node["id"]] = parse_level_optional(level_cell)
    return nodes, raw_levels


class ComplexYesCycleTests(unittest.TestCase):
    """実サンプル complex_yes（判断ノードの No 分岐で後方ループを含む）での安全性。"""

    def setUp(self) -> None:
        self.rows = _legacy_8_to_10_v2(_LEGACY_TEMPLATE_8["complex_yes"])

    def test_no_hang_and_no_crash_with_all_values_present(self) -> None:
        nodes, raw_levels = _nodes_with_raw_levels(self.rows)
        result = compute_level_proposals(nodes, raw_levels, mode="blank_only")
        self.assertFalse(result.cancelled)
        # 全セルに既存値があるため blank_only では何も提案しない
        self.assertEqual(result.proposals, [])

    def test_fills_blank_level_inside_cycle_from_settled_neighbor(self) -> None:
        # id=70（No分岐、dest_down=50で後方ループ）の列を空欄にしても
        # 他の確定値（id=60, 列0からの右継承）から固定点反復で埋まる。
        nodes, raw_levels = _nodes_with_raw_levels(self.rows, blank_ids=["70"])
        result = compute_level_proposals(nodes, raw_levels, mode="blank_only")
        self.assertFalse(result.cancelled)
        proposed = {p.node_id: p.proposed for p in result.proposals}
        self.assertIn("70", proposed)
        self.assertEqual(proposed["70"], 1)  # id=60(列0)の右接続 → +1


class NoAnchorCycleTests(unittest.TestCase):
    def test_two_nodes_pointing_at_each_other_with_no_anchor_are_unresolved(self) -> None:
        nodes = [
            {"id": "1", "dests_down": ["2"], "dests_right": []},
            {"id": "2", "dests_down": ["1"], "dests_right": []},
        ]
        raw_levels = {"1": None, "2": None}
        result = compute_level_proposals(nodes, raw_levels, mode="blank_only")
        self.assertFalse(result.cancelled)
        self.assertEqual(result.proposals, [])
        review_ids = {item.node_id for item in result.needs_review}
        self.assertEqual(review_ids, {"1", "2"})


class MultiDestSkipTests(unittest.TestCase):
    def test_row_with_multiple_dests_down_is_skipped_as_source_and_target(self) -> None:
        nodes = [
            {"id": "1", "dests_down": ["2", "3"], "dests_right": []},
            {"id": "2", "dests_down": [], "dests_right": []},
            {"id": "3", "dests_down": [], "dests_right": []},
        ]
        raw_levels = {"1": None, "2": None, "3": None}
        result = compute_level_proposals(nodes, raw_levels, mode="blank_only")
        self.assertFalse(result.cancelled)
        self.assertIn("1", result.skipped_multi_dest)
        proposed_ids = {p.node_id for p in result.proposals}
        # id=1 はカンマ区切り複数宛先の「行」自体ではないため対象になりうるが
        # ソースが無いため未提案。id=2/3 は id=1 からの推測が使われないため未確定。
        self.assertNotIn("2", proposed_ids)
        self.assertNotIn("3", proposed_ids)


class ConflictingInferenceTests(unittest.TestCase):
    def test_two_sources_implying_different_levels_becomes_review_item(self) -> None:
        nodes = [
            {"id": "1", "dests_down": ["3"], "dests_right": []},
            {"id": "2", "dests_down": [], "dests_right": ["3"]},
            {"id": "3", "dests_down": [], "dests_right": []},
        ]
        # id=1 の列0 → id=3 へ継承で0を推測。id=2 の列0 → id=3 へ +1 で1を推測。矛盾。
        raw_levels = {"1": 0, "2": 0, "3": None}
        result = compute_level_proposals(nodes, raw_levels, mode="blank_only")
        self.assertFalse(result.cancelled)
        self.assertEqual(result.proposals, [])
        review_ids = {item.node_id for item in result.needs_review}
        self.assertEqual(review_ids, {"3"})
        self.assertIn("矛盾", result.needs_review[0].reason)


class StopEventTests(unittest.TestCase):
    def test_preset_stop_event_cancels_without_computing(self) -> None:
        nodes = [{"id": "1", "dests_down": [], "dests_right": []}]
        raw_levels = {"1": None}
        stop_event = threading.Event()
        stop_event.set()
        result = compute_level_proposals(nodes, raw_levels, mode="blank_only", stop_event=stop_event)
        self.assertTrue(result.cancelled)
        self.assertEqual(result.proposals, [])


class BlankOnlyVsFullRecalcTests(unittest.TestCase):
    def test_blank_only_never_touches_existing_values(self) -> None:
        nodes = [
            {"id": "1", "dests_down": ["2"], "dests_right": []},
            {"id": "2", "dests_down": [], "dests_right": []},
        ]
        # id=1(列0)は id=2 へ継承のはずが id=2 は既に列5（矛盾する既存値）
        raw_levels = {"1": 0, "2": 5}
        result = compute_level_proposals(nodes, raw_levels, mode="blank_only")
        self.assertEqual(result.proposals, [])
        self.assertEqual(result.needs_review, [])

    def test_full_recalc_proposes_overwrite_when_source_disagrees(self) -> None:
        nodes = [
            {"id": "1", "dests_down": ["2"], "dests_right": []},
            {"id": "2", "dests_down": [], "dests_right": []},
        ]
        raw_levels = {"1": 0, "2": 5}
        result = compute_level_proposals(nodes, raw_levels, mode="full_recalc")
        proposed = {p.node_id: (p.current, p.proposed) for p in result.proposals}
        self.assertEqual(proposed.get("2"), (5, 0))

    def test_full_recalc_leaves_entry_row_untouched(self) -> None:
        # id=1 は誰からも接続先として指されない（エントリ行）ため推測規則が発火しない。
        nodes = [
            {"id": "1", "dests_down": ["2"], "dests_right": []},
            {"id": "2", "dests_down": [], "dests_right": []},
        ]
        raw_levels = {"1": 3, "2": None}
        result = compute_level_proposals(nodes, raw_levels, mode="full_recalc")
        proposed_ids = {p.node_id for p in result.proposals}
        self.assertNotIn("1", proposed_ids)


if __name__ == "__main__":
    unittest.main()
