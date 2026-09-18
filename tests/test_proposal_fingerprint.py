"""proposal_fingerprint のユニットテスト（C-2 鮮度チェック・区別①②③）。"""
import unittest

from app.core.proposal_fingerprint import capture_snapshot, hand_edited_ids, topology_changed

# table-10col-v2: ID, 種別, 色, 下, 右, 段, 列, Text1, Text2, Text3
_ROW_A = ("10", "端子", "", "20", "", 0, 0, "start", "", "")
_ROW_B = ("20", "処理", "", "", "", 1, 0, "step", "", "")


class TopologyChangedTests(unittest.TestCase):
    def test_no_change_is_false(self) -> None:
        old = capture_snapshot((_ROW_A, _ROW_B))
        new = capture_snapshot((_ROW_A, _ROW_B))
        self.assertFalse(topology_changed(old, new))

    def test_dest_down_text_change_is_topology_change(self) -> None:
        old = capture_snapshot((_ROW_A, _ROW_B))
        changed_row_a = ("10", "端子", "", "30", "", 0, 0, "start", "", "")
        new = capture_snapshot((changed_row_a, _ROW_B))
        self.assertTrue(topology_changed(old, new))

    def test_row_added_is_topology_change(self) -> None:
        old = capture_snapshot((_ROW_A, _ROW_B))
        new_row = ("30", "端子", "", "", "", 2, 0, "end", "", "")
        new = capture_snapshot((_ROW_A, _ROW_B, new_row))
        self.assertTrue(topology_changed(old, new))

    def test_row_removed_is_topology_change(self) -> None:
        old = capture_snapshot((_ROW_A, _ROW_B))
        new = capture_snapshot((_ROW_A,))
        self.assertTrue(topology_changed(old, new))

    def test_text_column_change_is_not_topology_change(self) -> None:
        old = capture_snapshot((_ROW_A, _ROW_B))
        changed_row_a = ("10", "端子", "", "20", "", 0, 0, "別のテキスト", "", "")
        new = capture_snapshot((changed_row_a, _ROW_B))
        self.assertFalse(topology_changed(old, new))

    def test_target_cell_change_alone_is_not_topology_change(self) -> None:
        old = capture_snapshot((_ROW_A, _ROW_B))
        changed_row_b = ("20", "処理", "", "", "", 1, 9, "step", "", "")
        new = capture_snapshot((_ROW_A, changed_row_b))
        self.assertFalse(topology_changed(old, new))


class HandEditedIdsTests(unittest.TestCase):
    def test_target_cell_edit_is_detected(self) -> None:
        old = capture_snapshot((_ROW_A, _ROW_B))
        changed_row_b = ("20", "処理", "", "", "", 1, 9, "step", "", "")
        new = capture_snapshot((_ROW_A, changed_row_b))
        self.assertEqual(hand_edited_ids(old, new), frozenset({"20"}))

    def test_no_edit_is_empty(self) -> None:
        old = capture_snapshot((_ROW_A, _ROW_B))
        new = capture_snapshot((_ROW_A, _ROW_B))
        self.assertEqual(hand_edited_ids(old, new), frozenset())

    def test_unrelated_text_column_change_is_not_hand_edit(self) -> None:
        old = capture_snapshot((_ROW_A, _ROW_B))
        changed_row_a = ("10", "端子", "", "20", "", 0, 0, "別のテキスト", "", "")
        new = capture_snapshot((changed_row_a, _ROW_B))
        self.assertEqual(hand_edited_ids(old, new), frozenset())


if __name__ == "__main__":
    unittest.main()
