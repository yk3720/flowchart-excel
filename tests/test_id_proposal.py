"""compute_id_proposals のユニットテスト（構想設計§6-2、F5: ID自動採番の提案）。"""
import unittest

from app.core.id_proposal import compute_id_proposals

# table-10col-v2: ID, 種別, 色, 下, 右, 段, 列, Text1, Text2, Text3
_ROW_A = ("10", "端子", "", "20", "", 0, 0, "start", "", "")
_ROW_B = ("20", "処理", "", "", "", 1, 5, "step", "", "")


def _new_row(**overrides):
    row = ["", "", "", "", "", None, None, "", "", ""]
    for key, value in overrides.items():
        idx = {"id": 0, "type": 1, "color": 2, "down": 3, "right": 4}[key]
        row[idx] = value
    return tuple(row)


class ComputeIdProposalsTests(unittest.TestCase):
    def test_no_blank_id_rows_returns_empty(self) -> None:
        result = compute_id_proposals((_ROW_A, _ROW_B))
        self.assertEqual(result.proposals, ())
        self.assertEqual(result.dangling_warnings, ())

    def test_single_new_row_proposes_max_plus_one(self) -> None:
        new_row = _new_row(type="処理")
        result = compute_id_proposals((_ROW_A, _ROW_B, new_row))
        self.assertEqual(len(result.proposals), 1)
        self.assertEqual(result.proposals[0].row_index, 2)
        self.assertEqual(result.proposals[0].proposed_id, "21")

    def test_fully_blank_row_is_not_proposed(self) -> None:
        blank_row = _new_row()  # ID以外も全て空欄
        result = compute_id_proposals((_ROW_A, _ROW_B, blank_row))
        self.assertEqual(result.proposals, ())

    def test_multiple_new_rows_get_sequential_non_duplicate_ids_in_row_order(self) -> None:
        row1 = _new_row(type="処理")
        row2 = _new_row(type="判断")
        result = compute_id_proposals((_ROW_A, _ROW_B, row1, row2))
        self.assertEqual([p.proposed_id for p in result.proposals], ["21", "22"])
        self.assertEqual([p.row_index for p in result.proposals], [2, 3])

    def test_no_existing_ids_starts_from_one(self) -> None:
        new_row = _new_row(type="端子")
        result = compute_id_proposals((new_row,))
        self.assertEqual(result.proposals[0].proposed_id, "1")

    def test_dangling_reference_matching_new_id_warns(self) -> None:
        # 既存行が参照する"30"というIDは現存しない（削除済み）。次の採番が30に当たる。
        dangling_ref_row = ("10", "端子", "", "30", "", 0, 0, "", "", "")
        existing_max_29 = ("29", "処理", "", "", "", 1, 5, "", "", "")
        new_row = _new_row(type="判断")
        result = compute_id_proposals((dangling_ref_row, existing_max_29, new_row))
        self.assertEqual(result.proposals[0].proposed_id, "30")
        self.assertEqual(len(result.dangling_warnings), 1)
        self.assertIn("30", result.dangling_warnings[0])

    def test_dangling_reference_not_matching_new_id_does_not_warn(self) -> None:
        dangling_ref_row = ("10", "端子", "", "999", "", 0, 0, "", "", "")
        new_row = _new_row(type="判断")
        result = compute_id_proposals((dangling_ref_row, new_row))
        self.assertEqual(result.dangling_warnings, ())

    def test_non_numeric_existing_id_ignored_for_max_computation(self) -> None:
        odd_row = ("abc", "端子", "", "", "", 0, 0, "", "", "")
        new_row = _new_row(type="判断")
        result = compute_id_proposals((odd_row, new_row))
        self.assertEqual(result.proposals[0].proposed_id, "1")


if __name__ == "__main__":
    unittest.main()
