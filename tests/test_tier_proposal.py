"""compute_tier_proposals のユニットテスト。"""
import unittest

from app.core.tier_proposal import compute_tier_proposals

_NODES = (
    {"id": "1", "ridx": 1},
    {"id": "2", "ridx": 2},
    {"id": "3", "ridx": 3},
)


class ComputeTierProposalsTests(unittest.TestCase):
    def test_blank_only_assigns_sequential_from_zero_when_all_blank(self) -> None:
        raw = {"1": None, "2": None, "3": None}
        result = compute_tier_proposals(_NODES, raw, mode="blank_only")
        self.assertEqual([p.proposed for p in result.proposals], [0, 1, 2])
        self.assertEqual([p.node_id for p in result.proposals], ["1", "2", "3"])

    def test_blank_only_continues_after_existing_max(self) -> None:
        raw = {"1": 0, "2": None, "3": None}
        result = compute_tier_proposals(_NODES, raw, mode="blank_only")
        self.assertEqual([p.node_id for p in result.proposals], ["2", "3"])
        self.assertEqual([p.proposed for p in result.proposals], [1, 2])

    def test_blank_only_skips_filled_rows(self) -> None:
        raw = {"1": 0, "2": 1, "3": 2}
        result = compute_tier_proposals(_NODES, raw, mode="blank_only")
        self.assertEqual(result.proposals, [])

    def test_full_recalc_ranks_unique_values_without_breaking_groups(self) -> None:
        # 段 10,10,30 → ランク 0,0,1（同じ段のグループは維持）
        raw = {"1": 10, "2": 10, "3": 30}
        result = compute_tier_proposals(_NODES, raw, mode="full_recalc")
        by_id = {p.node_id: p.proposed for p in result.proposals}
        self.assertEqual(by_id["1"], 0)
        self.assertEqual(by_id["2"], 0)
        self.assertEqual(by_id["3"], 1)

    def test_full_recalc_all_same_value_yields_no_change(self) -> None:
        raw = {"1": 0, "2": 0, "3": 0}
        result = compute_tier_proposals(_NODES, raw, mode="full_recalc")
        self.assertEqual(result.proposals, [])


if __name__ == "__main__":
    unittest.main()
