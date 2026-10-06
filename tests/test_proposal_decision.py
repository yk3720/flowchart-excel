"""decide_proposal_next_action のユニットテスト（構想設計§3の純粋関数抽出）。"""
import unittest

from app.core.level_writer import WriteResult
from app.core.proposal_decision import decide_proposal_next_action

_MAX_RETRIES = 3


class DecideProposalNextActionTests(unittest.TestCase):
    def test_success_pushes_result_and_resets_retries(self) -> None:
        result = WriteResult(ok=True, updated_count=2, excluded_count=0)
        decision, retries = decide_proposal_next_action(result, 2, max_retries=_MAX_RETRIES)
        self.assertEqual(decision, "push_result")
        self.assertEqual(retries, 0)

    def test_failure_without_stale_topology_pushes_result_and_resets_retries(self) -> None:
        """書き込み自体の失敗（COMエラー等）はstale_topologyではないため push_result 扱い。"""
        result = WriteResult(ok=False, error="更新に失敗しました: boom")
        decision, retries = decide_proposal_next_action(result, 1, max_retries=_MAX_RETRIES)
        self.assertEqual(decision, "push_result")
        self.assertEqual(retries, 0)

    def test_stale_topology_within_limit_recomputes_and_increments_retries(self) -> None:
        result = WriteResult(ok=False, stale_topology=True)
        decision, retries = decide_proposal_next_action(result, 0, max_retries=_MAX_RETRIES)
        self.assertEqual(decision, "recompute")
        self.assertEqual(retries, 1)

    def test_stale_topology_at_limit_still_recomputes(self) -> None:
        result = WriteResult(ok=False, stale_topology=True)
        decision, retries = decide_proposal_next_action(
            result, _MAX_RETRIES - 1, max_retries=_MAX_RETRIES
        )
        self.assertEqual(decision, "recompute")
        self.assertEqual(retries, _MAX_RETRIES)

    def test_stale_topology_beyond_limit_gives_up(self) -> None:
        result = WriteResult(ok=False, stale_topology=True)
        decision, retries = decide_proposal_next_action(
            result, _MAX_RETRIES, max_retries=_MAX_RETRIES
        )
        self.assertEqual(decision, "give_up")
        self.assertEqual(retries, _MAX_RETRIES + 1)


if __name__ == "__main__":
    unittest.main()
