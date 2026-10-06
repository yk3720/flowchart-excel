"""C-2「更新」書き込み結果から次のUIアクションを決める純粋関数（構想設計§3）。

副作用（再計算の実行・COM呼び出し・`request_id`管理・スレッド起動等）は
呼び出し側（`app/ui/embedded_preview.py`）に残し、本モジュールは
「書き込み結果＋現在のリトライ回数 → 次に取るべきアクション」の判定のみを担う。
"""
from __future__ import annotations

from typing import Literal

from app.core.level_writer import WriteResult

ProposalDecision = Literal["push_result", "recompute", "give_up"]


def decide_proposal_next_action(
    write_result: WriteResult,
    current_retries: int,
    *,
    max_retries: int,
) -> tuple[ProposalDecision, int]:
    """書き込み結果とリトライ回数から、次に取るべきアクションを決める。

    - `stale_topology=False`（通常完了・エラー含む）→ `"push_result"`
      （リトライ回数は0に戻す）
    - `stale_topology=True`かつリトライ上限超過 → `"give_up"`
      （リトライ回数は上限超のまま返す）
    - `stale_topology=True`かつリトライ上限内 → `"recompute"`
      （リトライ回数を1増やして返す）
    """
    if not write_result.stale_topology:
        return "push_result", 0

    new_retries = current_retries + 1
    if new_retries > max_retries:
        return "give_up", new_retries
    return "recompute", new_retries
