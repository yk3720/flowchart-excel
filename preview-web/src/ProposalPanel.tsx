import { useCallback, useEffect, useRef, useState } from "react";

import type {
  LevelProposal,
  ProposalMode,
  ProposalResultPayload,
  ProposalScope,
  ProposalUpdateResultPayload,
  ReviewItem,
} from "./vite-env";

function makeRequestId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function sendProposalAction(
  scope: ProposalScope,
  action: "compute" | "update" | "cancel",
  mode: ProposalMode,
): string {
  const requestId = makeRequestId();
  window.__proposalAction = { requestId, scope, action, mode };
  return requestId;
}

type Status = "idle" | "computing" | "ready" | "updating" | "error";

export type ProposalPanelProps = {
  /** C-3（段の再採番）が同じパネルを流用するときに切り替える。 */
  scope?: ProposalScope;
  /** 表示ラベル（列 / 段）。C-3 で "段" を渡す想定。 */
  targetLabel?: string;
};

export function ProposalPanel({ scope = "level", targetLabel = "列" }: ProposalPanelProps) {
  const [mode, setMode] = useState<ProposalMode>("blank_only");
  const [status, setStatus] = useState<Status>("idle");
  const [proposals, setProposals] = useState<LevelProposal[]>([]);
  const [needsReview, setNeedsReview] = useState<ReviewItem[]>([]);
  const [skippedMultiDest, setSkippedMultiDest] = useState<string[]>([]);
  const [message, setMessage] = useState<string | null>(null);
  const [errorText, setErrorText] = useState<string | null>(null);
  const pendingRequestId = useRef<string | null>(null);

  useEffect(() => {
    window.setProposalResult = (result: ProposalResultPayload) => {
      if (result.requestId !== pendingRequestId.current) return;
      if (result.cancelled) {
        setStatus("idle");
        setMessage("計算をキャンセルしました。");
        return;
      }
      if (result.error) {
        setStatus("error");
        setErrorText(result.error);
        return;
      }
      setProposals(result.proposals ?? []);
      setNeedsReview(result.needsReview ?? []);
      setSkippedMultiDest(result.skippedMultiDest ?? []);
      setErrorText(null);
      setStatus("ready");
    };
    window.setProposalUpdateResult = (result: ProposalUpdateResultPayload) => {
      if (result.requestId !== pendingRequestId.current) return;
      if (result.staleTopology) {
        setMessage(
          "Excelの表が変更されたため提案を再計算しました。内容を確認してから改めて更新してください。",
        );
        setStatus("ready");
        return;
      }
      if (!result.ok) {
        setStatus("error");
        setErrorText(result.error ?? "更新に失敗しました。");
        return;
      }
      const excludedNote =
        result.excludedCount > 0
          ? `（${result.excludedCount}件は提案計算後に値が入力されていたため対象外でした）`
          : "";
      setMessage(`${result.updatedCount}件更新しました${excludedNote}`);
      setStatus("idle");
      setProposals([]);
      setNeedsReview([]);
    };
    return () => {
      delete window.setProposalResult;
      delete window.setProposalUpdateResult;
    };
  }, []);

  const handleCompute = useCallback(() => {
    setStatus("computing");
    setMessage(null);
    setErrorText(null);
    pendingRequestId.current = sendProposalAction(scope, "compute", mode);
  }, [scope, mode]);

  const handleUpdate = useCallback(() => {
    setStatus("updating");
    setMessage(null);
    setErrorText(null);
    pendingRequestId.current = sendProposalAction(scope, "update", mode);
  }, [scope, mode]);

  const handleCancel = useCallback(() => {
    sendProposalAction(scope, "cancel", mode);
  }, [scope, mode]);

  const isBusy = status === "computing" || status === "updating";
  const isFullRecalc = mode === "full_recalc";

  return (
    <div className="flex h-full flex-col overflow-hidden bg-flow-surface-muted">
      <div className="shrink-0 border-b border-flow-border bg-flow-surface px-4 py-3">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="text-sm font-bold text-flow-text">{targetLabel}の提案・更新</div>
          <label className="flex items-center gap-1.5 text-xs text-flow-text-muted">
            <input
              type="checkbox"
              checked={isFullRecalc}
              onChange={(e) => setMode(e.target.checked ? "full_recalc" : "blank_only")}
            />
            全部再計算（既存値も上書き対象にします）
          </label>
        </div>
        <p className="mt-1 text-xs text-flow-text-muted">
          「提案の計算」は Excel のセルを一切変更しません。内容を確認し「更新」を押した場合だけ書き込まれます。
        </p>
      </div>

      <div className="flex-1 overflow-auto px-4 py-3">
        {isBusy ? (
          <div className="mb-3 flex items-center gap-2 text-sm text-flow-text-muted">
            <span>{status === "computing" ? "計算中…" : "更新中…"}</span>
            <button
              type="button"
              className="rounded border border-flow-border px-2 py-0.5 text-xs hover:bg-flow-surface-subtle"
              onClick={handleCancel}
            >
              キャンセル
            </button>
          </div>
        ) : null}

        {errorText ? (
          <div className="mb-3 rounded-md border border-flow-danger-border bg-flow-danger-muted px-3 py-2 text-xs text-flow-danger-text">
            {errorText}
          </div>
        ) : null}

        {message ? (
          <div className="mb-3 rounded-md border border-flow-accent-muted-border bg-flow-accent-muted px-3 py-2 text-xs text-flow-accent-muted-text">
            {message}
          </div>
        ) : null}

        {isFullRecalc ? (
          <div className="mb-3 rounded-md border border-flow-warning-border-strong bg-flow-warning-bg px-3 py-2 text-xs text-flow-warning-text">
            全部再計算モードでは、手入力した値も上書き対象になります。
          </div>
        ) : null}

        {proposals.length > 0 ? (
          <table className="w-full border-collapse text-xs">
            <thead>
              <tr className="border-b border-flow-border text-left text-flow-text-muted">
                <th className="py-1 pr-2 font-medium">ID</th>
                <th className="py-1 pr-2 font-medium">現在値</th>
                <th className="py-1 pr-2 font-medium">提案値</th>
                <th className="py-1 font-medium">根拠</th>
              </tr>
            </thead>
            <tbody>
              {proposals.map((p) => (
                <tr key={p.nodeId} className="border-b border-flow-border/50">
                  <td className="py-1 pr-2 text-flow-text-body">{p.nodeId}</td>
                  <td className="py-1 pr-2 text-flow-text-body">
                    {p.current ?? "（空欄）"}
                  </td>
                  <td className="py-1 pr-2 font-medium text-flow-accent">{p.proposed}</td>
                  <td className="py-1 text-flow-text-muted">{p.reason}</td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : status === "ready" ? (
          <div className="text-xs text-flow-text-muted">提案はありません。</div>
        ) : null}

        {needsReview.length > 0 ? (
          <div className="mt-4">
            <div className="mb-1 text-xs font-bold text-flow-text">要確認</div>
            <ul className="space-y-1 text-xs text-flow-text-muted">
              {needsReview.map((item) => (
                <li key={item.nodeId}>
                  行{item.nodeId}: {item.reason}
                </li>
              ))}
            </ul>
          </div>
        ) : null}

        {skippedMultiDest.length > 0 ? (
          <div className="mt-4 text-xs text-flow-text-muted">
            複数宛先の行（対象外・未補完のまま）: {skippedMultiDest.join(", ")}
          </div>
        ) : null}

        {proposals.length > 0 ? (
          <div className="mt-3 text-xs text-flow-text-muted">
            この提案が不要な場合は、Excel で直接値を入力すると次回の提案から除外されます。
          </div>
        ) : null}
      </div>

      <div className="flex shrink-0 items-center justify-end gap-2 border-t border-flow-border bg-flow-surface px-4 py-3">
        <button
          type="button"
          className="rounded-md border border-flow-border bg-flow-surface px-4 py-2 text-sm text-flow-text-body hover:bg-flow-surface-subtle disabled:opacity-40"
          disabled={isBusy}
          onClick={handleCompute}
        >
          提案の計算
        </button>
        <button
          type="button"
          className="rounded-md bg-flow-accent px-4 py-2 text-sm font-bold text-white hover:bg-flow-accent-hover disabled:opacity-40"
          disabled={isBusy || proposals.length === 0}
          onClick={handleUpdate}
        >
          更新
        </button>
      </div>
    </div>
  );
}
