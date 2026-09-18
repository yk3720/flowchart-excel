/// <reference types="vite/client" />

export type PreviewPayload = {
  title: string;
  isFullMode: boolean;
  schema?: string;
  table: (string | number | null)[][];
  layout: {
    width: number;
    heightMin: number;
    gapV: number;
    gapH: number;
    baseLeft: number;
    baseTop: number;
  };
  meta?: {
    nodeCount?: number;
    colCount?: number;
    live?: boolean;
    embedded?: boolean;
    watch?: Record<string, unknown>;
  };
};

export type PreviewHostApi = {
  confirm: () => Promise<void> | void;
  cancel: () => Promise<void> | void;
};

export type FlowchartValidationState = {
  ok: boolean;
  errorCount: number;
};

// C-2: 「提案」タブ（段・列の自動計算・提案機能）の JS↔Python ポーリングブリッジ型。
// scope は "level"（C-2・実装済み）/ "tier"（C-3・未実装）を多重化できるよう最初から持つ。
export type ProposalScope = "level" | "tier";
export type ProposalMode = "blank_only" | "full_recalc";
export type ProposalActionKind = "compute" | "update" | "cancel";

export type ProposalActionRequest = {
  requestId: string;
  scope: ProposalScope;
  action: ProposalActionKind;
  mode: ProposalMode;
};

export type LevelProposal = {
  nodeId: string;
  current: number | null;
  proposed: number;
  reason: string;
};

export type ReviewItem = {
  nodeId: string;
  reason: string;
};

export type ProposalResultPayload = {
  requestId: string;
  cancelled?: boolean;
  error?: string;
  proposals?: LevelProposal[];
  needsReview?: ReviewItem[];
  skippedMultiDest?: string[];
};

export type ProposalUpdateResultPayload = {
  requestId: string;
  ok: boolean;
  updatedCount: number;
  excludedCount: number;
  staleTopology: boolean;
  error: string | null;
};

declare global {
  interface Window {
    __PREVIEW_PAYLOAD__?: PreviewPayload;
    setPreviewPayload?: (payload: PreviewPayload | null) => void;
    pywebview?: { api: PreviewHostApi };
    /** 埋め込み1窓モードのみ: Python側がevaluate_jsでポーリングして読む。 */
    __flowchartValidation?: FlowchartValidationState | null;
    /** C-2: React → Python への要求。Python側が ExecuteScriptAsync でポーリングして読む。 */
    __proposalAction?: ProposalActionRequest | null;
    /** C-2: Python → React への「提案の計算」結果 push（ExecuteScriptAsync 経由で呼ばれる）。 */
    setProposalResult?: (result: ProposalResultPayload) => void;
    /** C-2: Python → React への「更新」結果 push。 */
    setProposalUpdateResult?: (result: ProposalUpdateResultPayload) => void;
  }
}

export {};
