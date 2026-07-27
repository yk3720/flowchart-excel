import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { Edge, Node } from "@xyflow/react";

import { FlowCanvas, type FlowCanvasHandle } from "@/frontend/src/components/flowchart/FlowCanvas";
import { generateFlowchart } from "@/lib/flowchart/graph/generate";
import {
  toReactFlow,
  type FlowNodeData,
} from "@/lib/flowchart/graph/toReactFlow";
import type { LayoutConfig } from "@/lib/flowchart/model/types";
import type { PreviewPayload } from "./vite-env";

function readInitialPayload(): PreviewPayload | null {
  return window.__PREVIEW_PAYLOAD__ ?? null;
}

function callHost(action: "confirm" | "cancel") {
  const api = window.pywebview?.api;
  if (api?.[action]) {
    void api[action]();
    return;
  }
  // ブラウザ単独確認用
  console.info(`[preview-web] ${action}`);
}

export function App() {
  const [payload, setPayload] = useState<PreviewPayload | null>(readInitialPayload);
  const canvasRef = useRef<FlowCanvasHandle | null>(null);
  const [zoomPercent, setZoomPercent] = useState(100);

  useEffect(() => {
    window.setPreviewPayload = (next) => setPayload(next);
    return () => {
      delete window.setPreviewPayload;
    };
  }, []);

  const generated = useMemo(() => {
    if (!payload) return null;
    const layout: LayoutConfig = {
      width: payload.layout.width,
      heightMin: payload.layout.heightMin,
      gapV: payload.layout.gapV,
      gapH: payload.layout.gapH,
      baseLeft: payload.layout.baseLeft,
      baseTop: payload.layout.baseTop,
    };
    return generateFlowchart(payload.table, layout, payload.schema);
  }, [payload]);

  const rf = useMemo(() => {
    if (!generated || !generated.ok) {
      return { nodes: [] as Node<FlowNodeData>[], edges: [] as Edge[] };
    }
    return toReactFlow(generated.placed, generated.edges);
  }, [generated]);

  const onConfirm = useCallback(() => callHost("confirm"), []);
  const onCancel = useCallback(() => callHost("cancel"), []);

  if (!payload) {
    return (
      <div className="flex h-full items-center justify-center text-flow-text-muted">
        プレビューデータを待機中…
      </div>
    );
  }

  const mode = payload.isFullMode ? "表全体" : "選択範囲";
  const errorText =
    generated && !generated.ok ? generated.errors.join(" / ") : null;
  const embedded = Boolean(payload.meta?.embedded);

  return (
    <div className="flex h-full flex-col bg-flow-surface-muted">
      <header className="shrink-0 border-b border-flow-border bg-flow-surface px-4 py-3">
        <div className="flex flex-wrap items-center gap-2">
          <div className="text-sm font-bold text-flow-text">
            {payload.title}（{mode}）
          </div>
          {payload.meta?.live ? (
            <span className="rounded border border-flow-accent/30 bg-flow-accent/10 px-2 py-0.5 text-[11px] font-medium text-flow-accent">
              ライブ更新
            </span>
          ) : null}
        </div>
        <div className="mt-1 text-xs text-flow-text-muted">
          {generated?.ok
            ? `ノード ${generated.placed.length} · 接続 ${generated.edges.length} — Excel の表を直すと自動で再描画`
            : "生成エラー — 表を直すと自動で再試行します"}
        </div>
        {errorText ? (
          <div className="mt-2 text-xs text-flow-danger">{errorText}</div>
        ) : null}
      </header>

      <div className="relative min-h-0 flex-1">
        {generated?.ok ? (
          <>
            <FlowCanvas
              canvasRef={canvasRef}
              nodes={rf.nodes}
              edges={rf.edges}
              fillContainer
              onViewportZoomChange={setZoomPercent}
            />
            <div className="pointer-events-none absolute bottom-3 right-3 flex items-center gap-1 rounded-md border border-flow-border bg-flow-surface px-2 py-1 text-xs text-flow-text-body shadow-sm">
              <button
                type="button"
                className="pointer-events-auto min-h-6 min-w-6 rounded border border-flow-border px-2"
                aria-label="縮小"
                onClick={() => canvasRef.current?.zoomOut()}
              >
                −
              </button>
              <span className="min-w-10 text-center tabular-nums">{zoomPercent}%</span>
              <button
                type="button"
                className="pointer-events-auto min-h-6 min-w-6 rounded border border-flow-border px-2"
                aria-label="拡大"
                onClick={() => canvasRef.current?.zoomIn()}
              >
                +
              </button>
              <button
                type="button"
                className="pointer-events-auto ml-1 rounded border border-flow-border px-2 py-0.5"
                onClick={() => canvasRef.current?.fitView()}
              >
                ホーム
              </button>
            </div>
          </>
        ) : (
          <div className="flex h-full items-center justify-center p-6 text-sm text-flow-text-muted">
            表示できるフローがありません
          </div>
        )}
      </div>

      {!embedded ? (
      <footer className="flex shrink-0 items-center justify-between gap-3 border-t border-flow-border bg-flow-surface px-4 py-3">
        <p className="text-xs text-flow-text-muted">
          約0.75秒ごとに Excel を再読込します（セル編集中はスキップ）。
        </p>
        <div className="flex gap-2">
          <button
            type="button"
            className="rounded-md border border-flow-border bg-flow-surface px-4 py-2 text-sm text-flow-text-body hover:bg-flow-surface-subtle"
            onClick={onCancel}
          >
            キャンセル
          </button>
          <button
            type="button"
            className="rounded-md bg-flow-accent px-4 py-2 text-sm font-bold text-white hover:bg-flow-accent-hover disabled:opacity-40"
            disabled={!generated?.ok}
            onClick={onConfirm}
          >
            Excelに作成
          </button>
        </div>
      </footer>
      ) : null}
    </div>
  );
}
