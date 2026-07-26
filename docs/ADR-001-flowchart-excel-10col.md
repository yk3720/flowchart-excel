# ADR-001: flowchart-excel 新設（10列表駆動）

**日付:** 2026-07-21  
**状態:** Accepted

## 背景

- flowchart-studio（Web）は `table-10col-v2` を正本とする
- 旧 MZ0000_FlowchartTool rev014 は 8列 + 行順 Y 配置
- Excel AutoShape 出力ニーズは残る

## 決定

1. `yk-application/flowchart-excel` を **MZ0000 rev014 の fork** として新設
2. パース · レイアウトを flowchart-studio の `parseTable.ts` / `layoutGrid.ts` に合わせる
3. 8列は後方互換として維持
4. JSON 直接取込は **スコープ外**（Phase 2 候補）

## 列 SSOT

`flowchart-studio/lib/flowchart/table/tableColumns.ts` — `TABLE_HEADERS_10_V2`
