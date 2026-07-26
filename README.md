# flowchart-excel

Excel 上の **10列表** から AutoShape フローチャートを生成する Windows デスクトップツール。  
[flowchart-studio](../flowchart-studio/)（Web）と同じ `table-10col-v2` 列定義を SSOT とする。

## 位置づけ

| ツール | 役割 |
|--------|------|
| **flowchart-studio** | Web で表編集 · クラウド保存 · PNG/SVG |
| **flowchart-excel**（本リポ） | Excel AutoShape 出力 · 社内 Excel 配布向け |

元は `MZ0000_FlowchartTool_rev014`（8列）を fork し、10列（段+列）レイアウトに対応。

## 操作フロー（プレビュー必須 · studio 同等）

1. Excel で表を選択
2. 「表全体を確認して作成」または「選択範囲を確認して作成」
3. **WebView プレビュー**（flowchart-studio と同じ React Flow 描画）で確認
4. 問題なければ「Excelに作成」→ AutoShape 描画

キャンセルすれば Excel には何も書きません。スマート・パレット（単体図形）はプレビュー対象外です。

初回またはプレビュー更新後は `preview-web` のビルドが必要です。

```powershell
cd preview-web
npm install
npm run build
cd ..
pip install -r requirements.txt
python main.py
```

## 列定義（table-10col-v2）

`ID | 図形種別 | 色 | 接続先(下) | 接続先(右) | 段 | 列 | Text1 | Text2 | Text3`

- **段（tier）:** 縦位置 — 同じ段 = 同じ高さで横並び
- **列（level）:** 横位置 — 分岐の並び

8列（旧 MZ0000 形式）も後方互換で読み込めます。

## flowchart-studio との連携

1. Web で表を編集
2. 表を TSV コピー（または JSON から表部分を Excel に貼付）
3. Excel で範囲を選択
4. 本ツールで「選択範囲を確認して作成」（プレビュー必須）

JSON ファイルの直接取込は **未実装**（今後の拡張候補）。

## 必要環境

- Windows 10/11
- Microsoft Excel（COM 操作）
- Python 3.14+（開発時）

## セットアップ

```powershell
cd c:\yk-application\flowchart-excel
python setup_venv.py
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python main.py
```

## テスト

```powershell
python -m unittest discover -s tests -p "test_*.py"
```

## ビルド（exe）

```powershell
python build_exe.py
```

## 関連ドキュメント

- flowchart-studio データモデル: `../flowchart-studio/docs/03_技術仕様/データモデル.md`
- 旧 MZ0000: `c:\1.cursor\5.Python\3.作成中\MZ0000_FlowchartTool_rev014\`
