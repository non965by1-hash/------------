# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

Excel VBA マクロの開発プロジェクト。`シーケンス入力シート.xlsm`（ブックA）に VBA マクロを実装し、`parts database.xlsm`（ブックB）から部品情報を検索・入力補助する機能を作る。

## 対象ファイル

- **`シーケンス入力シート.xlsm`** — マクロを実装するブック（ブックA）。ファイル名は変わる可能性があるため、コード内では必ず `ThisWorkbook` を基準に参照する。対象シート名は固定で **「入力シート」**。
- **`parts database.xlsm`** — 部品データベース（ブックB）。ファイル名固定。社内ネットワーク上の所定パスに配置予定。**読み取り専用で開くこと（ReadOnly:=True）**。

## 実装するマクロの仕様

### 発火タイミング
- ブックA「入力シート」の **G列（2行目以降）** に文字が入力されたタイミング（`Worksheet_Change`）
- 空文字になった場合は何もしない（E/F/G/J列も変えない）
- 無限ループ防止フラグ（`g_bIgnoreChange` 等）を必ず実装する（G列に書き戻す処理があるため）

### 検索仕様
- ブックBの**全シートのC列（2行目以降）**を検索対象とする
- 毎回全走査しない。**Workbook_Open または初回起動時にインデックスを作成**（`Scripting.Dictionary` 推奨）
- インデックスが保持する情報：SheetName, Row, A列値, B列値, C列値, D列値, E列値
- 候補の並び順：**グループ1（完全一致）→ グループ2（部分一致）**
  - グループ内は一致度スコア順（前方一致 > マッチ位置が早い > 長さ差が小さい > C列昇順）
  - 最大表示件数は **50件**（超える場合は「候補が多いので絞ってください」を案内）

### UserForm（frmPartsPick）
- キーボードのみで操作できること
- コントロール構成：
  - `txtFilter`（TextBox）— 初期値=入力キー。可能なら文字入力に応じてリスト絞り込み（`txtFilter_Change`）
  - `lstCandidates`（ListBox、3列）— 表示列の順番：**C列 / D列 / A列**
  - `lblHint`（Label）— 操作説明（↑↓選択、Enter決定、Escキャンセル）
  - `btnOK`（`Default=True`、Enterで決定）
  - `btnCancel`（`Cancel=True`、Escでキャンセル）
- **候補が1件のみでもフォームを表示する**（自動確定しない）
- 候補0件の場合：「見つかりません」表示 → Escで安全に終了
- Escキャンセル時は何も書き込まない

### 転記処理（決定時）
ブックB選択行 → ブックA「入力シート」の編集した行へ：
- DB A列 → ブックA **E列**
- DB B列 → ブックA **F列**
- DB C列 → ブックA **G列**

さらに、ブックAの**C列**に今回確定したDB C列値と同じ文字が入っている**全行**に対しても同様にE/F/Gを自動入力する。

### J列の背景色ルール（B案：ON/OFFする）
DBの同一行のE列値を参照し、ブックA「入力シート」の対象行J列の背景を制御する：
- **E列が「1」とみなす条件**：文字列化→Trim→全角スペースも含む全空白除去後、結果が "1" なら1扱い
- E列が1扱い → `Interior.Color = RGB(255, 0, 0)`（赤）
- E列が1扱いでない → `Interior.Pattern = xlNone`（塗りつぶし無しに戻す）
- このON/OFFは「編集した行」と「C列一致で一括反映した行」の**両方**に適用する
- J列の書式変更はこのルールのみ許可。**他セルの書式・値は変更禁止**

### DBファイルのパス（ネットワーク対応）
```vba
Const PARTS_DB_FULLPATH As String = "\\server\share\folder\parts database.xlsm"
```
- `PARTS_DB_FULLPATH` が空の場合はブックAと同じフォルダからフォールバック検索
- 見つからない場合はメッセージ表示して終了（書き込みなし）
- すでに開いている場合はそのインスタンスを使う

### パフォーマンス制御
```vba
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
' ... 処理 ...
' エラー時も含め、必ず復帰する（On Error / Finally相当）
```

## 実装配置（期待するファイル構成）

| 場所 | 内容 |
|------|------|
| `ThisWorkbook` モジュール | `Workbook_Open`（インデックス初期化） |
| シート「入力シート」コード | `Worksheet_Change`（G列トリガ） |
| 標準モジュール `modPartsIndex` | インデックス生成・検索・フォーム呼び出し・転記処理 |
| UserForm `frmPartsPick` | 候補選択フォーム |

## テスト手順

1. G2に文字入力 → フォーム表示 → ↑↓移動 → Enterで決定 → E2/F2/G2が埋まる
2. Escでキャンセル → E/F/G/Jは変化なし
3. DB E列=1の行を選択 → 入力シート同行Jが赤になる
4. DB E列が1でない行を選択 → 入力シート同行Jが塗りつぶし無しになる
5. C列に同じ文字の行が複数 → それらのE/F/Gも同じ値になり、J列も適切に制御される
6. 完全一致が上、部分一致が下に表示されることを確認
7. 候補が51件以上 → 50件表示＋案内メッセージ
8. 候補0件 → エラーなくEscで抜けられる
