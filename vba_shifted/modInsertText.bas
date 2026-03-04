Attribute VB_Name = "modInsertText"
Option Explicit

' ============================================================
' modInsertText - 可視選択セルへの文字列先頭/末尾追加マクロ
' ============================================================
' ソート・フィルタ後に表示されているセルのうち、
' ユーザーが選択している範囲内の可視セルだけを対象に、
' 任意の文字列を先頭または末尾に追加する。
'
' 使い方:
'   1. シート上で対象セル範囲を選択
'   2. Alt+F8 → 「文字列追加」を実行
'   3. フォームで追加文字列・位置・オプションを設定
'   4. OK → 確認ダイアログ → 実行
' ============================================================

' --- フォーム連携用グローバル変数 ---
Public g_InsertText As String       ' 追加する文字列
Public g_InsertToHead As Boolean    ' True=先頭, False=末尾
Public g_SkipDuplicate As Boolean   ' True=二重追加スキップ
Public g_AddToEmpty As Boolean      ' True=空白セルにも追加
Public g_FormOK As Boolean          ' True=OKで閉じた

' ============================================================
' メインマクロ（Alt+F8 から実行）
' ============================================================
Public Sub 文字列追加()
    Dim visibleCells As Range
    Dim cell As Range
    Dim area As Range
    Dim calcMode As XlCalculation
    Dim processedCount As Long
    Dim skippedCount As Long
    Dim totalVisible As Long
    Dim cellValue As String
    Dim confirmMsg As String
    Dim posLabel As String

    ' --- Selection の妥当性チェック ---
    If Selection Is Nothing Then
        MsgBox "セルが選択されていません。", vbExclamation, "文字列追加"
        Exit Sub
    End If
    If Not TypeOf Selection Is Range Then
        MsgBox "セル範囲を選択してください。", vbExclamation, "文字列追加"
        Exit Sub
    End If

    ' --- 可視セルの取得 ---
    On Error Resume Next
    Set visibleCells = Selection.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If visibleCells Is Nothing Then
        MsgBox "選択範囲内に可視セルがありません。" & vbCrLf & _
               "フィルタや非表示行により全て隠れている可能性があります。", _
               vbExclamation, "文字列追加"
        Exit Sub
    End If

    ' 可視セル数をカウント
    totalVisible = visibleCells.Count

    ' --- フォーム表示 ---
    g_FormOK = False
    frmInsertText.Show vbModal

    ' キャンセルされた場合
    If Not g_FormOK Then Exit Sub

    ' 入力チェック
    If Len(g_InsertText) = 0 Then
        MsgBox "追加文字列が空です。処理を中止します。", vbExclamation, "文字列追加"
        Exit Sub
    End If

    ' --- 確認ダイアログ ---
    If g_InsertToHead Then
        posLabel = "先頭"
    Else
        posLabel = "末尾"
    End If

    confirmMsg = "以下の内容で文字列を追加します。" & vbCrLf & vbCrLf & _
                 "対象セル数: " & totalVisible & " セル" & vbCrLf & _
                 "追加文字列: 「" & g_InsertText & "」" & vbCrLf & _
                 "追加位置: " & posLabel & vbCrLf & _
                 "二重追加スキップ: " & IIf(g_SkipDuplicate, "ON", "OFF") & vbCrLf & _
                 "空白セルにも追加: " & IIf(g_AddToEmpty, "ON", "OFF") & vbCrLf & vbCrLf & _
                 "※元に戻せません。実行しますか？"

    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "文字列追加 - 確認") <> vbYes Then
        Exit Sub
    End If

    ' --- パフォーマンス設定 ---
    On Error GoTo ErrorHandler
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' --- メイン処理 ---
    processedCount = 0
    skippedCount = 0

    For Each area In visibleCells.Areas
        For Each cell In area.Cells
            ' 結合セルはスキップ（エラー回避）
            If cell.MergeCells Then
                skippedCount = skippedCount + 1
                GoTo NextCell
            End If

            ' セルの値を文字列として取得
            ' ※ .Value を CStr するため、数値・日付は内部値の文字列表現になる
            '    表示形式はそのまま残るが、値は文字列に変換される点に注意
            cellValue = SafeCellToString(cell)

            ' 空白セルの処理
            If Len(cellValue) = 0 Then
                If Not g_AddToEmpty Then
                    skippedCount = skippedCount + 1
                    GoTo NextCell
                End If
            End If

            ' 二重追加スキップ判定
            If g_SkipDuplicate And Len(cellValue) > 0 Then
                If g_InsertToHead Then
                    ' 先頭が既に追加文字列で始まっている場合スキップ
                    If Len(cellValue) >= Len(g_InsertText) Then
                        If Left(cellValue, Len(g_InsertText)) = g_InsertText Then
                            skippedCount = skippedCount + 1
                            GoTo NextCell
                        End If
                    End If
                Else
                    ' 末尾が既に追加文字列で終わっている場合スキップ
                    If Len(cellValue) >= Len(g_InsertText) Then
                        If Right(cellValue, Len(g_InsertText)) = g_InsertText Then
                            skippedCount = skippedCount + 1
                            GoTo NextCell
                        End If
                    End If
                End If
            End If

            ' 文字列追加
            If g_InsertToHead Then
                cell.Value = g_InsertText & cellValue
            Else
                cell.Value = cellValue & g_InsertText
            End If
            processedCount = processedCount + 1

NextCell:
        Next cell
    Next area

    ' --- 完了メッセージ ---
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = calcMode

    MsgBox "完了しました。" & vbCrLf & vbCrLf & _
           "処理セル数: " & processedCount & vbCrLf & _
           "スキップ数: " & skippedCount, vbInformation, "文字列追加"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error Resume Next
    Application.Calculation = calcMode
    On Error GoTo 0
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation, "文字列追加"
End Sub

' ============================================================
' セル値を安全に文字列化（エラー値・Null・Empty対応）
' ============================================================
Private Function SafeCellToString(ByVal cell As Range) As String
    Dim v As Variant
    v = cell.Value
    If IsError(v) Then
        SafeCellToString = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        SafeCellToString = ""
    Else
        SafeCellToString = CStr(v)
    End If
End Function
