Attribute VB_Name = "modKatashikiCopy"
Option Explicit

' ============================================================
' modKatashikiCopy - 型式・入線本数コピー マクロ
' ============================================================
' Alt+F8 → 「型式・入線本数コピー」を選んで手動実行する。
'
' 処理概要:
'   1. 入力シートのC列でグループ化（同一C値 = 1グループ）
'   2. 全グループについてE/F/G列の値が一致しているか事前チェック
'   3. 不一致が1件でもあれば一切変更せず、不一致C値を表示して終了
'   4. 全グループ一致の場合のみ、各グループのマスター行（最初の行）の
'      E/F/G/X値とJ列背景色を同グループ全行に統一
'
' 変更対象: E/F/G/X列の値、J列の背景色(Interior)のみ
' それ以外の列・書式は絶対に変更しない
' ============================================================

Private Const TARGET_SHEET_NAME As String = "入力シート"
Private Const MAX_SHOW As Long = 200

' ============================================================
' メインマクロ（Alt+F8 から実行）
' ============================================================
Public Sub 型式・入線本数コピー()
    KatashikiCopyMain
End Sub

' --- 英名ラッパー（環境互換用） ---
Public Sub Katashiki_Nyusen_Copy()
    KatashikiCopyMain
End Sub

' ============================================================
' 本体処理
' ============================================================
Private Sub KatashikiCopyMain()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim calcMode As XlCalculation
    Dim r As Long
    Dim cVal As String

    ' --- シート取得 ---
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "シート「" & TARGET_SHEET_NAME & "」が見つかりません。", _
               vbExclamation, "型式・入線本数コピー"
        Exit Sub
    End If

    ' --- 最終行取得 ---
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "対象データがありません（2行目以降が空）。", _
               vbInformation, "型式・入線本数コピー"
        Exit Sub
    End If

    ' --- パフォーマンス設定 ---
    On Error GoTo ErrorHandler
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' ============================================================
    ' フェーズ1: C値でグループ化（Dictionary: C値 → Collection of 行番号）
    ' ============================================================
    Dim dictGroups As Object
    Set dictGroups = CreateObject("Scripting.Dictionary")
    dictGroups.CompareMode = vbTextCompare

    Dim arrC As Variant
    arrC = ws.Range("C2:C" & lastRow).Value

    For r = 1 To UBound(arrC, 1)
        cVal = NormalizeValue(arrC(r, 1))
        If Len(cVal) = 0 Then GoTo NextRow1

        If Not dictGroups.Exists(cVal) Then
            Set dictGroups(cVal) = New Collection
        End If
        dictGroups(cVal).Add r + 1  ' 実際の行番号（2行目起点）

NextRow1:
    Next r

    If dictGroups.Count = 0 Then
        GoTo CleanUpNoChange
    End If

    ' ============================================================
    ' フェーズ2: 不一致チェック（読み取りのみ、絶対に書き換えない）
    ' ============================================================
    Dim mismatchList As Object  ' Dictionary（重複排除用）
    Set mismatchList = CreateObject("Scripting.Dictionary")

    ' E/F/G列を配列で一括読み込み（読み取り専用）
    Dim arrE As Variant: arrE = ws.Range("E2:E" & lastRow).Value
    Dim arrF As Variant: arrF = ws.Range("F2:F" & lastRow).Value
    Dim arrG As Variant: arrG = ws.Range("G2:G" & lastRow).Value

    Dim groupKey As Variant
    Dim rowCol As Collection
    Dim masterRow As Long
    Dim masterE As String, masterF As String, masterG As String
    Dim curE As String, curF As String, curG As String
    Dim idx As Long
    Dim rn As Long
    Dim hasMismatch As Boolean

    For Each groupKey In dictGroups.Keys
        Set rowCol = dictGroups(groupKey)
        If rowCol.Count < 2 Then GoTo NextGroup2  ' 1行グループは不一致なし

        ' マスター行（グループ最初の行）の値を取得
        masterRow = rowCol(1)
        masterE = NormalizeValue(arrE(masterRow - 1, 1))
        masterF = NormalizeValue(arrF(masterRow - 1, 1))
        masterG = NormalizeValue(arrG(masterRow - 1, 1))

        ' グループ内の他の行と比較
        hasMismatch = False
        For idx = 2 To rowCol.Count
            rn = rowCol(idx)
            curE = NormalizeValue(arrE(rn - 1, 1))
            curF = NormalizeValue(arrF(rn - 1, 1))
            curG = NormalizeValue(arrG(rn - 1, 1))

            If curE <> masterE Or curF <> masterF Or curG <> masterG Then
                hasMismatch = True
                Exit For
            End If
        Next idx

        If hasMismatch Then
            mismatchList(CStr(groupKey)) = True
        End If

NextGroup2:
    Next groupKey

    ' ============================================================
    ' フェーズ2.5: 不一致があれば即終了（一切変更しない）
    ' ============================================================
    If mismatchList.Count > 0 Then
        ' 復帰してからメッセージ表示
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = calcMode

        ' 不一致C値のメッセージ構築
        Dim msg As String
        Dim showCount As Long
        Dim mmKey As Variant
        Dim mmKeys As Variant

        mmKeys = mismatchList.Keys
        showCount = mismatchList.Count
        If showCount > MAX_SHOW Then showCount = MAX_SHOW

        msg = "EFG列の値が一致しないC値があります。" & vbCrLf & _
              "処理を中止しました。" & vbCrLf

        Dim cnt As Long
        cnt = 0
        For Each mmKey In mmKeys
            cnt = cnt + 1
            If cnt > MAX_SHOW Then Exit For
            msg = msg & vbCrLf & CStr(mmKey)
        Next mmKey

        If mismatchList.Count > MAX_SHOW Then
            msg = msg & vbCrLf & vbCrLf & _
                  "（ほか" & (mismatchList.Count - MAX_SHOW) & "件）"
        End If

        MsgBox msg, vbExclamation, "型式・入線本数コピー"
        Exit Sub
    End If

    ' ============================================================
    ' フェーズ3: 全グループ一致 → マスター行の値で統一（ここで初めて書き換える）
    ' ============================================================
    Dim masterX As String
    Dim curX As String
    Dim arrX As Variant: arrX = ws.Range("X2:X" & lastRow).Value

    ' J列背景色のマスター情報を一時保持する型
    Dim mPattern As Long
    Dim mColor As Long
    Dim mColorIndex As Long
    Dim mPatternColor As Long
    Dim mPatternColorIndex As Long
    Dim mTintAndShade As Double

    For Each groupKey In dictGroups.Keys
        Set rowCol = dictGroups(groupKey)
        If rowCol.Count < 2 Then GoTo NextGroup3  ' 1行なら更新不要

        masterRow = rowCol(1)

        ' マスター行の E/F/G/X 値
        masterE = NormalizeValue(arrE(masterRow - 1, 1))
        masterF = NormalizeValue(arrF(masterRow - 1, 1))
        masterG = NormalizeValue(arrG(masterRow - 1, 1))
        masterX = NormalizeValue(arrX(masterRow - 1, 1))

        ' マスター行の J列 Interior を読み取り（値を変数にコピー、セルは触らない）
        With ws.Cells(masterRow, "J").Interior
            mPattern = .Pattern
            mColor = .Color
            mColorIndex = .ColorIndex
            mPatternColor = .PatternColor
            mPatternColorIndex = .PatternColorIndex
            On Error Resume Next
            mTintAndShade = .TintAndShade
            If Err.Number <> 0 Then mTintAndShade = 0
            On Error GoTo ErrorHandler
        End With

        ' グループ内の他の行を更新
        For idx = 2 To rowCol.Count
            rn = rowCol(idx)

            ' E/F/G/X列の値を更新
            ws.Cells(rn, "E").Value = ws.Cells(masterRow, "E").Value
            ws.Cells(rn, "F").Value = ws.Cells(masterRow, "F").Value
            ws.Cells(rn, "G").Value = ws.Cells(masterRow, "G").Value
            ws.Cells(rn, "X").Value = ws.Cells(masterRow, "X").Value

            ' J列の背景色を更新
            CopyInterior ws.Cells(masterRow, "J"), ws.Cells(rn, "J")
        Next idx

NextGroup3:
    Next groupKey

    ' ============================================================
    ' 完了
    ' ============================================================
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = calcMode

    MsgBox "完了しました。", vbInformation, "型式・入線本数コピー"
    Exit Sub

CleanUpNoChange:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = calcMode
    MsgBox "対象となるグループがありません。", vbInformation, "型式・入線本数コピー"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error Resume Next
    Application.Calculation = calcMode
    On Error GoTo 0
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation, "型式・入線本数コピー"
End Sub

' ============================================================
' 値の正規化（Value2 → CStr → Trim）
' Null/Empty/Error → ""、それ以外は CStr + Trim
' ============================================================
Private Function NormalizeValue(ByVal v As Variant) As String
    If IsError(v) Then
        NormalizeValue = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        NormalizeValue = ""
    Else
        NormalizeValue = Trim(CStr(v))
    End If
End Function

' ============================================================
' Interior のコピー（srcセル → dstセル）
' 塗りつぶし無し(xlNone)も正確に反映する
' ============================================================
Private Sub CopyInterior(ByVal src As Range, ByVal dst As Range)
    With src.Interior
        If .Pattern = xlNone Then
            ' 塗りつぶし無し
            dst.Interior.Pattern = xlNone
        Else
            dst.Interior.Pattern = .Pattern
            dst.Interior.Color = .Color
            dst.Interior.PatternColor = .PatternColor

            On Error Resume Next
            dst.Interior.TintAndShade = .TintAndShade
            On Error GoTo 0
        End If
    End With
End Sub
