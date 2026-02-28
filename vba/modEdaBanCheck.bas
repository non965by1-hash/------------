Attribute VB_Name = "modEdaBanCheck"
Option Explicit

' ============================================================
' modEdaBanCheck - 枝番チェック マクロ
' ============================================================
' Alt+F8 → 「枝番チェック」を選んで手動実行する。
'
' 判定ロジック：
'   1. I列の値でグループ化（I列空白行はスキップ）
'   2. 各(I, A)単位で対象VコードのDistinct数を数える → vCountSum
'   3. Iグループ内のA列Distinct種類数 >= 2 なら aBonus = 1
'   4. Iグループ内のY列が「1」の行数 → yCount
'   5. total = vCountSum + aBonus + yCount >= 3 なら対象
'      → そのI値の全行のJ列を薄いブルーに塗る（既存色があれば変更しない）
' ============================================================

' --- 薄いブルーの色定数 ---
Private Const LIGHT_BLUE_R As Long = 221
Private Const LIGHT_BLUE_G As Long = 235
Private Const LIGHT_BLUE_B As Long = 247

' --- 対象Vコードリスト ---
Private Const TARGET_V_CODES As String = "FDR,FDR1,FDR2,FDL,FDL1,FDL2,BDR,BDR1,BDR2,BDL,BDL1,BDL2"

' --- 判定閾値 ---
Private Const THRESHOLD As Long = 3

' ============================================================
' メインマクロ（Alt+F8 から実行）
' ============================================================
Public Sub 枝番チェック()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim iVal As String
    Dim aVal As String
    Dim vVal As String
    Dim yVal As String
    Dim iaKey As String

    ' --- 対象Vコードを Dictionary に格納（高速判定用） ---
    Dim dictTargetV As Object
    Set dictTargetV = CreateObject("Scripting.Dictionary")
    dictTargetV.CompareMode = vbTextCompare
    Dim vCodes As Variant
    Dim vc As Variant
    vCodes = Split(TARGET_V_CODES, ",")
    For Each vc In vCodes
        dictTargetV(Trim(CStr(vc))) = True
    Next vc

    ' --- 集計用 Dictionary ---
    ' dictIAVCodes: key="I値" & Chr(0) & "A値" → Dictionary(Vコード→True)
    Dim dictIAVCodes As Object
    Set dictIAVCodes = CreateObject("Scripting.Dictionary")

    ' dictIATypes: key="I値" → Dictionary(A値→True)
    Dim dictIATypes As Object
    Set dictIATypes = CreateObject("Scripting.Dictionary")

    ' dictIYCount: key="I値" → Y列が「1」の行数
    Dim dictIYCount As Object
    Set dictIYCount = CreateObject("Scripting.Dictionary")

    ' dictIRows: key="I値" → Collection(行番号)
    Dim dictIRows As Object
    Set dictIRows = CreateObject("Scripting.Dictionary")

    ' --- シート・範囲の準備 ---
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Worksheets("入力シート")

    ' 最終行を A/I/V/Y 列の最大から決定
    lastRow = GetMaxLastRow(ws, Array("A", "I", "V", "Y"))
    If lastRow < 2 Then
        MsgBox "データがありません（2行目以降が空）。", vbInformation, "枝番チェック"
        Exit Sub
    End If

    ' パフォーマンス設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- A/I/V/Y列を配列で一括読み込み ---
    Dim arrA As Variant: arrA = ws.Range("A2:A" & lastRow).Value
    Dim arrI As Variant: arrI = ws.Range("I2:I" & lastRow).Value
    Dim arrV As Variant: arrV = ws.Range("V2:V" & lastRow).Value
    Dim arrY As Variant: arrY = ws.Range("Y2:Y" & lastRow).Value

    ' --- 集計ループ ---
    For r = 1 To UBound(arrI, 1)
        iVal = Trim(CStr(SafeVariant(arrI(r, 1))))
        If Len(iVal) = 0 Then GoTo NextRow  ' I列空白はスキップ

        aVal = Trim(CStr(SafeVariant(arrA(r, 1))))
        vVal = Trim(CStr(SafeVariant(arrV(r, 1))))
        yVal = Trim(CStr(SafeVariant(arrY(r, 1))))

        ' --- I単位の行番号リストに追加 ---
        If Not dictIRows.Exists(iVal) Then
            Set dictIRows(iVal) = New Collection
        End If
        dictIRows(iVal).Add r + 1  ' 実際の行番号（2行目開始）

        ' --- I単位の A種類を記録 ---
        If Not dictIATypes.Exists(iVal) Then
            Set dictIATypes(iVal) = CreateObject("Scripting.Dictionary")
            dictIATypes(iVal).CompareMode = vbTextCompare
        End If
        If Len(aVal) > 0 Then
            dictIATypes(iVal)(aVal) = True
        End If

        ' --- (I,A)単位の Vコード distinct を記録 ---
        If Len(vVal) > 0 And dictTargetV.Exists(vVal) Then
            iaKey = iVal & Chr(0) & aVal
            If Not dictIAVCodes.Exists(iaKey) Then
                Set dictIAVCodes(iaKey) = CreateObject("Scripting.Dictionary")
                dictIAVCodes(iaKey).CompareMode = vbTextCompare
            End If
            dictIAVCodes(iaKey)(vVal) = True
        End If

        ' --- I単位の Y列「1」カウント ---
        If IsYValueOne(yVal) Then
            If Not dictIYCount.Exists(iVal) Then
                dictIYCount(iVal) = CLng(0)
            End If
            dictIYCount(iVal) = dictIYCount(iVal) + 1
        End If

NextRow:
    Next r

    ' --- 判定＆J列着色 ---
    Dim iKey As Variant
    Dim iaKeys As Variant
    Dim iak As Variant
    Dim vCountSum As Long
    Dim aTypeCount As Long
    Dim aBonus As Long
    Dim yCount As Long
    Dim total As Long
    Dim targetICount As Long
    Dim lightBlue As Long

    lightBlue = RGB(LIGHT_BLUE_R, LIGHT_BLUE_G, LIGHT_BLUE_B)
    targetICount = 0

    For Each iKey In dictIRows.Keys
        ' --- vCountSum: (I,A)ごとの distinct Vコード数の合計 ---
        vCountSum = 0
        iaKeys = dictIAVCodes.Keys
        For Each iak In iaKeys
            If Left(CStr(iak), Len(CStr(iKey)) + 1) = CStr(iKey) & Chr(0) Then
                vCountSum = vCountSum + dictIAVCodes(iak).Count
            End If
        Next iak

        ' --- aBonus: A列種類数 >= 2 なら +1 ---
        aTypeCount = 0
        If dictIATypes.Exists(CStr(iKey)) Then
            aTypeCount = dictIATypes(CStr(iKey)).Count
        End If
        If aTypeCount >= 2 Then
            aBonus = 1
        Else
            aBonus = 0
        End If

        ' --- yCount: Y列が「1」の行数 ---
        yCount = 0
        If dictIYCount.Exists(CStr(iKey)) Then
            yCount = dictIYCount(CStr(iKey))
        End If

        ' --- 最終判定 ---
        total = vCountSum + aBonus + yCount

        If total >= THRESHOLD Then
            targetICount = targetICount + 1
            ' このI値の全行のJ列を着色（既存の背景色があれば変更しない）
            Dim rowCol As Collection
            Set rowCol = dictIRows(CStr(iKey))
            Dim idx As Long
            For idx = 1 To rowCol.Count
                Dim rn As Long
                rn = rowCol(idx)
                If ws.Cells(rn, "J").Interior.Pattern = xlNone Then
                    ws.Cells(rn, "J").Interior.Color = lightBlue
                End If
            Next idx
        End If
    Next iKey

    ' --- 完了 ---
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "枝番チェック完了。" & vbCrLf & _
           "対象Iグループ数: " & targetICount, vbInformation, "枝番チェック"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation, "枝番チェック"
End Sub

' ============================================================
' Y列の「1」判定（Trim＋空白除去して "1" かチェック）
' ============================================================
Private Function IsYValueOne(ByVal s As String) As Boolean
    s = Trim(s)
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(&H3000), "")
    IsYValueOne = (s = "1")
End Function

' ============================================================
' 複数列から最終行の最大値を取得
' ============================================================
Private Function GetMaxLastRow(ByVal ws As Worksheet, ByVal cols As Variant) As Long
    Dim mx As Long
    Dim c As Variant
    Dim lr As Long
    mx = 1
    For Each c In cols
        lr = ws.Cells(ws.Rows.Count, CStr(c)).End(xlUp).Row
        If lr > mx Then mx = lr
    Next c
    GetMaxLastRow = mx
End Function

' ============================================================
' Variant のセーフ変換（エラー値・Null・Empty 対応）
' ============================================================
Private Function SafeVariant(ByVal v As Variant) As String
    If IsError(v) Then
        SafeVariant = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        SafeVariant = ""
    Else
        SafeVariant = CStr(v)
    End If
End Function
