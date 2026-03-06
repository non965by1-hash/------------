Option Explicit

'=============================================================================
' modBunkatsuSenban - 分割線番一覧マクロ（標準モジュール）
'
' 「入力シート」のI列グループごとに、ユーザーが選択した面名（V列）と
' 別の面名が共存するグループを抽出し、「分割線番」シートに出力する。
'
' ■配置先：標準モジュール（例: Module1）
'=============================================================================

'---------------------------------------------------------------------
' メインマクロ：マクロウインドウから「分割線番一覧」で実行する
'---------------------------------------------------------------------
Public Sub 分割線番一覧()
    ' ユーザーフォームを表示する
    Dim frm As frmBunkatsu
    Set frm = New frmBunkatsu
    frm.Show
    Set frm = Nothing
End Sub

'---------------------------------------------------------------------
' フォームからOK押下時に呼ばれる処理本体
' selectedMenName : ユーザーが選択した面名
'---------------------------------------------------------------------
Public Sub ExecuteBunkatsu(ByVal selectedMenName As String)

    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ws As Worksheet

    ' --- パフォーマンス制御 ---
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim calcMode As XlCalculation
    calcMode = Application.Calculation
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrHandler

    ' --- 入力シートの取得 ---
    Dim wsFound As Boolean
    wsFound = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "入力シート" Then
            wsFound = True
            Set wsInput = ws
            Exit For
        End If
    Next
    If Not wsFound Then
        MsgBox "「入力シート」が見つかりません。", vbExclamation
        GoTo CleanUp
    End If

    ' --- 最終行の取得（I列基準、V列も考慮） ---
    Dim lastRowI As Long, lastRowV As Long
    lastRowI = wsInput.Cells(wsInput.Rows.Count, "I").End(xlUp).Row
    lastRowV = wsInput.Cells(wsInput.Rows.Count, "V").End(xlUp).Row
    lastRow = lastRowI
    If lastRowV > lastRow Then lastRow = lastRowV
    If lastRow < 2 Then
        MsgBox "データがありません。", vbInformation
        GoTo CleanUp
    End If

    ' --- データ収集 ---
    ' Iグループごとに以下を収集する:
    '   - 選択面名の行が存在するか
    '   - 別面名の行が存在するか
    '   - 選択面名の行のM列・N列値（最初に見つかったもの）

    ' dictHasSelected: key=I値, value=True（選択面名の行が存在）
    Dim dictHasSelected As Object
    Set dictHasSelected = CreateObject("Scripting.Dictionary")

    ' dictHasOther: key=I値, value=True（別面名の行が存在）
    Dim dictHasOther As Object
    Set dictHasOther = CreateObject("Scripting.Dictionary")

    ' dictM: key=I値, value=選択面名行のM列値（最初に見つかったもの）
    Dim dictM As Object
    Set dictM = CreateObject("Scripting.Dictionary")

    ' dictN: key=I値, value=選択面名行のN列値（最初に見つかったもの）
    Dim dictN As Object
    Set dictN = CreateObject("Scripting.Dictionary")

    ' dictOrder: key=I値, value=初出行番号（出力順を安定させるため）
    Dim dictOrder As Object
    Set dictOrder = CreateObject("Scripting.Dictionary")

    Dim cellI As String, cellV As String

    For i = 2 To lastRow
        cellI = Trim(CStr(wsInput.Cells(i, "I").Value))
        cellV = Trim(CStr(wsInput.Cells(i, "V").Value))

        ' I列が空白の行はスキップ
        If cellI = "" Then GoTo NextRow

        ' 初出順を記録
        If Not dictOrder.Exists(cellI) Then
            dictOrder.Add cellI, i
        End If

        ' V列が空白の行は面名判定対象外
        If cellV = "" Then GoTo NextRow

        If cellV = selectedMenName Then
            ' 選択面名の行
            If Not dictHasSelected.Exists(cellI) Then
                dictHasSelected.Add cellI, True
                ' 最初に見つかった選択面名行のM列・N列を記録
                dictM.Add cellI, CStr(wsInput.Cells(i, "M").Value)
                dictN.Add cellI, CStr(wsInput.Cells(i, "N").Value)
            End If
        Else
            ' 別面名の行
            If Not dictHasOther.Exists(cellI) Then
                dictHasOther.Add cellI, True
            End If
        End If

NextRow:
    Next i

    ' --- 出力対象の抽出（選択面名あり AND 別面名あり） ---
    Dim arrOutput() As String  ' (n, 0)=I値, (n, 1)=M値, (n, 2)=N値
    Dim outCount As Long
    outCount = 0

    ' dictOrder の初出順にソートして出力順を安定させる
    Dim allIKeys As Variant
    If dictOrder.Count = 0 Then
        GoTo OutputPhase
    End If
    allIKeys = dictOrder.keys

    ' 初出行番号順にソート（シェルソート）
    Dim allIRows As Variant
    allIRows = dictOrder.Items
    Dim gap As Long, ii As Long, jj As Long
    Dim tmpKey As Variant, tmpRow As Variant
    gap = dictOrder.Count \ 2
    Do While gap > 0
        For ii = gap To UBound(allIKeys)
            tmpKey = allIKeys(ii)
            tmpRow = allIRows(ii)
            jj = ii
            Do While jj >= gap
                If allIRows(jj - gap) > tmpRow Then
                    allIKeys(jj) = allIKeys(jj - gap)
                    allIRows(jj) = allIRows(jj - gap)
                    jj = jj - gap
                Else
                    Exit Do
                End If
            Loop
            allIKeys(jj) = tmpKey
            allIRows(jj) = tmpRow
        Next ii
        gap = gap \ 2
    Loop

    ' 条件を満たすものを抽出
    For i = 0 To UBound(allIKeys)
        Dim iKey As String
        iKey = CStr(allIKeys(i))

        If dictHasSelected.Exists(iKey) And dictHasOther.Exists(iKey) Then
            outCount = outCount + 1
            ReDim Preserve arrOutput(1 To outCount, 1 To 3)

            ' ReDim Preserve は最後の次元しか変更できないので、
            ' 一旦別の方法で格納する
        End If
    Next i

    ' ReDim Preserve の制約があるため、二次元配列を使い直す
    outCount = 0
    Dim tmpI() As String, tmpM() As String, tmpN() As String
    ' まずカウント
    For i = 0 To UBound(allIKeys)
        iKey = CStr(allIKeys(i))
        If dictHasSelected.Exists(iKey) And dictHasOther.Exists(iKey) Then
            outCount = outCount + 1
        End If
    Next i

    If outCount > 0 Then
        ReDim tmpI(1 To outCount)
        ReDim tmpM(1 To outCount)
        ReDim tmpN(1 To outCount)
        Dim idx As Long
        idx = 0
        For i = 0 To UBound(allIKeys)
            iKey = CStr(allIKeys(i))
            If dictHasSelected.Exists(iKey) And dictHasOther.Exists(iKey) Then
                idx = idx + 1
                tmpI(idx) = iKey
                tmpM(idx) = dictM(iKey)
                tmpN(idx) = dictN(iKey)
            End If
        Next i
    End If

OutputPhase:
    ' --- 「分割線番」シートの準備 ---
    ' 既存シートがあれば削除
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "分割線番" Then
            ws.Delete
            Exit For
        End If
    Next
    Application.DisplayAlerts = True

    ' 新規作成
    Set wsOutput = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOutput.Name = "分割線番"

    ' 全セルの表示形式を文字列に
    wsOutput.Cells.NumberFormatLocal = "@"

    ' --- 1行目: タイトル（選択した面名） ---
    wsOutput.Cells(1, 1).Value = selectedMenName

    ' --- 2行目: 見出し行 ---
    wsOutput.Cells(2, 1).Value = "線番"
    wsOutput.Cells(2, 2).Value = "線サイズ"
    wsOutput.Cells(2, 3).Value = "線色"

    ' --- 3行目以降: データ出力 ---
    If outCount > 0 Then
        For i = 1 To outCount
            wsOutput.Cells(i + 2, 1).Value = tmpI(i)
            wsOutput.Cells(i + 2, 2).Value = tmpM(i)
            wsOutput.Cells(i + 2, 3).Value = tmpN(i)
        Next i
    End If

    ' --- オートフィルター設定（2行目の見出しに対して） ---
    wsOutput.Range("A2:C2").AutoFilter

    ' --- 列幅自動調整 ---
    wsOutput.Cells.EntireColumn.AutoFit

    ' --- 完了メッセージ ---
    wsOutput.Activate
    wsOutput.Range("A1").Select

    If outCount = 0 Then
        MsgBox "該当データはありませんでした。" & vbCrLf & _
               "面名: " & selectedMenName, vbInformation
    Else
        MsgBox "分割線番一覧の出力が完了しました。" & vbCrLf & _
               "面名: " & selectedMenName & vbCrLf & _
               "出力件数: " & outCount & " 件", vbInformation
    End If

CleanUp:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical
    Resume CleanUp
End Sub
