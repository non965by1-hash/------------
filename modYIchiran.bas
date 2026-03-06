Option Explicit

'=============================================================================
' modYIchiran - Y一覧出力マクロ
'
' 「入力シート」のI列値をA列文頭の列番番号でグループ化し、
' 2種類以上の列番番号にまたがるものだけを「Y一覧」シートに出力する
'=============================================================================

' I列文字ごとの情報を保持する型
Private Type tItemInfo
    IValue          As String       ' I列の文字列
    ColNums()       As Long         ' 出現する列番番号の配列（ソート済み）
    ColNumCount     As Long         ' 出現する列番番号の種類数
    IsEarth         As Boolean      ' EARTH判定
    IsTwist         As Boolean      ' TWIST判定
    MaxM            As Double       ' M列の最大値（通常データ用）
    SortCategory    As Long         ' 0=通常, 1=TWIST, 2=EARTH
    HeadPriority    As Long         ' I列文頭による優先順位
End Type

Public Sub CreateYIchiran()

    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long, k As Long
    Dim cellA As String, cellI As String, cellM As String, cellQ As String
    Dim colNum As Long

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
    Dim ws As Worksheet
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

    ' --- 最終行の取得（A列とI列の大きい方） ---
    Dim lastRowA As Long, lastRowI As Long
    lastRowA = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row
    lastRowI = wsInput.Cells(wsInput.Rows.Count, "I").End(xlUp).Row
    lastRow = lastRowA
    If lastRowI > lastRow Then lastRow = lastRowI

    If lastRow < 2 Then
        MsgBox "データがありません。", vbInformation
        GoTo CleanUp
    End If

    ' --- 入力チェック: A列の列番番号形式チェック ---
    ' A列またはI列に値がある行を対象にチェック
    For i = 2 To lastRow
        cellA = Trim(CStr(wsInput.Cells(i, "A").Value))
        cellI = Trim(CStr(wsInput.Cells(i, "I").Value))

        ' A列またはI列に値がある行をチェック
        If cellA <> "" Or cellI <> "" Then
            If Not IsValidColNum(cellA) Then
                MsgBox "盤記号の前に列番番号を記入してください。", vbExclamation
                GoTo CleanUp
            End If
        End If
    Next i

    ' --- データ収集 ---
    ' dictColNums: key=I列文字, value=Dictionary(key=列番番号, value=True)
    Dim dictColNums As Object
    Set dictColNums = CreateObject("Scripting.Dictionary")

    ' dictQ: key=I列文字, value=Dictionary(key=Q値(大文字), value=True)
    Dim dictQ As Object
    Set dictQ = CreateObject("Scripting.Dictionary")

    ' dictM: key=I列文字, value=M列最大値
    Dim dictM As Object
    Set dictM = CreateObject("Scripting.Dictionary")

    ' 全列番番号を収集 + 列番番号ごとの代表A列値（ヘッダー用）
    Dim dictAllColNums As Object
    Set dictAllColNums = CreateObject("Scripting.Dictionary")

    ' dictColHeader: key=列番番号(Long), value=A列の値（最初に見つかったもの）
    Dim dictColHeader As Object
    Set dictColHeader = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        cellA = Trim(CStr(wsInput.Cells(i, "A").Value))
        cellI = Trim(CStr(wsInput.Cells(i, "I").Value))

        ' A列から列番番号を取得
        colNum = ExtractColNum(cellA)
        If colNum = 0 Then GoTo NextRow  ' 列番番号が取れない行はスキップ

        ' 全列番番号を記録 + ヘッダー用のA列値を保持
        If Not dictAllColNums.Exists(colNum) Then
            dictAllColNums.Add colNum, True
            dictColHeader.Add colNum, cellA  ' 最初に見つかったA列値をヘッダーにする
        End If

        ' I列が空の行はここでスキップ（列番番号の収集だけは上で済ませる）
        If cellI = "" Then GoTo NextRow

        ' I列文字ごとに列番番号を記録
        If Not dictColNums.Exists(cellI) Then
            Dim innerDict As Object
            Set innerDict = CreateObject("Scripting.Dictionary")
            Set dictColNums(cellI) = innerDict
        End If
        If Not dictColNums(cellI).Exists(colNum) Then
            dictColNums(cellI).Add colNum, True
        End If

        ' Q列の収集
        cellQ = UCase(Trim(CStr(wsInput.Cells(i, "Q").Value)))
        If Not dictQ.Exists(cellI) Then
            Dim qDict As Object
            Set qDict = CreateObject("Scripting.Dictionary")
            Set dictQ(cellI) = qDict
        End If
        If cellQ <> "" Then
            If Not dictQ(cellI).Exists(cellQ) Then
                dictQ(cellI).Add cellQ, True
            End If
        End If

        ' M列の最大値収集
        cellM = Trim(CStr(wsInput.Cells(i, "M").Value))
        Dim mVal As Double
        If IsNumeric(cellM) And cellM <> "" Then
            mVal = CDbl(cellM)
        Else
            mVal = 0
        End If
        If Not dictM.Exists(cellI) Then
            dictM.Add cellI, mVal
        Else
            If mVal > dictM(cellI) Then
                dictM(cellI) = mVal
            End If
        End If

NextRow:
    Next i

    ' --- 2種類以上の列番番号にまたがるI値を抽出 ---
    Dim arrItems() As tItemInfo
    Dim itemCount As Long
    itemCount = 0

    Dim keys As Variant
    keys = dictColNums.keys

    For i = 0 To UBound(keys)
        Dim iVal As String
        iVal = keys(i)

        If dictColNums(iVal).Count >= 2 Then
            itemCount = itemCount + 1
            ReDim Preserve arrItems(1 To itemCount)

            With arrItems(itemCount)
                .IValue = iVal
                .ColNumCount = dictColNums(iVal).Count

                ' 列番番号配列を作成
                ReDim .ColNums(1 To .ColNumCount)
                Dim cnKeys As Variant
                cnKeys = dictColNums(iVal).keys
                For j = 0 To UBound(cnKeys)
                    .ColNums(j + 1) = cnKeys(j)
                Next j
                ' 列番番号をソート（昇順）
                SortLongArray .ColNums, 1, .ColNumCount

                ' EARTH/TWIST判定
                .IsEarth = False
                .IsTwist = False
                If dictQ.Exists(iVal) Then
                    If dictQ(iVal).Exists("EARTH") Then
                        .IsEarth = True
                    End If
                    If Not .IsEarth Then
                        If dictQ(iVal).Exists("TWIST") Then
                            .IsTwist = True
                        End If
                    End If
                End If

                ' ソートカテゴリ: 0=通常, 1=TWIST, 2=EARTH
                If .IsEarth Then
                    .SortCategory = 2
                ElseIf .IsTwist Then
                    .SortCategory = 1
                Else
                    .SortCategory = 0
                End If

                ' M列最大値
                If dictM.Exists(iVal) Then
                    .MaxM = dictM(iVal)
                Else
                    .MaxM = 0
                End If

                ' 文頭優先順位
                .HeadPriority = GetHeadPriority(iVal)
            End With
        End If
    Next i

    If itemCount = 0 Then
        MsgBox "2種類以上の列番番号にまたがるI列データがありません。", vbInformation
        GoTo CleanUp
    End If

    ' --- ソート（バブルソート） ---
    ' ソート順:
    '   1. SortCategory昇順（0=通常 → 1=TWIST → 2=EARTH）
    '   2. 通常: M列降順 → 文頭優先順位昇順 → 文字列昇順
    '   3. TWIST/EARTH: 文頭優先順位昇順 → 文字列昇順
    Dim swapped As Boolean
    Dim tempItem As tItemInfo
    Do
        swapped = False
        For i = 1 To itemCount - 1
            If ShouldSwap(arrItems(i), arrItems(i + 1)) Then
                tempItem = arrItems(i)
                arrItems(i) = arrItems(i + 1)
                arrItems(i + 1) = tempItem
                swapped = True
            End If
        Next i
    Loop While swapped

    ' --- Y一覧シートの準備 ---
    ' 既存のY一覧シートがあれば削除
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Y一覧" Then
            ws.Delete
            Exit For
        End If
    Next
    Application.DisplayAlerts = True

    ' 新規作成
    Set wsOutput = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOutput.Name = "Y一覧"

    ' 全セルの表示形式を文字列に
    wsOutput.Cells.NumberFormatLocal = "@"

    ' --- 全列番番号リストをソート（数値昇順） ---
    Dim allColNums() As Long
    Dim allColNumCount As Long
    Dim acnKeys As Variant
    acnKeys = dictAllColNums.keys
    allColNumCount = dictAllColNums.Count
    ReDim allColNums(1 To allColNumCount)
    For i = 0 To UBound(acnKeys)
        allColNums(i + 1) = acnKeys(i)
    Next i
    SortLongArray allColNums, 1, allColNumCount

    ' 列番番号→出力列位置のマッピング（奇数列: 1, 3, 5, 7, ...）
    Dim dictColToPos As Object
    Set dictColToPos = CreateObject("Scripting.Dictionary")
    For i = 1 To allColNumCount
        dictColToPos.Add allColNums(i), (i - 1) * 2 + 1
    Next i

    ' --- 1行目: ヘッダー（列番番号＋盤記号） ---
    ' サンプルに合わせて、A列の値全体をヘッダーにする
    Dim outCol As Long
    For i = 1 To allColNumCount
        outCol = dictColToPos(allColNums(i))
        If dictColHeader.Exists(allColNums(i)) Then
            wsOutput.Cells(1, outCol).Value = dictColHeader(allColNums(i))
        Else
            wsOutput.Cells(1, outCol).Value = Format(allColNums(i), "00") & "_"
        End If
    Next i

    ' --- 2行目: Y連番 ---
    For i = 1 To allColNumCount
        outCol = dictColToPos(allColNums(i))
        wsOutput.Cells(2, outCol).Value = "Y" & CStr(i)
    Next i

    ' --- 3行目以降: データ出力 ---
    ' サンプルに合わせたレイアウト:
    '   - 全ての奇数列に行番号（行単位で同じ番号）を出力
    '   - 偶数列にはI値、または「-」、または空欄
    For i = 1 To itemCount
        Dim outRow As Long
        outRow = i + 2  ' 3行目から

        With arrItems(i)
            ' この I値が出現する列番番号の最小・最大を取得
            Dim minCN As Long, maxCN As Long
            minCN = .ColNums(1)
            maxCN = .ColNums(.ColNumCount)

            ' 全列番番号について処理
            For j = 1 To allColNumCount
                Dim cn As Long
                cn = allColNums(j)
                outCol = dictColToPos(cn)

                ' 奇数列: 全列に行通し番号を出力（サンプル準拠）
                wsOutput.Cells(outRow, outCol).Value = CStr(i)

                ' 偶数列: I値 / - / 空欄
                If cn >= minCN And cn <= maxCN Then
                    ' 出現範囲内
                    If ExistsInArray(.ColNums, .ColNumCount, cn) Then
                        ' この列番番号に存在する → I値を書き込み
                        wsOutput.Cells(outRow, outCol + 1).Value = .IValue
                    Else
                        ' 範囲内だが存在しない → 「-」
                        wsOutput.Cells(outRow, outCol + 1).Value = "-"
                    End If
                End If
                ' 範囲外（左端より左、右端より右）は空欄のまま
            Next j
        End With
    Next i

    ' --- 列幅自動調整 ---
    wsOutput.Cells.EntireColumn.AutoFit

    ' --- 完了メッセージ ---
    wsOutput.Activate
    wsOutput.Range("A1").Select
    MsgBox "Y一覧の出力が完了しました。" & vbCrLf & _
           "抽出件数: " & itemCount & " 件", vbInformation

CleanUp:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

'=============================================================================
' A列の文頭が「2桁数字_」形式かどうかを判定
'=============================================================================
Private Function IsValidColNum(ByVal s As String) As Boolean
    If Len(s) < 3 Then
        IsValidColNum = False
        Exit Function
    End If
    ' 先頭2文字が数字、3文字目が_
    If Mid(s, 1, 1) >= "0" And Mid(s, 1, 1) <= "9" And _
       Mid(s, 2, 1) >= "0" And Mid(s, 2, 1) <= "9" And _
       Mid(s, 3, 1) = "_" Then
        IsValidColNum = True
    Else
        IsValidColNum = False
    End If
End Function

'=============================================================================
' A列文字列から列番番号（数値）を抽出
' 例: "01_ABC" → 1, "05_xxx" → 5
'=============================================================================
Private Function ExtractColNum(ByVal s As String) As Long
    If IsValidColNum(s) Then
        ExtractColNum = CLng(Left(s, 2))
    Else
        ExtractColNum = 0
    End If
End Function

'=============================================================================
' I列文頭による優先順位を返す
' U*, V*, W*, A*, B*, C*, F*, R*, S*, T*, O*, P*, M*, N* の順
' 該当しない場合は最後尾扱い（99）
'=============================================================================
Private Function GetHeadPriority(ByVal s As String) As Long
    If Len(s) = 0 Then
        GetHeadPriority = 99
        Exit Function
    End If

    Dim firstChar As String
    firstChar = UCase(Left(s, 1))

    Select Case firstChar
        Case "U": GetHeadPriority = 1
        Case "V": GetHeadPriority = 2
        Case "W": GetHeadPriority = 3
        Case "A": GetHeadPriority = 4
        Case "B": GetHeadPriority = 5
        Case "C": GetHeadPriority = 6
        Case "F": GetHeadPriority = 7
        Case "R": GetHeadPriority = 8
        Case "S": GetHeadPriority = 9
        Case "T": GetHeadPriority = 10
        Case "O": GetHeadPriority = 11
        Case "P": GetHeadPriority = 12
        Case "M": GetHeadPriority = 13
        Case "N": GetHeadPriority = 14
        Case Else: GetHeadPriority = 99
    End Select
End Function

'=============================================================================
' ソート比較: item1がitem2の後ろに来るべきならTrue
'=============================================================================
Private Function ShouldSwap(ByRef item1 As tItemInfo, ByRef item2 As tItemInfo) As Boolean
    ShouldSwap = False

    ' 1. SortCategory昇順（通常=0 → TWIST=1 → EARTH=2）
    If item1.SortCategory <> item2.SortCategory Then
        ShouldSwap = (item1.SortCategory > item2.SortCategory)
        Exit Function
    End If

    ' 同じカテゴリ内のソート
    If item1.SortCategory = 0 Then
        ' 通常データ: M列降順（大きい方が上）
        If item1.MaxM <> item2.MaxM Then
            ShouldSwap = (item1.MaxM < item2.MaxM)
            Exit Function
        End If
    End If

    ' 文頭優先順位昇順（小さい方が上）
    If item1.HeadPriority <> item2.HeadPriority Then
        ShouldSwap = (item1.HeadPriority > item2.HeadPriority)
        Exit Function
    End If

    ' 文字列昇順
    If item1.IValue <> item2.IValue Then
        ShouldSwap = (item1.IValue > item2.IValue)
        Exit Function
    End If
End Function

'=============================================================================
' Long型配列のバブルソート（昇順）
'=============================================================================
Private Sub SortLongArray(ByRef arr() As Long, ByVal lb As Long, ByVal ub As Long)
    Dim i As Long, j As Long
    Dim tmp As Long
    For i = lb To ub - 1
        For j = lb To ub - (i - lb) - 1
            If arr(j) > arr(j + 1) Then
                tmp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = tmp
            End If
        Next j
    Next i
End Sub

'=============================================================================
' Long型配列に値が存在するかチェック
'=============================================================================
Private Function ExistsInArray(ByRef arr() As Long, ByVal cnt As Long, ByVal val As Long) As Boolean
    Dim i As Long
    For i = 1 To cnt
        If arr(i) = val Then
            ExistsInArray = True
            Exit Function
        End If
    Next i
    ExistsInArray = False
End Function
