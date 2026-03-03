Attribute VB_Name = "modPartsIndex"
Option Explicit

' ============================================================
' modPartsIndex - 部品コード抽出マクロ（標準モジュール）
' ============================================================
' parts database.xlsm の全シートC列をインデックス化し、
' G列入力時に候補を検索・表示・転記する。
' ============================================================

' --- 定数 ---
' ネットワークパスを設定する場合はここを変更
' 例: "\\server\share\folder\parts database.xlsm"
' 空文字の場合、ブックAと同じフォルダから検索する（フォールバック）
Public Const PARTS_DB_FULLPATH As String = ""
Public Const PARTS_DB_FILENAME As String = "parts database.xlsm"
Public Const MAX_CANDIDATES As Long = 50
Public Const TARGET_SHEET_NAME As String = "入力シート"

' --- インデックスレコード ---
Public Type PartsRecord
    SheetName As String
    Row As Long
    ColA As String       ' DB A列
    ColB As String       ' DB B列
    ColC As String       ' DB C列（検索対象）
    ColD As String       ' DB D列
    ColE As String       ' DB E列（J列赤制御用）
    NormalizedC As String ' 検索用正規化済みC列
End Type

' --- 候補レコード（スコア付き） ---
Public Type CandidateRecord
    Rec As PartsRecord
    Score As Long        ' 1000=完全一致, 800=前方一致, 650=境界一致, 500=部分一致
    MatchPos As Long     ' InStr位置（タイブレーク用）
    LenDiff As Long      ' |Len(C)-Len(key)|（タイブレーク用）
End Type

' --- グローバル変数 ---
Public g_Index() As PartsRecord
Public g_IndexCount As Long
Public g_IndexBuilt As Boolean
Public g_bIgnoreChange As Boolean       ' 無限ループ防止フラグ

' --- フォーム連携用グローバル ---
Public g_InitialKey As String           ' フォームに渡す初期検索文字
Public g_SelectedIndex As Long          ' -1=キャンセル, >0=選択された候補番号
Public g_FilteredCandidates() As CandidateRecord
Public g_FilteredCount As Long
Public g_TotalMatchCount As Long        ' 50件制限前の総マッチ数

' ============================================================
' セーフ文字列変換（エラー値・Null・Empty 対応）
' ============================================================
Private Function SafeCStr(ByVal v As Variant) As String
    If IsError(v) Then
        SafeCStr = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        SafeCStr = ""
    Else
        SafeCStr = CStr(v)
    End If
End Function

' ============================================================
' 検索用正規化（Trim + LCase）
' ============================================================
Public Function NormalizeForSearch(ByVal s As String) As String
    s = Trim(s)
    s = LCase(s)
    NormalizeForSearch = s
End Function

' ============================================================
' E列「1」判定（全空白除去して "1" かチェック）
' 文字列化 → Trim → 半角/全角スペース全除去 → "1" なら True
' ============================================================
Public Function IsValueOne(ByVal v As Variant) As Boolean
    Dim s As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        IsValueOne = False
        Exit Function
    End If
    s = CStr(v)
    s = Trim(s)
    ' 半角スペース除去
    s = Replace(s, " ", "")
    ' 全角スペース除去
    s = Replace(s, ChrW(&H3000), "")
    IsValueOne = (s = "1")
End Function

' ============================================================
' DBファイルパス取得
' ============================================================
Private Function GetPartsDBPath() As String
    Dim fPath As String

    ' フルパス指定がある場合
    If Len(PARTS_DB_FULLPATH) > 0 Then
        If Dir(PARTS_DB_FULLPATH) <> "" Then
            GetPartsDBPath = PARTS_DB_FULLPATH
            Exit Function
        End If
    End If

    ' フォールバック：ブックAと同じフォルダ
    fPath = ThisWorkbook.Path & Application.PathSeparator & PARTS_DB_FILENAME
    If Dir(fPath) <> "" Then
        GetPartsDBPath = fPath
        Exit Function
    End If

    GetPartsDBPath = ""
End Function

' ============================================================
' DBブックを開く（または既に開いているものを返す）
' ============================================================
Private Function OpenPartsDB(ByVal dbPath As String, ByRef wasOpen As Boolean) As Workbook
    Dim wb As Workbook
    Dim fName As String

    fName = Dir(dbPath)

    ' 既に開いているか確認
    On Error Resume Next
    For Each wb In Application.Workbooks
        If LCase(wb.Name) = LCase(fName) Then
            Set OpenPartsDB = wb
            wasOpen = True
            Exit Function
        End If
    Next wb
    On Error GoTo 0

    ' ReadOnlyで開く
    On Error Resume Next
    Set wb = Workbooks.Open(Filename:=dbPath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo 0

    If Not wb Is Nothing Then
        Set OpenPartsDB = wb
        wasOpen = False
    End If
End Function

' ============================================================
' インデックス構築（Workbook_Open または初回呼び出し時）
' parts database の全シート C列(2行目以降)を走査し、
' A/B/C/D/E列の値をメモリに保持する。
' ============================================================
Public Sub BuildPartsIndex()
    Dim wbDB As Workbook
    Dim ws As Worksheet
    Dim dbPath As String
    Dim lastRow As Long
    Dim r As Long
    Dim dataArr As Variant
    Dim wasOpen As Boolean
    Dim cValue As String

    ' パス決定
    dbPath = GetPartsDBPath()
    If dbPath = "" Then
        MsgBox "parts database.xlsm が見つかりません。" & vbCrLf & _
               "ブックと同じフォルダに配置するか、" & vbCrLf & _
               "PARTS_DB_FULLPATH 定数を設定してください。", _
               vbExclamation, "エラー"
        g_IndexBuilt = False
        Exit Sub
    End If

    ' ブックを開く
    Set wbDB = OpenPartsDB(dbPath, wasOpen)
    If wbDB Is Nothing Then
        MsgBox "parts database.xlsm を開けませんでした。" & vbCrLf & _
               "パス: " & dbPath, vbExclamation, "エラー"
        g_IndexBuilt = False
        Exit Sub
    End If

    ' インデックス初期化（倍増方式で拡張）
    g_IndexCount = 0
    ReDim g_Index(1 To 2000)

    ' 全シートを走査してインデックス構築
    For Each ws In wbDB.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
        If lastRow >= 2 Then
            ' A~E列を一括読み込み（高速化）
            dataArr = ws.Range("A2:E" & lastRow).Value

            For r = 1 To UBound(dataArr, 1)
                cValue = SafeCStr(dataArr(r, 3))
                If Len(Trim(cValue)) > 0 Then
                    g_IndexCount = g_IndexCount + 1

                    ' 配列拡張（倍増方式）
                    If g_IndexCount > UBound(g_Index) Then
                        ReDim Preserve g_Index(1 To UBound(g_Index) * 2)
                    End If

                    With g_Index(g_IndexCount)
                        .SheetName = ws.Name
                        .Row = r + 1
                        .ColA = SafeCStr(dataArr(r, 1))
                        .ColB = SafeCStr(dataArr(r, 2))
                        .ColC = cValue
                        .ColD = SafeCStr(dataArr(r, 4))
                        .ColE = SafeCStr(dataArr(r, 5))
                        .NormalizedC = NormalizeForSearch(cValue)
                    End With
                End If
            Next r
        End If
    Next ws

    ' 正確なサイズにリサイズ
    If g_IndexCount > 0 Then
        ReDim Preserve g_Index(1 To g_IndexCount)
    Else
        Erase g_Index
    End If

    ' 元々開いていなかったら閉じる
    If Not wasOpen Then
        wbDB.Close SaveChanges:=False
    End If

    g_IndexBuilt = True
End Sub

' ============================================================
' 候補検索（スコア付き・ソート済み）
' インデックスに対して検索するため、DB走査は発生しない。
'
' スコアリング仕様:
'   1000: 完全一致
'    800: 前方一致（Cがkeyで始まる）
'    650: 単語境界一致（スペース/記号区切り直後にkeyが出る）
'    500: 通常の部分一致
'
' タイブレーク:
'   1. マッチ位置が早い（InStr位置昇順）
'   2. 長さ差が小さい（|Len(C)-Len(key)| 昇順）
'   3. C列昇順
'
' 戻り値: 50件制限前の総マッチ数
' ============================================================
Public Function SearchCandidates(ByVal key As String) As Long
    Dim normalizedKey As String
    Dim i As Long
    Dim totalCount As Long
    Dim Score As Long
    Dim MatchPos As Long
    Dim LenDiff As Long
    Dim normalizedC As String
    Dim prevChar As String

    normalizedKey = NormalizeForSearch(key)

    If Len(normalizedKey) = 0 Then
        g_FilteredCount = 0
        g_TotalMatchCount = 0
        SearchCandidates = 0
        Exit Function
    End If

    ' インデックスが未構築なら構築
    If Not g_IndexBuilt Then BuildPartsIndex
    If g_IndexCount = 0 Then
        g_FilteredCount = 0
        g_TotalMatchCount = 0
        SearchCandidates = 0
        Exit Function
    End If

    ' 一時候補配列
    Dim tempCandidates() As CandidateRecord
    ReDim tempCandidates(1 To g_IndexCount)
    totalCount = 0

    For i = 1 To g_IndexCount
        normalizedC = g_Index(i).NormalizedC
        MatchPos = InStr(1, normalizedC, normalizedKey, vbTextCompare)

        If MatchPos > 0 Then
            totalCount = totalCount + 1

            ' スコア計算
            If normalizedC = normalizedKey Then
                ' 完全一致（グループ1: 最優先）
                Score = 1000
            ElseIf MatchPos = 1 Then
                ' 前方一致
                Score = 800
            Else
                ' 単語境界チェック（マッチ位置の直前の文字）
                prevChar = Mid(normalizedC, MatchPos - 1, 1)
                If prevChar = " " Or prevChar = "-" Or prevChar = "_" Or _
                   prevChar = "/" Or prevChar = "." Or prevChar = "(" Or _
                   prevChar = "," Or prevChar = ";" Or _
                   prevChar = ChrW(&H3000) Then
                    ' 単語境界一致
                    Score = 650
                Else
                    ' 通常の部分一致
                    Score = 500
                End If
            End If

            LenDiff = Abs(CLng(Len(normalizedC)) - CLng(Len(normalizedKey)))

            With tempCandidates(totalCount)
                .Rec = g_Index(i)
                .Score = Score
                .MatchPos = MatchPos
                .LenDiff = LenDiff
            End With
        End If
    Next i

    g_TotalMatchCount = totalCount
    SearchCandidates = totalCount

    If totalCount = 0 Then
        g_FilteredCount = 0
        Exit Function
    End If

    ' ソート（スコア降順 → マッチ位置昇順 → 長さ差昇順 → C列昇順）
    ReDim Preserve tempCandidates(1 To totalCount)
    SortCandidates tempCandidates, 1, totalCount

    ' 最大50件に制限（完全一致を優先して詰め、残り枠に部分一致）
    g_FilteredCount = totalCount
    If g_FilteredCount > MAX_CANDIDATES Then
        g_FilteredCount = MAX_CANDIDATES
    End If

    ReDim g_FilteredCandidates(1 To g_FilteredCount)
    For i = 1 To g_FilteredCount
        g_FilteredCandidates(i) = tempCandidates(i)
    Next i
End Function

' ============================================================
' クイックソート（候補配列をスコア順にソート）
' ============================================================
Private Sub SortCandidates(ByRef arr() As CandidateRecord, ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub

    Dim i As Long, j As Long
    Dim pivot As CandidateRecord
    Dim temp As CandidateRecord

    i = lo
    j = hi
    pivot = arr((lo + hi) \ 2)

    Do While i <= j
        Do While CompareCandidates(arr(i), pivot) < 0
            i = i + 1
        Loop
        Do While CompareCandidates(arr(j), pivot) > 0
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If lo < j Then SortCandidates arr, lo, j
    If i < hi Then SortCandidates arr, i, hi
End Sub

' ============================================================
' 候補比較（ソート順序定義）
' 戻り値: 負=aが先, 正=bが先, 0=同等
' ============================================================
Private Function CompareCandidates(ByRef a As CandidateRecord, ByRef b As CandidateRecord) As Long
    ' スコア降順（高いスコアが先＝完全一致が上）
    If a.Score > b.Score Then
        CompareCandidates = -1: Exit Function
    ElseIf a.Score < b.Score Then
        CompareCandidates = 1: Exit Function
    End If

    ' マッチ位置昇順（小さい位置が先）
    If a.MatchPos < b.MatchPos Then
        CompareCandidates = -1: Exit Function
    ElseIf a.MatchPos > b.MatchPos Then
        CompareCandidates = 1: Exit Function
    End If

    ' 長さ差昇順（近い長さが先）
    If a.LenDiff < b.LenDiff Then
        CompareCandidates = -1: Exit Function
    ElseIf a.LenDiff > b.LenDiff Then
        CompareCandidates = 1: Exit Function
    End If

    ' C列昇順（最終タイブレーク）
    CompareCandidates = StrComp(a.Rec.ColC, b.Rec.ColC, vbTextCompare)
End Function

' ============================================================
' J列背景色制御
' DB E列=1 → 赤（RGB(255,0,0)）
' DB E列≠1 → 塗りつぶし無し（xlNone）
' ============================================================
Private Sub ApplyJColor(ByVal ws As Worksheet, ByVal Row As Long, ByVal eValue As String)
    If IsValueOne(eValue) Then
        ws.Cells(Row, "J").Interior.Color = RGB(255, 0, 0)
    Else
        ws.Cells(Row, "J").Interior.Pattern = xlNone
    End If
End Sub

' ============================================================
' 転記処理（E/F/G列 + J列背景色 + C列一致行への一括反映）
'
' 転記ルール:
'   DB A列 → ブックA E列
'   DB B列 → ブックA F列
'   DB C列 → ブックA G列
'
' さらに、ブックA C列に同じ値を持つ全行にも同様に転記し、
' J列の赤背景ON/OFFも同時に適用する。
' ============================================================
Public Sub TransferData(ByVal targetRow As Long, ByVal candidateIdx As Long)
    Dim ws As Worksheet
    Dim rec As PartsRecord
    Dim lastRow As Long
    Dim r As Long
    Dim cVal As String
    Dim cRange As Variant

    Set ws = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)
    rec = g_FilteredCandidates(candidateIdx).Rec

    ' --- 対象行に転記 ---
    ws.Cells(targetRow, "E").Value = rec.ColA   ' DB A列 → E列
    ws.Cells(targetRow, "F").Value = rec.ColB   ' DB B列 → F列
    ws.Cells(targetRow, "G").Value = rec.ColC   ' DB C列 → G列
    ApplyJColor ws, targetRow, rec.ColE         ' J列背景制御

    ' --- C列一致行への一括反映 ---
    cVal = rec.ColC
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    If lastRow >= 2 Then
        ' C列を配列で一括読み込み（高速化）
        cRange = ws.Range("C2:C" & lastRow).Value

        ' 単一セルの場合の対応
        If Not IsArray(cRange) Then
            If CStr(cRange) = cVal And 2 <> targetRow Then
                ws.Cells(2, "E").Value = rec.ColA
                ws.Cells(2, "F").Value = rec.ColB
                ws.Cells(2, "G").Value = rec.ColC
                ApplyJColor ws, 2, rec.ColE
            End If
        Else
            For r = 1 To UBound(cRange, 1)
                If (r + 1) <> targetRow Then
                    If SafeCStr(cRange(r, 1)) = cVal Then
                        ws.Cells(r + 1, "E").Value = rec.ColA
                        ws.Cells(r + 1, "F").Value = rec.ColB
                        ws.Cells(r + 1, "G").Value = rec.ColC
                        ApplyJColor ws, r + 1, rec.ColE
                    End If
                End If
            Next r
        End If
    End If
End Sub

' ============================================================
' メイン処理（Worksheet_Change から呼ばれる）
' ============================================================
Public Sub ProcessGColumnInput(ByVal targetRow As Long, ByVal key As String)
    Dim totalCount As Long
    Dim calcMode As XlCalculation

    On Error GoTo ErrorHandler

    ' パフォーマンス設定を保存・制御
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' インデックス確認（未構築なら構築）
    If Not g_IndexBuilt Then BuildPartsIndex
    If Not g_IndexBuilt Then GoTo CleanUp

    ' 検索実行
    totalCount = SearchCandidates(key)

    ' フォーム用の初期キーを設定
    g_InitialKey = key
    g_SelectedIndex = -1

    ' フォーム表示のため画面更新を復帰
    Application.ScreenUpdating = True
    Application.Calculation = calcMode

    ' フォーム表示（モーダル）- 候補0件でも表示する
    frmPartsPick.Show vbModal

    ' 結果処理
    If g_SelectedIndex > 0 And g_SelectedIndex <= g_FilteredCount Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

        ' ガードフラグON（G列書き戻しの無限ループ防止）
        g_bIgnoreChange = True

        ' 転記
        TransferData targetRow, g_SelectedIndex

        g_bIgnoreChange = False
    End If
    ' Escキャンセル時（g_SelectedIndex = -1）は何もしない

    GoTo CleanUp

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation, "部品コード抽出エラー"

CleanUp:
    ' 必ず復帰（エラー時も含む）
    g_bIgnoreChange = False
    Application.ScreenUpdating = True
    On Error Resume Next
    Application.Calculation = calcMode
    On Error GoTo 0
End Sub

' ============================================================
' インデックス再構築（手動実行用マクロ）
' Alt+F8 → RebuildIndex で実行可能
' ============================================================
Public Sub RebuildIndex()
    g_IndexBuilt = False
    BuildPartsIndex
    If g_IndexBuilt Then
        MsgBox "インデックスを再構築しました。" & vbCrLf & _
               "レコード数: " & g_IndexCount, vbInformation, "完了"
    End If
End Sub
