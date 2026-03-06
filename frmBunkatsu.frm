'=============================================================================
' frmBunkatsu - 分割線番一覧 ユーザーフォーム
'
' ■フォーム名：frmBunkatsu
' ■配置先：VBEでユーザーフォームを挿入し、Name を frmBunkatsu に設定
'
' ■フォーム上に配置するコントロール一覧：
'   1. ラベル      Name: lblTitle     Caption: 分割される面名を選んでください
'   2. コンボボックス Name: cboMen      Style: 2 - fmStyleDropDownList
'   3. OKボタン    Name: btnOK        Caption: OK        Default: True
'   4. キャンセルボタン Name: btnCancel   Caption: キャンセル  Cancel: True
'
' ■フォームのプロパティ：
'   Caption: 分割線番一覧
'   StartUpPosition: 1 - CenterOwner（画面中央表示）
'
' ■推奨サイズ：
'   Width: 300, Height: 180 程度
'
' ■コントロール配置イメージ：
'   ┌──────────────────────────┐
'   │ 分割される面名を選んでください       │  ← lblTitle
'   │                                    │
'   │ [コンボボックス▼               ]    │  ← cboMen
'   │                                    │
'   │      [ OK ]    [キャンセル]         │  ← btnOK, btnCancel
'   └──────────────────────────┘
'=============================================================================

Option Explicit

'---------------------------------------------------------------------
' フォーム初期化：コンボボックスに面名一覧をセットする
'---------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim wsInput As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellV As String

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
        Unload Me
        Exit Sub
    End If

    ' --- V列の最終行を取得 ---
    lastRow = wsInput.Cells(wsInput.Rows.Count, "V").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "V列にデータがありません。", vbInformation
        Unload Me
        Exit Sub
    End If

    ' --- V列2行目以降の値を重複なしで収集 ---
    Dim dictMen As Object
    Set dictMen = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        cellV = Trim(CStr(wsInput.Cells(i, "V").Value))
        ' 空白は除外
        If cellV <> "" Then
            If Not dictMen.Exists(cellV) Then
                dictMen.Add cellV, True
            End If
        End If
    Next i

    If dictMen.Count = 0 Then
        MsgBox "V列に面名データがありません。", vbInformation
        Unload Me
        Exit Sub
    End If

    ' --- 面名一覧を昇順ソートしてコンボボックスに追加 ---
    Dim menKeys As Variant
    menKeys = dictMen.keys

    ' シェルソート（昇順）
    Dim gap As Long, ii As Long, jj As Long
    Dim tmpVal As Variant
    gap = UBound(menKeys) \ 2
    Do While gap > 0
        For ii = gap To UBound(menKeys)
            tmpVal = menKeys(ii)
            jj = ii
            Do While jj >= gap
                If CStr(menKeys(jj - gap)) > CStr(tmpVal) Then
                    menKeys(jj) = menKeys(jj - gap)
                    jj = jj - gap
                Else
                    Exit Do
                End If
            Loop
            menKeys(jj) = tmpVal
        Next ii
        gap = gap \ 2
    Loop

    ' コンボボックスに追加
    For i = 0 To UBound(menKeys)
        cboMen.AddItem CStr(menKeys(i))
    Next i

    ' 最初の項目を選択状態にする（任意）
    If cboMen.ListCount > 0 Then
        cboMen.ListIndex = 0
    End If
End Sub

'---------------------------------------------------------------------
' OKボタン押下時の処理
'---------------------------------------------------------------------
Private Sub btnOK_Click()
    ' --- 選択チェック ---
    If cboMen.ListIndex < 0 Then
        MsgBox "面名を選択してください。", vbExclamation
        cboMen.SetFocus
        Exit Sub
    End If

    Dim selectedMenName As String
    selectedMenName = cboMen.Value

    If Trim(selectedMenName) = "" Then
        MsgBox "面名を選択してください。", vbExclamation
        cboMen.SetFocus
        Exit Sub
    End If

    ' --- フォームを閉じてから処理を実行 ---
    Unload Me
    Call ExecuteBunkatsu(selectedMenName)
End Sub

'---------------------------------------------------------------------
' キャンセルボタン押下時の処理
'---------------------------------------------------------------------
Private Sub btnCancel_Click()
    Unload Me
End Sub

'---------------------------------------------------------------------
' フォーム右上×ボタンでも安全に閉じる
'---------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ' ×ボタンで閉じた場合は何もせず終了
        Cancel = False
    End If
End Sub
