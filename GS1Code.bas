Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim searchRange As Range
    Dim cell As Range
    Dim inputValue As String
    Dim searchValue As String
    Dim result As Variant
    Dim lastRow As Long
    Dim searchText As String
    Dim isJANCode As Boolean
    Dim thirdDigit As String

    ' 処理範囲をA5:A500に限定
    If Not Intersect(Target, Me.Range("A5:A500")) Is Nothing Then
        Application.EnableEvents = False ' イベントの無限ループを防ぐ

        ' 入力されたセルの値を取得
        For Each cell In Target
            If Not IsEmpty(cell.Value) Then
                inputValue = cell.Value

                ' 入力値が数字のみかどうかを判定（0で始まる可能性を考慮）
                isJANCode = IsNumeric(inputValue)

                ' 入力値が3文字以上の場合、三文字検索を実行
                If Len(inputValue) >= 3 Then
                    searchText = inputValue

                    ' ユーザーフォームのリストボックスをクリア
                    frmSearch.lstResults.Clear

                    ' シート3を使用（数字の場合）
                    If isJANCode Then
                        Set ws = ThisWorkbook.Worksheets(3)
                    Else
                        Set ws = ThisWorkbook.Worksheets("tmp_tana")
                    End If

                    ' 検索シートの最終行を取得
                    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

                    ' 部分一致検索
                    For i = 2 To lastRow
                        If InStr(1, ws.Cells(i, 2).Value, searchText, vbTextCompare) > 0 Then
                            frmSearch.lstResults.AddItem ws.Cells(i, 2).Value
                        End If
                    Next i

                    ' リストボックスの結果を表示
                    If frmSearch.lstResults.ListCount > 0 Then
                        frmSearch.Show vbModal
                    Else
                        MsgBox "該当する項目が見つかりませんでした。"
                    End If
                End If

                ' GS1Code（16桁以上の数字）が入力された場合、完全一致検索
                If isJANCode And Len(inputValue) = 16 Then
                    thirdDigit = Mid(inputValue, 3, 1) ' 左から3桁目を取得

                    ' シート3を参照
                    Set ws = ThisWorkbook.Worksheets("Sheet3")

                    If thirdDigit = "1" Then
                        ' パターン①：右から14桁をG列で完全一致検索
                        searchValue = Right(inputValue, 14)
                        Set searchRange = ws.Columns("G")
                    ElseIf thirdDigit = "0" Then
                        ' パターン②：右から13桁をE列で完全一致検索
                        searchValue = Right(inputValue, 13)
                        Set searchRange = ws.Columns("E")
                    Else
                        GoTo SkipCell ' 3桁目が1でも0でもない場合は処理をスキップ
                    End If

                    ' 完全一致検索
                    On Error Resume Next
                    result = Application.Match(searchValue, searchRange, 0)
                    On Error GoTo 0

                    If Not IsError(result) Then
                        ' 一致する値が見つかった場合、その値をB列に代入
                        cell.Offset(0, 1).Value = searchRange.Cells(result).Value
                    Else
                        ' 一致する値が見つからない場合、B列を空白にする
                        cell.Offset(0, 1).Value = "一致なし"
                    End If
                End If
            End If
SkipCell:
        Next cell

        Application.EnableEvents = True ' イベントを再有効化
    End If
End Sub