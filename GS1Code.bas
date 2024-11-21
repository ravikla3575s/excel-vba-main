Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim searchRange As Range
    Dim cell As Range
    Dim inputValue As String
    Dim searchValue As String
    Dim result As Variant

    ' 処理範囲をA5:A500に限定
    If Not Intersect(Target, Me.Range("A5:A500")) Is Nothing Then
        Application.EnableEvents = False ' イベントの無限ループを防ぐ

        ' 入力されたセルの値を取得
        For Each cell In Target
            If Not IsEmpty(cell.Value) Then
                inputValue = cell.Value

                ' 入力値が16桁の数字であることを確認
                If Len(inputValue) = 16 And IsNumeric(inputValue) Then
                    Dim thirdDigit As String
                    thirdDigit = Mid(inputValue, 3, 1) ' 左から3桁目の数字を取得

                    ' シート3を参照
                    Set ws = ThisWorkbook.Worksheets("Sheet3")

                    If thirdDigit = "1" Then
                        ' パターン①：右から14桁をG列で検索
                        searchValue = Right(inputValue, 14)
                        Set searchRange = ws.Columns("G")
                    ElseIf thirdDigit = "0" Then
                        ' パターン②：右から13桁をE列で検索
                        searchValue = Right(inputValue, 13)
                        Set searchRange = ws.Columns("E")
                    Else
                        GoTo SkipCell ' 3桁目が1でも0でもない場合は処理をスキップ
                    End If

                    ' 検索して一致する値を取得
                    On Error Resume Next
                    result = Application.Match(searchValue, searchRange, 0)
                    On Error GoTo 0

                    If Not IsError(result) Then
                        ' 一致する値が見つかった場合、その値をB列に代入
                        cell.Offset(0, 1).Value = searchRange.Cells(result).Value
                    Else
                        ' 一致する値が見つからない場合、B列を空白にする
                        cell.Offset(0, 1).Value = ""
                    End If
                Else
                    ' 入力が16桁の数字でない場合はエラー表示
                    MsgBox "入力された値は16桁の数字である必要があります。", vbExclamation, "入力エラー"
                    cell.Value = "" ' 無効な値を削除
                End If
            End If
SkipCell:
        Next cell

        Application.EnableEvents = True ' イベントを再有効化
    End If
End Sub