Function GetStartRow(ws As Worksheet, category As String) As Long
    Dim lastRow As Long
    Dim i As Long
    Dim keyword As String
    
    ' 最終行取得
    lastRow = ws.Cells(Rows.Count, "H").End(xlUp).Row

    ' キーワードに応じた開始行を決定
    Select Case category
        Case "社保返戻再請求"
            keyword = "国家→医本"
        Case "国保返戻再請求"
            keyword = "⑨返戻分再請求分（医保）"
        Case "社保月遅れ請求"
            keyword = "⑨返戻分再請求分"
        Case "国保月遅れ請求"
            keyword = "⑩月遅れ請求分（医保）"
        Case Else
            GetStartRow = lastRow + 1
            Exit Function
    End Select

    ' H列を検索（2行目から）
    For i = 2 To lastRow
        If ws.Cells(i, 8).Value = keyword Then
            GetStartRow = i + 1
            Exit Function
        End If
    Next i

    ' 該当なしの場合
    GetStartRow = lastRow + 1
End Function