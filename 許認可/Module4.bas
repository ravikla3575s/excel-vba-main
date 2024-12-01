Sub UpdatePharmacistInfo()
    Dim Kws As Worksheet, updWS As Worksheet
    Set Kws = ThisWorkbook.Sheets("所属変更")
    Set updWS = ThisWorkbook.Sheets("届出一覧テーブル")
    
    Dim updateRow As Long
    Dim updateColumn As Long
    Dim startColumn As Long
    Dim strPharmacistArray(1 To 10) As String
    Dim pharmacistInfoColumn As Long
    Dim i As Long, j As Long
    
    ' findUpdateRow関数の呼び出し
    updateRow = findUpdateRow(ws)
    
    ' findUpdateColumn関数の呼び出し
    startColumn = findUpdateColumn(ws, "非常勤薬剤師10")
    
    ' 薬剤師情報を仮に設定
    For i = 1 To 10
        strPharmacistArray(i) = "Pharmacist" & i
    Next i
    
    ' 基準カラムの右隣から30カラムの範囲を使用
    For i = 0 To 9
        pharmacistInfoColumn = startColumn + (i * 3) ' 3つずつの範囲に入力
        ws.Cells(updateRow, pharmacistInfoColumn).value = strPharmacistArray(i + 1)
        ws.Cells(updateRow, pharmacistInfoColumn + 1).value = "その他薬剤師" & (i + 1)
        ws.Cells(updateRow, pharmacistInfoColumn + 2).value = "Pharmacist Info " & (i + 1)
    Next i
    
End Sub

Function findUpdateRow(ws As Worksheet) As Long
    Dim lastRow As Long
    Dim i As Long
    Dim storeName As String
    storeName = "対象店舗名" ' ここに店舗名を設定
    
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    
    For i = 1 To lastRow
        If ws.Cells(i, 2).value = storeName Then
            findUpdateRow = i
            Exit Function
        End If
    Next i
    
    findUpdateRow = 0 ' 店舗名が見つからなかった場合
End Function

Function findUpdateColumn(ws As Worksheet, headerName As String) As Long
    Dim lastColumn As Long
    Dim i As Long
    
    lastColumn = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    For i = 1 To lastColumn
        If ws.Cells(1, i).value = headerName Then
            findUpdateColumn = i + 1 ' 基準カラムの右隣
            Exit Function
        End If
    Next i
    
    findUpdateColumn = 0 ' 基準カラムが見つからなかった場合
End Function
