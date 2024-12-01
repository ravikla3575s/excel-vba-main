Sub UpdatePharmacistInfoWithClass()
    Dim ws As Worksheet, Kws As Worksheet
    Set ws = ThisWorkbook.Sheets("届出一覧テーブル")
    Set Kws = ThisWorkbook.Sheets("所属変更")
    Dim strName As String
    Dim updateRow As Long
    Dim startColumn As Long
    Dim pharmacists() As Class1
    Dim i As Long, j As Long, count As Long
    
    '店名をstrNameに入れる
    strName = Kws.Cells(2, 1)
    
    ' findUpdateRow関数の呼び出し
    updateRow = findUpdateRow(ws, strName)
    
    ' 開始カラム (開発時は140列目)
    For startColumn = 1 To 215
    If ws.Cells(1, startColumn) = "非常勤薬剤師10" Then
        startColumn = startColumn + 1
        Exit For
    End If
    Next startColumn
    
    ' 10人分のデータをクラスに格納
    ReDim pharmacists(1 To 10)
    
    For i = 1 To 10
        Set pharmacists(i) = New Class1
        
        ' EmployeeNumberの処理
        empNum = ws.Cells(updateRow, startColumn + (i - 1) * 3).value
        If IsNumeric(empNum) And Len(empNum) <= 7 Then
            pharmacists(i).EmployeeNumber = CLng(empNum)
        Else
            pharmacists(i).EmployeeNumber = 0 ' 無効な場合は0に設定
        End If
        
        ' PharmacistNameの処理
        pharmacists(i).PharmacistName = ws.Cells(updateRow, startColumn + (i - 1) * 3 + 1).value
        
        ' WorkHourの処理
        workHr = ws.Cells(updateRow, startColumn + (i - 1) * 3 + 2).value
        If IsNumeric(workHr) Then
            pharmacists(i).WorkHour = CSng(workHr)
        Else
            pharmacists(i).WorkHour = 0 ' 無効な場合は0に設定
        End If
    Next i
    
    ' 空のインスタンスを詰める処理
    count = 0
    For i = 1 To 10
        If pharmacists(i).EmployeeNumber <> 0 Or Len(pharmacists(i).PharmacistName) > 0 Or pharmacists(i).WorkHour <> 0 Then
            count = count + 1
            If count <> i Then
                Set pharmacists(count) = pharmacists(i)
            End If
        End If
    Next i
    
    ' 整理された配列を再度シートに書き込む
    For i = 1 To count
        ws.Cells(updateRow, startColumn + (i - 1) * 3).value = pharmacists(i).EmployeeNumber
        ws.Cells(updateRow, startColumn + (i - 1) * 3 + 1).value = pharmacists(i).PharmacistName
        ws.Cells(updateRow, startColumn + (i - 1) * 3 + 2).value = pharmacists(i).WorkHour
    Next i
    
    ' 残りのセルをクリアする
    For i = count + 1 To 10
        ws.Cells(updateRow, startColumn + (i - 1) * 3).ClearContents
        ws.Cells(updateRow, startColumn + (i - 1) * 3 + 1).ClearContents
        ws.Cells(updateRow, startColumn + (i - 1) * 3 + 2).ClearContents
    Next i

    Dim fullTimePharmacists() As String
    Dim partTimePharmacists() As String
    Dim fullTimeCount As Long
    Dim partTimeCount As Long
    Dim col As Long

    ' 常勤/非常勤の分類用配列を初期化
    ReDim fullTimePharmacists(1 To 10)
    ReDim partTimePharmacists(1 To 5)
    fullTimeCount = 0
    partTimeCount = 0

    ' 常勤/非常勤を分類
    For i = 1 To count
        If pharmacists(i).WorkHour > 32 Then
            fullTimeCount = fullTimeCount + 1
            fullTimePharmacists(fullTimeCount) = pharmacists(i).PharmacistName
        Else
            partTimeCount = partTimeCount + 1
            partTimePharmacists(partTimeCount) = pharmacists(i).PharmacistName
        End If
    Next i

    ' 常勤薬剤師の登録
    col = findColumn(ws, "常勤薬剤師1")
    If col > 0 Then
        For i = 1 To fullTimeCount
            ws.Cells(updateRow, col + (i - 1)).value = fullTimePharmacists(i)
        Next i
        ' 残りのセルをクリアする
        For i = fullTimeCount + 1 To 10
            ws.Cells(updateRow, col + (i - 1)).ClearContents
        Next i
    End If

    ' 非常勤薬剤師の登録
    col = findColumn(ws, "非常勤薬剤師1")
    If col > 0 Then
        For i = 1 To partTimeCount
            ws.Cells(updateRow, col + (i - 1)).value = partTimePharmacists(i)
        Next i
        ' 残りのセルをクリアする
        For i = partTimeCount + 1 To 5
            ws.Cells(updateRow, col + (i - 1)).ClearContents
        Next i
    End If

End Sub

' カラムを見つける関数
Function findColumn(ws As Worksheet, headerName As String) As Long
    Dim lastColumn As Long
    Dim i As Long
    
    lastColumn = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    For i = 1 To lastColumn
        If ws.Cells(1, i).value = headerName Then
            findColumn = i
            Exit Function
        End If
    Next i
    
    findColumn = 0 ' 見つからなかった場合
End Function
Function findUpdateRow(ws As Worksheet, strName As String) As Long
    Dim lastRow As Long
    Dim i As Long
    Dim storeName As String
    storeName = strName ' ここに店舗名を設定
    
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    
    For i = 1 To lastRow
        If ws.Cells(i, 2).value = storeName Then
            findUpdateRow = i
            Exit Function
        End If
    Next i
    
    findUpdateRow = 0 ' 店舗名が見つからなかった場合
End Function
