' CSVデータを転記
Sub ImportCSVData(csvFile As String, ws As Worksheet, fileType As String)
    Dim colMap As Object
    Dim fso As Object
    Dim ts As Object
    Dim lineText As String
    Dim dataArray As Variant
    Dim i As Long, j As Long, key

    ' 項目マッピングを取得
    Set colMap = GetColumnMapping(fileType)

    ' 1行目に項目名を転記
    j = 1
    For Each key In colMap.Keys
        ws.Cells(1, j).Value = colMap(key)
        j = j + 1
    Next key

    ' CSVデータを読み込んで転記
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(csvFile, 1, False, -2)

    ' データを転記
    i = 2
    Do While Not ts.AtEndOfStream
        lineText = ts.ReadLine
        dataArray = Split(lineText, ",")

        If i > 2 Then
            j = 1
            For Each key In colMap.Keys
                If key <= UBound(dataArray) Then
                    ws.Cells(i - 1, j).Value = Trim(dataArray(key - 1))
                End If
                j = j + 1
            Next key
        End If
        i = i + 1
    Loop
    ts.Close

    ' 列幅を自動調整
    ws.Cells.EntireColumn.AutoFit
End Sub