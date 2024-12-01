Sub ProcessSupporterFilesAndCreatePDFs()
    Dim folderPath As String
    Dim exfileName As String
    Dim wb As Workbook
    Dim Kws As Worksheet
    Dim supporterData As Variant
    Dim i As Long
    Dim supporterName As String
    Dim startDate As String
    Dim endDate As String
    Dim parsedStartDate As Variant
    Dim parsedEndDate As Variant
    Dim createdStartDate As Variant
    Dim createdEndDate As Variant
    Dim storeName As String
    Dim nameParts() As String
    
    ' Excelファイルのあるフォルダパスを指定
    folderPath = "/Users/yoshipc/Desktop/令和6年3月応援者リスト/" ' <-- フォルダパスを適宜変更してください
    
    ' 所属変更シートの設定
    Set Kws = ThisWorkbook.Sheets("所属変更")
    
    ' フォルダ内の最初のファイルを取得
    exfileName = Dir(folderPath & "*.xlsx")
    Debug.Print "1st ：" & Dir()
    ' フォルダ内の全てのファイルをループ
    'On Error GoTo ErrLabel
    Do While exfileName <> ""
        ' 各ファイルを開く
        Set wb = Workbooks.Open(folderPath & exfileName)
        
        storeName = wb.Sheets(1).Range("A1").value
        ' A1セルの店舗名を取得し、"　" または " " で区切って左側の文字列を店舗名とする
       If storeName = "" Then
            storeName = 0
        End If
        nameParts = Split(storeName, "　") ' 全角スペースで分割
        
        If UBound(nameParts) = 0 Then
            nameParts = Split(storeName, " ") ' 半角スペースで再分割
        End If
        
        storeName = nameParts(0) ' 左側の文字列を取得
        
        ' 店舗名の末尾に"店"が付いていない場合、"店"を付加
        If Right(storeName, 1) <> "店" Then
            storeName = storeName & "店"
        End If
        
        storeName = FindPartialMatchStoreName(storeName, ThisWorkbook.Sheets("届出一覧テーブル").Range("B2:B67"))
        
        ' 店舗名をA2セルにセット
        Kws.Cells(2, 1).value = storeName
        Kws.Range("E2").value = "非常勤"
        Kws.Range("B3:D11").ClearContents
        Kws.Range("B13:B17").ClearContents
        
        ' 新書式か旧書式かを判定
        If IsNewFormat(wb.Sheets(1)) Then
            ' 新書式の場合の処理
            supporterData = GetSupporterDataFromSheet(wb.Sheets(1))
            
            ' データを所属変更シートに反映
            For i = LBound(supporterData, 1) + 2 To UBound(supporterData, 1)
                supporterName = supporterData(i, 1)
                startDate = supporterData(i, 2)
                endDate = supporterData(i, 3)
                
                If supporterName = "" Then
                    Exit For
                End If
                
                ' startDate の処理
                If IsDate(startDate) Then
                    createdStartDate = startDate
                Else
                    parsedStartDate = ParseAndValidateDates(startDate)
                    If IsArray(parsedStartDate) Then
                        createdStartDate = parsedStartDate(0) ' 最初の日付を使用
                    End If
                End If
                
                ' endDate の処理
                If endDate <> "" Then
                    If IsDate(endDate) Then
                        createdEndDate = endDate
                    Else
                        parsedEndDate = ParseAndValidateDates(endDate)
                        If IsArray(parsedEndDate) Then
                            createdEndDate = parsedEndDate(UBound(parsedEndDate)) ' 最後の日付を使用
                        End If
                    End If
                Else
                    If IsArray(parsedStartDate) Then
                        createdEndDate = parsedStartDate(UBound(parsedStartDate))
                    Else: createdEndDate = createdStartDate
                    End If
                End If
                
                ' 所属変更シートにデータを更新
                UpdateSupporterInSheet Kws, supporterName, createdStartDate, createdEndDate
            Next i
        Else
            ' 旧書式の場合の処理
            supporterName = wb.Sheets(1).Range("C4").value
            Call SortAndIndexDates(wb.Sheets(1))
            startDate = wb.Sheets(1).Range("B4").value
            endDate = wb.Sheets(1).Range("B4").End(xlDown).value
            
            ' startDate の処理
            If IsDate(startDate) Then
                createdStartDate = startDate
            Else
                parsedStartDate = ParseAndValidateDates(startDate)
                If IsArray(parsedStartDate) Then
                    createdStartDate = parsedStartDate(0) ' 最初の日付を使用
                End If
            End If
            
            ' endDate の処理
                If endDate <> "" Then
                    If IsDate(endDate) Then
                        createdEndDate = endDate
                    Else
                        parsedEndDate = ParseAndValidateDates(endDate)
                        If IsArray(parsedEndDate) Then
                            createdEndDate = parsedEndDate(UBound(parsedEndDate)) ' 最後の日付を使用
                        End If
                    End If
                Else
                    If IsArray(parsedStartDate) Then
                        createdEndDate = parsedStartDate(UBound(parsedStartDate))
                    Else: createdEndDate = createdStartDate
                    End If
                End If
            
            ' 所属変更シートにデータを更新
            Call UpdateSupporterInSheet(Kws, supporterName, createdStartDate, createdEndDate)
        End If
        'Debug.Print "2nd ：" & Dir()
        ' B12セル〜B16セルを上から検索し、値が入っている場合はB13から順に移し取る
        Call CopyValuesToThisWorkbook(wb.Sheets(1), Kws)
        
        ' 処理が終わったらファイルを閉じる
        wb.Close False
        
        Call UpdateMultiplePharmacists
        ' PDFを作成（必要に応じて）
        ThisWorkbook.Activate
        Call 厚生局所属変更書類PDF
        'Debug.Print "3d ：" & Dir()
        ' 次のファイルを取得
        exfileName = Dir()
    Loop
    
    Kws.Range("B3:D11").ClearContents
    Kws.Range("B13:B17").ClearContents
    Exit Sub
ErrLabel:
    MsgBox "エラーが発生しました"
    wb.Close False
    exfileName = Dir()
End Sub
Function ParseAndValidateDates(dateStr As String, Optional baseMonth As Long = 0, Optional baseYear As Long = 0) As Variant
    Dim dateParts() As String
    Dim parsedDates() As Variant
    Dim i As Integer
    Dim tempDate As Date
    
    ' 複数の区切り文字で日付を分割
    dateStr = Replace(dateStr, "ー", "-")
    dateStr = Replace(dateStr, "〜", "-")
    dateStr = Replace(dateStr, ".", "-")
    dateStr = Replace(dateStr, ",", "-")
    
    dateParts = Split(dateStr, "-")
    
    ' 分割された日付を格納する配列の初期化
    ReDim parsedDates(LBound(dateParts) To UBound(dateParts))
    
    'On Error GoTo ErrorHandler
    
    For i = LBound(dateParts) To UBound(dateParts)
        ' 日付が月/日のみの場合、ベース月と年を使用して完全な日付にする
        If InStr(dateParts(i), "/") > 0 Then
            If baseMonth > 0 And baseYear > 0 Then
                tempDate = DateSerial(baseYear, baseMonth, CInt(Split(dateParts(i), "/")(1)))
                If IsDate(tempDate) Then
                    parsedDates(i) = tempDate
                End If
            Else
                parsedDates(i) = CDate(dateParts(i))
            End If
        ElseIf IsDate(dateParts(i)) Then
            ' 年月日がフルに入っている場合
            parsedDates(i) = CDate(dateParts(i))
        Else
            ' 日付の形式が不正な場合はエラーメッセージを返す
            ParseAndValidateDates = "Invalid date in string: " & dateParts(i)
            Exit Function
        End If
    Next i
    
    ' 正常終了した場合、日付の配列を返す
    ParseAndValidateDates = parsedDates
    Exit Function
    
ErrorHandler:
    ' エラーが発生した場合、エラーメッセージを返す
    ParseAndValidateDates = "Invalid date in string: " & dateParts(i)
End Function
Function FindPartialMatchStoreName(storeName As String, searchRange As Range) As String
    Dim cell As Range
    
    ' 検索範囲をループ
    For Each cell In searchRange
        If InStr(1, cell.value, storeName, vbTextCompare) > 0 Then
            ' 部分一致する店舗名を見つけた場合、その名前を返す
            FindPartialMatchStoreName = cell.value
            Exit Function
        End If
    Next cell
    
    ' 一致する名前が見つからなかった場合は空の文字列を返す
    FindPartialMatchStoreName = ""
End Function
Function IsNewFormat(ws As Worksheet) As Boolean
    ' D1セルの値をチェックして、"←店舗名を入力してください"であるかを判定します
    If ws.Range("D1").value = "←店舗名を入力してください" Then
        IsNewFormat = True  ' 新書式
    Else
        IsNewFormat = False ' 旧書式
    End If
End Function

Sub SortAndIndexDates(ws As Worksheet)
    Dim rng As Range
    Dim cell As Range
    Dim dateArray() As Date
    Dim i As Long, j As Long
    Dim tempDate As Date
    
    ' 日付が入力されている範囲（B列）を指定
    Set rng = ws.Range("B4:B" & ws.Cells(ws.Rows.count, "B").End(xlUp).Row)
    
    ' 日付を配列に格納
    ReDim dateArray(1 To rng.Rows.count)
    i = 1
    For Each cell In rng
        If IsDate(cell.value) Then
            dateArray(i) = CDate(cell.value)
            i = i + 1
        End If
    Next cell
    
    ' 日付配列をソート
    For i = LBound(dateArray) To UBound(dateArray) - 1
        For j = i + 1 To UBound(dateArray)
            If dateArray(i) > dateArray(j) Then
                tempDate = dateArray(i)
                dateArray(i) = dateArray(j)
                dateArray(j) = tempDate
            End If
        Next j
    Next i
    
    ' ソートされた日付をシートに反映
    i = 1
    For Each cell In rng
        cell.value = dateArray(i)
        i = i + 1
    Next cell
End Sub

Sub UpdateSupporterInSheet(Kws As Worksheet, Name As String, startDate As Variant, endDate As Variant)
    Dim lastRow As Long
    Dim found As Range
    
    ' 名前がすでにシートにあるかを確認
    Set found = Kws.Columns("B").Find(Name, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' 日付データがDate型でない場合、変換
    If Not IsDate(startDate) Then
        startDate = CDate(startDate)
    End If
    If Not IsDate(endDate) Then
        endDate = CDate(endDate)
    End If
    
    If Not found Is Nothing Then
        ' 名前が見つかった場合、最終日付を更新
        If found.Offset(0, 2).value < endDate Then
            found.Offset(0, 2).value = endDate
        End If
    Else
        ' 新しい行に追加
        lastRow = Kws.Cells(11, "B").End(xlUp).Row + 1
        Kws.Cells(lastRow, 2).value = Name
        Kws.Cells(lastRow, 3).value = startDate
        Kws.Cells(lastRow, 4).value = endDate
    End If
End Sub

Sub CopyValuesToThisWorkbook(srcWs As Worksheet, destWs As Worksheet)
    Dim i As Long
    For i = 12 To 16
        If srcWs.Cells(i, 2).value <> "" Then
            destWs.Cells(i - 11 + 12, 2).value = srcWs.Cells(i, 2).value
        End If
    Next i
End Sub

Function GetSupporterDataFromSheet(ws As Worksheet) As Variant
    Dim lastRow As Long
    Dim dataRange As Range
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    
    ' データ範囲を設定
    Set dataRange = ws.Range("B2:D" & lastRow)
    
    ' データを配列に変換して返す
    GetSupporterDataFromSheet = dataRange.value
End Function

Sub ExampleDir()
    Dim folderPath As String
    Dim fileName As String
    
    ' 検索するフォルダを指定
    folderPath = "/Users/yoshipc/Desktop/令和6年3月応援者リスト/"
    
    ' 最初のファイルを取得
    fileName = Dir(folderPath & "*.xlsx")
    
    ' ループでフォルダ内のすべての .xlsx ファイルを取得
    Do While fileName <> ""
        ' ファイル名を出力
        Debug.Print "Found file: " & fileName
        
        ' 次のファイルを取得
        fileName = Dir
    Loop
End Sub

