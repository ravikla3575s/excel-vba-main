Sub SearchDatesAndInsertValues()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim currentYear As Integer
    Dim searchDate As Date
    Dim i As Integer
    Dim insertRow As Long
    Dim targetValue As Variant
    
    ' シートの参照設定
    Set ws1 = ThisWorkbook.Sheets(1) ' Sheet(1)
    Set ws2 = ThisWorkbook.Sheets(2) ' Sheet(2)
    
    ' 現在の年を取得
    currentYear = Year(Date)
    
    ' 値を挿入する開始行（Sheet(1)のCells(2, 4)の位置を基準）
    insertRow = ws1.Cells(2, 4).Row
    
    ' Sheet(2)のCells(43, 3)からCells(43, 36)を検索
    For i = 3 To 36
        searchDate = ws2.Cells(43, i).Value ' 43行目のi列の値を取得
        
        ' 年が現在の年と同じかどうかを確認
        If Year(searchDate) = currentYear Then
            ' 年が一致する場合、同じ列の5行目の値を取得
            targetValue = ws2.Cells(5, i).Value
            
            ' Sheet(1)の挿入行にその値を入力
            ws1.Cells(insertRow, 4).Value = targetValue
            
            ' 次の行に値を挿入するため、挿入行を1つ下げる
            insertRow = insertRow + 1
        End If
    Next i
    
    MsgBox "該当する値の検索と挿入が完了しました。"
End Sub
Sub SearchStringsAndDisplayMismatch()

    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim searchRange As Range, cell As Range
    Dim searchValue As String
    Dim foundCell As Range
    Dim msg As String
    Dim lastCol As Long
    Dim lastRow As Long

    ' シートを設定
    Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Set ws2 = ThisWorkbook.Sheets("Sheet2")

    ' シート2の6行目の最終列を取得
    lastCol = ws2.Cells(6, ws2.Columns.Count).End(xlToLeft).Column

    ' シート1のD列5行目以降の範囲を設定
    lastRow = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row
    Set searchRange = ws1.Range("D5:D" & lastRow)

    ' 一致しない文字列を格納するメッセージ用変数
    msg = ""

    ' シート2の6行目を順に検索
    For col = 3 To lastCol
        searchValue = ws2.Cells(6, col).Value

        ' シート1のD列で文字列を検索
        Set foundCell = searchRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)

        ' 一致しなかった場合、メッセージボックスに表示
        If foundCell Is Nothing Then
            msg = msg & searchValue & vbNewLine
        End If
    Next col

    ' 一致しなかった文字列を表示
    If msg <> "" Then
        MsgBox "一致しなかった文字列: " & vbNewLine & msg
    Else
        MsgBox "全ての文字列が一致しました。"
    End If

End Sub
Sub ImportCSVAndTransferSpecificData()
    Dim ws As Worksheet
    Dim csvWs As Worksheet
    Dim csvFilePath As String
    Dim csvData As Variant
    Dim i As Long
    Dim rowNum As Long
    Dim lastRow As Long
    
    ' CSVファイルのパスを指定
    csvFilePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "CSVファイルを選択")
    
    If csvFilePath = "False" Then
        MsgBox "CSVファイルが選択されませんでした。"
        Exit Sub
    End If
    
    ' CSVファイルを開く
    Workbooks.Open csvFilePath
    Set csvWs = ActiveSheet
    
    ' データの最終行を取得
    lastRow = csvWs.Cells(csvWs.Rows.Count, 1).End(xlUp).Row
    
    ' データを配列に読み込む
    csvData = csvWs.Range("A2:BR" & lastRow).Value ' 1行目は無視し、2行目から読み込む
    
    ' CSVファイルを閉じる
    Workbooks(csvWs.Parent.Name).Close False
    
    ' 調剤請求書（旭川市）のシートを指定
    Set ws = ThisWorkbook.Sheets("調剤請求書（旭川市）")
    
    ' 転記開始行は11行目
    rowNum = 11
    
    ' CSVデータをシートに転記
    For i = 1 To UBound(csvData, 1) ' 1行目（CSVの2行目）から読み込む
        If csvData(i, 1) <> "" Then ' 空でない行を処理
            ws.Cells(rowNum, 2).Value = csvData(i, 10) ' J列: 患者氏名
            ws.Cells(rowNum, 3).Value = csvData(i, 11) ' K列: 患者カナ氏名
            ws.Cells(rowNum, 4).Value = csvData(i, 12) ' L列: 生年月日
            ws.Cells(rowNum, 5).Value = csvData(i, 17) ' Q列: 公費
            ws.Cells(rowNum, 6).Value = csvData(i, 51) ' AY列: 生保受給者番号
            ws.Cells(rowNum, 7).Value = csvData(i, 38) ' AL列: 患者住所
            ws.Cells(rowNum, 8).Value = csvData(i, 34) ' AH列: 処方元の医療機関名
            ws.Cells(rowNum, 9).Value = csvData(i, 65) ' BM列: 処方元の医療機関コード
            rowNum = rowNum + 1
        End If
    Next i
    ' 新しいブックを作成し、調剤請求書シートのみをコピー
    Set newWb = Workbooks.Add
    ws.Copy Before:=newWb.Sheets(1)
    
    ' 元のワークブックに戻す
    ThisWorkbook.Activate
    
    ' フォルダ選択ダイアログを表示
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    fDialog.Title = "保存するフォルダを選択してください"
    
    If fDialog.Show = -1 Then
        folderPath = fDialog.SelectedItems(1)
    Else
        MsgBox "保存フォルダが選択されませんでした。処理を中止します。"
        Exit Sub
    End If
    
    ' 保存先のファイル名を設定
    saveFileName = "tyouzai_excel.xlsx"
    savePath = folderPath & "\" & saveFileName
    
    ' 新しいファイルとして保存
    newWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close False
    
    ' 呼び出し元シート2のD11:M500をクリア
    callingWs.Range("D11:M500").ClearContents
    
    ' 完了メッセージ
    MsgBox "データの転記と保存が完了しました。保存場所: " & savePath, vbInformation
End Sub
Sub ProcessMultipleCSVFiles()
    Dim ws As Worksheet
    Dim csvSheet As Worksheet
    Dim lastRow As Long
    Dim searchMonth As String
    Dim i As Long, j As Long
    Dim csvMonth As String
    Dim normalStartCol As Integer
    Dim reClaimStartCol As Integer
    Dim found As Boolean
    Dim folderPath As String
    Dim csvFile As String
    Dim wb As Workbook
    
    ' シート1（転記先のシート）
    Set ws = ThisWorkbook.Sheets(1)
    
    ' フォルダ選択ダイアログを表示
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "CSVファイルが保存されているフォルダを選択してください"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "フォルダが選択されませんでした。"
            Exit Sub
        End If
    End With
    
    ' フォルダ内のすべてのCSVファイルを処理
    csvFile = Dir(folderPath & "*.csv") ' フォルダ内のCSVファイルを取得
    
    Do While csvFile <> ""
        ' CSVファイルを開く
        Set wb = Workbooks.Open(folderPath & csvFile)
        Set csvSheet = wb.Sheets(1) ' CSVシートを取得
        
        ' CSVシートの最終行を取得
        lastRow = csvSheet.Cells(csvSheet.Rows.Count, "A").End(xlUp).Row
        
        ' シート1のA5からA16の各月のラベルを検索
        For i = 5 To 16 Step 2 ' A5からA16まで、2行セットで検索するのでStep 2に設定
            searchMonth = ws.Cells(i, 1).Value ' A列の各月のラベルを取得
            
            found = False ' 初期状態では見つかっていない
            
            ' CSVシートの該当するデータを検索
            csvMonth = Replace(csvSheet.Cells(1, 5).Value, "'", "") ' 'を削除
            csvMonth = ConvertZenkakuToHankaku(csvMonth) ' 全角数字を半角に変換
            
            ' 一致するか比較
            If csvMonth = searchMonth Then
                ' 一致した場合にデータを転記
                found = True ' 見つかったことを示す
                
                ' 通常請求分：社保請求データをE列から横方向に転記（縦→横に変換）
                normalStartCol = 5 ' E列
                ws.Cells(i, normalStartCol).Resize(1, 7).Value = WorksheetFunction.Transpose(csvSheet.Cells(3, 11).Resize(7, 1).Value)
                
                ' 再請求分：O列から横方向に転記（縦→横に変換）
                reClaimStartCol = 15 ' O列
                ws.Cells(i + 1, reClaimStartCol).Resize(1, 7).Value = WorksheetFunction.Transpose(csvSheet.Cells(12, 11).Resize(7, 1).Value)
                
                Exit For ' データを転記したらループを抜ける
            End If
            
            ' 該当する月が見つからなかった場合のエラーメッセージ
            If Not found Then
                MsgBox "対象年月日が見つかりません: " & searchMonth & " in " & csvFile, vbExclamation, "エラー"
            End If
        Next i
        
        ' CSVファイルを閉じる
        wb.Close False ' 保存せずに閉じる
        
        ' 次のCSVファイルを取得
        csvFile = Dir
    Loop
    
    MsgBox "すべてのCSVファイルの処理が完了しました。"
End Sub

' 全角数字を半角数字に変換する関数
Function ConvertZenkakuToHankaku(inputStr As String) As String
    Dim i As Integer
    Dim result As String
    Dim currentChar As String
    result = ""
    
    ' 全角の数字を半角に変換
    For i = 1 To Len(inputStr)
        currentChar = Mid(inputStr, i, 1)
        Select Case currentChar
            Case "０": result = result & "0"
            Case "１": result = result & "1"
            Case "２": result = result & "2"
            Case "３": result = result & "3"
            Case "４": result = result & "4"
            Case "５": result = result & "5"
            Case "６": result = result & "6"
            Case "７": result = result & "7"
            Case "８": result = result & "8"
            Case "９": result = result & "9"
            Case Else: result = result & currentChar ' 全角数字以外はそのまま
        End Select
    Next i
    
    ConvertZenkakuToHankaku = result
End Function
