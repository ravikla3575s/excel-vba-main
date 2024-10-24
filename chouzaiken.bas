Sub ImportCSVAndTransferDataAndSaveWithKanaFix()
    Dim ws As Worksheet
    Dim csvWs As Worksheet
    Dim newWb As Workbook
    Dim csvFilePath As String
    Dim csvData As Variant
    Dim i As Long
    Dim rowNum As Long
    Dim lastRow As Long
    Dim folderPath As String
    Dim saveFileName As String
    Dim savePath As String
    Dim fDialog As FileDialog
    Dim callingWs As Worksheet
    Dim tempData As String
    
    ' 呼び出し元のシート2を設定
    Set callingWs = ThisWorkbook.Sheets(2)
    
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
            ' 各データを変換処理（全角変換、シングルクォートとスペースの削除）
            ws.Cells(rowNum, 2).Value = FixKanaAndTrim(csvData(i, 10)) ' J列: 患者氏名
            ws.Cells(rowNum, 3).Value = FixKanaAndTrim(csvData(i, 11)) ' K列: 患者カナ氏名
            ws.Cells(rowNum, 4).Value = FixKanaAndTrim(csvData(i, 12)) ' L列: 生年月日
            ws.Cells(rowNum, 5).Value = FixKanaAndTrim(csvData(i, 17)) ' Q列: 公費
            ws.Cells(rowNum, 6).Value = FixKanaAndTrim(csvData(i, 51)) ' AY列: 生保受給者番号
            ws.Cells(rowNum, 7).Value = FixKanaAndTrim(csvData(i, 38)) ' AL列: 患者住所
            ws.Cells(rowNum, 8).Value = FixKanaAndTrim(csvData(i, 34)) ' AH列: 処方元の医療機関名
            ws.Cells(rowNum, 9).Value = FixKanaAndTrim(csvData(i, 65)) ' BM列: 処方元の医療機関コード
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

' 半角カナを全角に変換し、シングルクォートと半角スペースを削除する関数
Function FixKanaAndTrim(inputStr As String) As String
    Dim result As String
    result = Application.WorksheetFunction.Substitute(inputStr, "'", "") ' シングルクォートを削除
    result = Application.WorksheetFunction.Substitute(result, " ", "") ' 半角スペースを削除
    result = StrConv(result, vbWide) ' 半角カナを全角に変換
    FixKanaAndTrim = result
End Function