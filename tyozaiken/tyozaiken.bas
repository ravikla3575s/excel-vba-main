Sub ExportSeihoChouzaiken()
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
    Dim tempAddress As String
    Dim firstPublicCode As String
    Dim secondPublicCode As String
    Dim thirdPublicCode As String
    
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
            
            ' 患者の住所を取得して「旭川市」が含まれているか確認
            tempAddress = FixKanaAndTrim(csvData(i, 38)) ' 38列目: 患者住所
            
            ' 住所に「旭川市」が含まれていない場合はスキップ
            If InStr(tempAddress, "旭川市") <> 0 Then
            End If
            
            ' 第一公費・第二公費の種別番号を取得
            firstPublicCode = csvData(i, 22) ' 第一公費種別番号
            secondPublicCode = csvData(i, 26) ' 第二公費種別番号
            thirdPublicCode = csvData(i, 30) ' 第三公費種別番号
            
            ' 公費の種別に自立支援が第一・第二のどちらにあるか判定
            If firstPublicCode = "21" Or firstPublicCode = "15" Or firstPublicCode = "16" Then
                ws.Cells(rowNum, 12).Value = "◯" ' 第一公費に自立支援がある（K列）
            ElseIf secondPublicCode = "21" Or secondPublicCode = "15" Or secondPublicCode = "16" Then
                ws.Cells(rowNum, 12).Value = "◯" ' 第二公費に自立支援がある（K列）
            End If

            ' 公費の種別に重症が第一・第二のどちらにあるか判定
            If firstPublicCode = "54" Then
    ws.Cells(rowNum, 13).Value = "◯" ' 第一公費に重度がある（L列）
                ws.Cells(rowNum, 13).Value = "◯" ' 第一公費に自立支援がある（K列）
            ElseIf secondPublicCode = "54" Then
    ws.Cells(rowNum, 13).Value = "◯" ' 第二公費に重度がある（L列）
                ws.Cells(rowNum, 13).Value = "◯" ' 第二公費に自立支援がある（K列）
            End If
            
            ' 医療機関名・診療科・住所情報の転記
            ws.Cells(rowNum, 2).Value = ThisWorkbook.Sheets(1).Cells(1, 2).Value ' 薬局名 ' 医療機関名
            ws.Cells(rowNum, 3).Value = ThisWorkbook.Sheets(1).Cells(2, 2).Value ' 医療機関コード ' 診療科
            ws.Cells(rowNum, 4).Value = TrimSpaces(FixKana(csvData(i, 32))) ' 医療機関名
            ws.Cells(rowNum, 5).Value = TrimSpaces(FixKana(csvData(i, 65))) ' 医療機関コード
            ws.Cells(rowNum, 6).Value = TrimSpaces(FixKana(csvData(i, 58))) ' 受給者番号
            ws.Cells(rowNum, 7).Value = FixKanaAndTrim(csvData(i, 10)) ' 患者氏名
            ws.Cells(rowNum, 8).Value = FixKanaAndTrim(csvData(i, 11)) ' 患者カナ氏名
            ws.Cells(rowNum, 9).Value = TrimSpaces(FixKana(csvData(i, 12))) ' 生年月日
            ws.Cells(rowNum, 10).Value = TrimSpaces(FixKana(csvData(i, 56))) ' 診療年月日
            
            ' 生保患者受給者番号の転記
            ws.Cells(rowNum, 13).Value = IIf(firstPublicCode = "54" Or secondPublicCode = "54", "◯", "")
            
            ' 行番号を進める
            rowNum = rowNum + 1
        End If
        End If
Next i
    
    ' 新しいブックを作成し、シートをコピーし新しいブックとして作成
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
    saveFileName = "tyouzai_excel_2.xlsx"
    savePath = folderPath & "\" & saveFileName
    
    ' 新しいファイルとして保存
    newWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close False
    
    ' 呼び出し元シート2のB11:M500をクリア
    callingWs.Range("B11:M500").ClearContents
    
    ' 完了メッセージ
    MsgBox "データの転記と保存が完了しました。保存場所: " & savePath, vbInformation
End Sub
Function FixKana(inputStr As String) As String
    Dim result As String
    result = Application.WorksheetFunction.Substitute(inputStr, "'","") ' シングルクォートを削除
    result = Application.WorksheetFunction.Substitute(result, "(", "/")
    result = Application.WorksheetFunction.Substitute(result, ")", "")
    result = StrConv(result, vbWide) ' 半角カナを全角に変換
    FixKana = result
End Function

Function TrimSpaces(inputStr As String) As String
    TrimSpaces = Application.WorksheetFunction.Trim(inputStr) ' スペース削除
End Function
