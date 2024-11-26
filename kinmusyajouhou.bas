Sub ImportDataAndTransfer()

    ' ファイル選択ダイアログを表示
    Dim selectedFile As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim dataArray(1 To 13) As Variant
    Dim i As Long
    
    ' ファイル選択ダイアログを表示
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "勤務者情報ファイルを選択してください"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
        Else
            MsgBox "ファイルが選択されませんでした。", vbExclamation
            Exit Sub
        End If
    End With

    ' 選択されたファイルを開く
    On Error GoTo ErrorHandler
    Set wbSource = Workbooks.Open(selectedFile)
    Set wsSource = wbSource.Sheets(1) ' 勤務者情報ファイルの最初のシートを参照
    
    ' データを配列に格納
    dataArray(1) = wsSource.Range("B3").Value ' 社員番号
    dataArray(2) = wsSource.Range("B4").Value ' 氏名
    dataArray(3) = wsSource.Range("B5").Value ' シメイ
    dataArray(4) = wsSource.Range("B6").Value ' 保健薬剤師記号
    dataArray(5) = wsSource.Range("B7").Value ' 保健薬剤師登録番号
    dataArray(6) = wsSource.Range("B8").Value ' 薬剤師番号
    dataArray(7) = wsSource.Range("B9").Value ' 薬剤師番号登録日
    dataArray(8) = wsSource.Range("B10").Value ' 生年月日
    dataArray(9) = wsSource.Range("B11").Value ' 郵便番号
    dataArray(10) = wsSource.Range("B12").Value ' 都道府県
    dataArray(11) = wsSource.Range("B13").Value ' 住所
    dataArray(12) = wsSource.Range("B14").Value ' 週労働時間
    dataArray(13) = wsSource.Range("B15").Value ' 資格者区分
    
    ' 元ファイルを閉じる
    wbSource.Close SaveChanges:=False
    
    ' 転記先シートを設定
    Set wsTarget = ThisWorkbook.Sheets("届出一覧テーブル")
    
    ' データを転記
    For i = 1 To 13
        wsTarget.Cells(i + 1, 2).Value = dataArray(i) ' B列に転記（1行目はヘッダー想定）
    Next i

    MsgBox "データの転記が完了しました。", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    On Error GoTo 0
End Sub