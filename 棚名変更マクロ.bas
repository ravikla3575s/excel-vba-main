Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' A5:A500範囲でのみ実行
    If Not Intersect(Target, Me.Range("A5:A500")) Is Nothing Then
        ' エンターキーを押した時のみSearchOnEnterを実行
        Application.OnKey "~", "SearchOnEnter"
    Else
        ' 他のセルを選択したときはエンターキーの設定を解除
        Application.OnKey "~"
    End If
End Sub

Sub SearchOnEnter()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim searchText As String
    Dim i As Long
    Dim isJANCode As Boolean
    
    ' 現在のセルがA5:A500の範囲内であることを確認
    If Not Intersect(ActiveCell, ThisWorkbook.Sheets("Sheet1").Range("A5:A500")) Is Nothing Then
        searchText = ActiveCell.Value
        
        ' JANコードの判定（数字のみで構成されるか）
        isJANCode = IsNumeric(searchText) And Len(searchText) >= 3
        
        ' 三文字以上で検索を実行
        If Len(searchText) >= 3 Then
            ' ユーザーフォームのリストボックスをクリア
            frmSearch.lstResults.Clear
            
            ' 検索対象シートの選択
            If isJANCode Then
                Set ws = ThisWorkbook.Worksheets("Sheet3") ' JANコードならSheet3を使用
            Else
                Set ws = ThisWorkbook.Worksheets("tmp_tana") ' 文字列検索ならtmp_tanaシート
            End If
            
            ' 検索シートの最終行を取得
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            ' シートのデータを部分一致で検索し、リストボックスに追加
            For i = 2 To lastRow ' データが2行目から始まると仮定
                If InStr(1, ws.Cells(i, 2).Value, searchText, vbTextCompare) > 0 Then
                    frmSearch.lstResults.AddItem ws.Cells(i, 2).Value
                End If
            Next i
            
            ' ユーザーフォームを表示
            If frmSearch.lstResults.ListCount > 0 Then
                frmSearch.Show vbModal ' モーダル表示
            Else
                MsgBox "該当する項目が見つかりませんでした。"
            End If
        End If
    End If
End Sub

Private Sub cmdSelect_Click()
    ' リストボックスで選択したアイテムをB列に戻す
    If lstResults.ListIndex <> -1 Then ' アイテムが選択されている場合
        ActiveCell.Offset(0, 1).Value = lstResults.Value ' 選択されたセルのB列に入力
        Unload Me ' フォームを閉じる
    Else
        MsgBox "項目を選択してください。"
    End If
End Sub

Sub ExportToCSV(ws As Worksheet)
    Dim csvData As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim filePath As String
    Dim version As Integer
    Dim folderPath As String
    Dim baseFileName As String
    Dim fullFileName As String
    
    ' 保存先のフォルダパスと基本ファイル名を設定
    folderPath = ThisWorkbook.Path & Application.PathSeparator
    baseFileName = "updated_tmp_tana"
    
    ' フォルダ内の最新バージョンを確認し、次のバージョン番号を決定
    version = 1
    Do
        fullFileName = folderPath & baseFileName & "_v" & version & ".csv"
        If Dir(fullFileName) = "" Then Exit Do
        version = version + 1
    Loop
    
    ' シートの最終行と最終列を取得
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' データをCSV形式に変換
    For i = 1 To lastRow
        For j = 1 To lastCol
            csvData = csvData & ws.Cells(i, j).Value
            If j < lastCol Then csvData = csvData & ","
        Next j
        csvData = csvData & vbNewLine
    Next i
    
    ' バージョン付きファイル名で保存
    Open fullFileName For Output As #1
    Print #1, csvData
    Close #1
    
    MsgBox "CSVファイルが保存されました: " & fullFileName
End Sub