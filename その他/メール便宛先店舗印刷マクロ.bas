Sub ProcessAndPrintWithLinks()
    Dim controlSheet As Worksheet
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim wbSource As Workbook
    Dim sourceFilePath As String
    Dim sourceSheetName As String
    Dim targetSheetName As String
    Dim lastRow As Long
    Dim i As Long
    Dim formattedText As String
    Dim targetCell As Range
    Dim columnH As Integer
    Dim columnI As Integer
    
    ' デフォルトの列設定（H列: 8, I列: 9）
    columnH = 8
    columnI = 9
    
    ' マクロ起動ブックの操作用シートを取得
    On Error Resume Next
    Set controlSheet = ThisWorkbook.Sheets("操作用シート")
    On Error GoTo 0
    If controlSheet Is Nothing Then
        MsgBox "操作用シートが見つかりません。'操作用シート'を作成してください。", vbExclamation
        Exit Sub
    End If
    
    ' 紐付けデータを取得
    sourceFilePath = controlSheet.Range("B2").Value ' 引用元のファイルパス
    sourceSheetName = controlSheet.Range("B3").Value ' 引用元のシート名
    targetSheetName = controlSheet.Range("B4").Value ' 印刷フォーマット用シート名
    
    ' 引用元のブックを開く（紐付けが正しいかチェック）
    If Dir(sourceFilePath) = "" Or sourceSheetName = "" Then
        ' 紐付け情報が不正の場合は選択を促す
        Dim fileDialog As FileDialog
        Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
        fileDialog.Filters.Clear
        fileDialog.Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        fileDialog.AllowMultiSelect = False
        
        If fileDialog.Show = -1 Then
            sourceFilePath = fileDialog.SelectedItems(1)
            controlSheet.Range("B2").Value = sourceFilePath ' 選択したパスを保存
        Else
            MsgBox "ファイルが選択されませんでした。", vbExclamation
            Exit Sub
        End If
        
        ' シート名を入力させる
        sourceSheetName = Application.InputBox("引用元のシート名を入力してください", "シート選択", Type:=2)
        If sourceSheetName = "" Then Exit Sub
        controlSheet.Range("B3").Value = sourceSheetName ' 選択したシート名を保存
    End If
    
    ' 印刷フォーマット用シートのチェック
    If targetSheetName = "" Then
        targetSheetName = Application.InputBox("印刷フォーマット用のシート名を入力してください", "シート選択", Type:=2)
        If targetSheetName = "" Then Exit Sub
        controlSheet.Range("B4").Value = targetSheetName ' 選択したシート名を保存
    End If
    
    ' ブックを開く
    Set wbSource = Workbooks.Open(sourceFilePath)
    
    ' 引用元シートを取得
    On Error Resume Next
    Set wsSource = wbSource.Sheets(sourceSheetName)
    On Error GoTo 0
    If wsSource Is Nothing Then
        MsgBox "指定された引用元シートが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 印刷フォーマット用シートを取得
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(targetSheetName)
    On Error GoTo 0
    If wsTarget Is Nothing Then
        MsgBox "指定された印刷フォーマット用シートが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.Count, columnH).End(xlUp).Row
    
    ' 処理ループ
    For i = 2 To lastRow ' ヘッダーを飛ばして2行目から処理
        Dim storeNumber As String
        Dim storeName As String
        
        ' H列の数字を4桁にフォーマット
        storeNumber = Format(wsSource.Cells(i, columnH).Value, "0000")
        
        ' I列の店舗名を取得
        storeName = wsSource.Cells(i, columnI).Value
        
        ' 文字列結合
        formattedText = storeNumber & " - " & storeName
        
        ' 結果を挿入（例としてA1セルに挿入）
        Set targetCell = wsTarget.Cells(1, 1) ' 必要に応じて挿入箇所を変更
        targetCell.Value = formattedText
        
        ' 短辺両面印刷の設定
        With wsTarget.PageSetup
            .Orientation = xlPortrait
            .PrintHeadings = False
            .CenterHorizontally = True
            .CenterVertically = True
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .Duplex = xlDuplexSimplex
        End With
        
        ' 印刷
        wsTarget.PrintOut
    Next i
    
    ' 引用元ブックを閉じる
    wbSource.Close SaveChanges:=False
    
    ' 処理終了メッセージ
    MsgBox "処理が完了しました。", vbInformation
End Sub