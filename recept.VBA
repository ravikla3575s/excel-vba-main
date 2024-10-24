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
        
        ' ファイル形式を判定して対応する処理を実行
        If IsBillingConfirmation(csvSheet) Then
            ProcessBillingConfirmation csvSheet, ws
        ElseIf IsPaymentDetails(csvFile) Then
            ProcessPaymentDetails csvSheet, ws
        ElseIf IsDispensingFeeStatement(csvSheet) Then
            ProcessDispensingFeeStatement csvSheet, ws
        Else
            MsgBox "不明なCSV形式: " & csvFile, vbExclamation, "エラー"
        End If
        
        ' CSVファイルを閉じる
        wb.Close False ' 保存せずに閉じる
        
        ' 次のCSVファイルを取得
        csvFile = Dir
    Loop
    
    MsgBox "すべてのCSVファイルの処理が完了しました。"
End Sub

' ===== 判定関数 =====
Function IsBillingConfirmation(csvSheet As Worksheet) As Boolean
    ' 請求確定表かどうかを判定する（Cells(1,6)が'請求確定表'、年月日が'令和X年X月処理分'）
    IsBillingConfirmation = (csvSheet.Cells(1, 6).Value Like "'請求確定表'") 
End Function

Function IsPaymentDetails(csvFileName As String) As Boolean
    ' 振込額明細書かどうかを判定する（ファイル名がRTfmeiで始まる）
    IsPaymentDetails = Left(csvFileName, 6) = "RTfmei"
End Function

Function IsDispensingFeeStatement(csvSheet As Worksheet) As Boolean
    ' 調剤報酬明細書かどうかを判定する（Cells(1,1)がH、Cells(2,1)がR2）
    IsDispensingFeeStatement = (csvSheet.Cells(1, 1).Value = "H") And (csvSheet.Cells(2, 1).Value = "R2")
End Function

' ===== 請求確定表の処理 =====
Sub ProcessBillingConfirmation(csvSheet As Worksheet, ws As Worksheet)
    Dim lastRow As Long
    Dim searchMonth As String
    Dim i As Long, j As Long
    Dim csvMonth As String
    Dim normalStartCol As Integer
    Dim reClaimStartCol As Integer
    Dim found As Boolean
    
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
            ws.Cells(i, reClaimStartCol).Resize(1, 7).Value = WorksheetFunction.Transpose(csvSheet.Cells(12, 11).Resize(7, 1).Value)
            
            Exit For ' データを転記したらループを抜ける
        End If
        
        ' 該当する月が見つからなかった場合のエラーメッセージ
        If Not found Then
            MsgBox "対象年月日が見つかりません: " & searchMonth, vbExclamation, "エラー"
        End If
    Next i
End Sub

' ===== 振込額明細書の処理 =====
Sub ProcessPaymentDetails(csvSheet As Worksheet, ws As Worksheet)
    Dim i As Long
    Dim paymentAmount As Double
    
    ' 振込額明細書は82行目の振込額を取得
    paymentAmount = csvSheet.Cells(82, 1).Value
    
    ' 振込額をシート1の指定されたセルに転記（例: B20に転記する）
    ws.Cells(20, 2).Value = paymentAmount
End Sub

' ===== 調剤報酬明細書の処理 =====
Sub ProcessDispensingFeeStatement(csvSheet As Worksheet, ws As Worksheet)
    Dim referenceAmount As Double
    
    ' 調剤報酬明細書は1行目の33列目の振込参考金額を取得
    referenceAmount = csvSheet.Cells(1, 33).Value
    
    ' 振込参考金額をシート1の指定されたセルに転記（例: B25に転記する）
    ws.Cells(25, 2).Value = referenceAmount
End Sub

' ===== 全角数字を半角数字に変換する関数 =====
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