Sub ProcessMultipleCSVFiles()
    Dim ws As Worksheet
    Dim csvSheet As Worksheet
    Dim lastRow As Long
    Dim folderPath As String
    Dim csvFile As String
    Dim wb As Workbook
    
    ' シート1（転記先のシート）
    Set ws = ThisWorkbook.Sheets(1)
    
    folderPath = ThisWorkbook.Path & Application.PathSeparator
    
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
    ' 請求確定表かどうかを判定する（Cells(1,7)が'請求確定表'）
    IsBillingConfirmation = (csvSheet.Cells(1, 7).Value Like "*請求確定表*")
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
    For i = 5 To 16 ' A5からA16
        searchMonth = ws.Cells(i, 1).Value ' A列の各月のラベルを取得
        
        found = False ' 初期状態では見つかっていない
        
        ' CSVシートの該当するデータを検索
        csvMonth = Replace(csvSheet.Cells(1, 5).Value, "'", "") ' 'を削除
        csvMonth = Replace(csvMonth, " ", "") ' 半角スペースを削除
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
    Next i
    
    ' 該当する月が見つからなかった場合のエラーメッセージ
    If Not found Then
        MsgBox "対象年月日が見つかりません: " & csvMonth, vbExclamation, "エラー"
    End If
End Sub

' ===== 振込額明細書の処理 =====
Sub ProcessPaymentDetails(csvSheet As Worksheet, ws As Worksheet)
    Dim i As Long
    Dim paymentAmount As Double
    Dim totalAmount As Double
    Dim paymentAgencyCode As String
    Dim depositColumn As Integer
    Dim returnManagementSheet As Worksheet
    Dim nextEmptyRow As Long
    Dim diagnosisDate As String
    Dim requestPoints As Double
    Dim finalPoints As Double
    Dim difference As Double
    Dim status As String
    
    ' "返戻管理"シートを取得
    Set returnManagementSheet = ThisWorkbook.Sheets("返戻管理")
    
    ' 次にデータを転記する行を取得（返戻管理シートの最終行+1）
    nextEmptyRow = returnManagementSheet.Cells(returnManagementSheet.Rows.Count, "A").End(xlUp).Row + 1
    
    ' 支払機関コードを取得（例：7桁目の値を使用）
    paymentAgencyCode = Mid(csvSheet.Name, 7, 1) ' ファイル名から7桁目を取得
    
    ' 支払機関コードに応じて転記先の列を設定
    Select Case paymentAgencyCode
        Case "1"
            depositColumn = 5 ' 支払機関コードが1の場合、5列目に転記
        Case "2"
            depositColumn = 6 ' 支払機関コードが2の場合、6列目に転記
        Case "3"
            depositColumn = 8 ' 支払機関コードが3の場合、8列目に転記
        Case Else
            MsgBox "不明な支払機関コードです: " & paymentAgencyCode, vbExclamation, "エラー"
            Exit Sub
    End Select
    
    ' 診療年月を取得（例: Cells(1,2)の5桁の値）
    diagnosisDate = csvSheet.Cells(1, 2).Value
    
    ' 82列目の合計額を計算し、空欄を無視
    totalAmount = 0
    For i = 3 To csvSheet.Cells(csvSheet.Rows.Count, "A").End(xlUp).Row
        If IsNumeric(csvSheet.Cells(i, 82).Value) Then
            totalAmount = totalAmount + csvSheet.Cells(i, 82).Value
        Else
            ' 82列目が空欄の場合、返戻管理シートに転記
            returnManagementSheet.Cells(nextEmptyRow, 1).Value = paymentAgencyCode ' 支払機関
            returnManagementSheet.Cells(nextEmptyRow, 2).Value = diagnosisDate ' 診療年月
            returnManagementSheet.Cells(nextEmptyRow, 3).Value = csvSheet.Cells(i, 14).Value ' 患者名
            returnManagementSheet.Cells(nextEmptyRow, 4).Value = "振込なし" ' 返戻日
            returnManagementSheet.Cells(nextEmptyRow, 5).Value = csvSheet.Cells(i, 22).Value ' 請求時点数
            returnManagementSheet.Cells(nextEmptyRow, 6).Value = csvSheet.Cells(i, 23).Value ' 返戻後再請求時点数
            returnManagementSheet.Cells(nextEmptyRow, 7).Value = 0 ' 振込額（なし）
            returnManagementSheet.Cells(nextEmptyRow, 8).Value = csvSheet.Cells(i, 22).Value ' 差額（請求時点数）
            returnManagementSheet.Cells(nextEmptyRow, 9).Value = "返戻" ' 請求状況
            nextEmptyRow = nextEmptyRow + 1
        End If
        
        ' 22列目と23列目の差異をチェック
        If IsNumeric(csvSheet.Cells(i, 22).Value) And IsNumeric(csvSheet.Cells(i, 23).Value) Then
            requestPoints = csvSheet.Cells(i, 22).Value
            finalPoints = csvSheet.Cells(i, 23).Value
            difference = requestPoints - finalPoints
            
            If difference <> 0 Then
                ' 差異がある場合、返戻管理シートに転記
                returnManagementSheet.Cells(nextEmptyRow, 1).Value = paymentAgencyCode ' 支払機関
                returnManagementSheet.Cells(nextEmptyRow, 2).Value = diagnosisDate ' 診療年月
                returnManagementSheet.Cells(nextEmptyRow, 3).Value = csvSheet.Cells(i, 14).Value ' 患者名
                returnManagementSheet.Cells(nextEmptyRow, 4).Value = Now ' 返戻日
                returnManagementSheet.Cells(nextEmptyRow, 5).Value = requestPoints ' 請求時点数
                returnManagementSheet.Cells(nextEmptyRow, 6).Value = finalPoints ' 返戻後再請求時点数
                returnManagementSheet.Cells(nextEmptyRow, 7).Value = csvSheet.Cells(i, 82).Value ' 振込額
                returnManagementSheet.Cells(nextEmptyRow, 8).Value = difference ' 請求振込差額
                returnManagementSheet.Cells(nextEmptyRow, 9).Value = "差異あり" ' 請求状況
                nextEmptyRow = nextEmptyRow + 1
            End If
        End If
    Next i
    
    ' 合計額を指定されたセルに転記
    ws.Cells(15, depositColumn).Value = totalAmount
End Sub

' ===== 調剤報酬明細書の処理 =====
Sub ProcessDispensingFeeStatement(csvSheet As Worksheet, ws As Worksheet)
    Dim referenceAmount As Double
    Dim searchMonth As String
    Dim i As Long, j As Long
    Dim csvMonth As String
    Dim found As Boolean
    
    ' シート1のA5からA16の各月のラベルを検索
    For i = 5 To 16 ' A5からA16
        searchMonth = ws.Cells(i, 1).Value ' A列の各月のラベルを取得
        
        found = False ' 初期状態では見つかっていない
        
        ' CSVシートの該当するデータを検索
        csvMonth = Replace(csvSheet.Cells(1, 5).Value, "'", "") ' 'を削除
        csvMonth = ConvertZenkakuToHankaku(csvMonth) ' 全角数字を半角に変換
        csvMonth = Format(Format(csvMonth, "@@@@/@@/@@"), "ggge年m月処理分") ' 令和◯年◯月処理分
        
        ' 一致するか比較
        If csvMonth = searchMonth Then
            ' 一致した場合にデータを転記
            found = True ' 見つかったことを示す
            
            ' 調剤報酬明細書は1行目の33列目の振込参考金額を取得
            referenceAmount = csvSheet.Cells(1, 33).Value
            
            ' 振込参考金額をシート1の指定されたセルに転記（例: B25に転記する）
            ws.Cells(i, 2).Value = referenceAmount
            
            Exit For ' データを転記したらループを抜ける
        End If
    Next i
    
    ' 該当する月が見つからなかった場合のエラーメッセージ
    If Not found Then
        MsgBox "対象年月日が見つかりません: " & csvMonth, vbExclamation, "エラー"
    End If
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