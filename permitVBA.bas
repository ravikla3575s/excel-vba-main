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
        For i = 5 To 16 ' A5からA16
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
        Next i
        
        ' CSVファイルを閉じる
        wb.Close False ' 保存せずに閉じる
        ' 該当する月が見つからなかった場合のエラーメッセージ
        If Not found Then
            MsgBox "対象年月日が見つかりません: "
        End If
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