Sub ProcessMultipleCSVFiles()
    Dim ws As Worksheet
    Dim csvSheet As Worksheet
    Dim folderPath As String
    Dim csvFile As String
    Dim wb As Workbook
    Dim errorFiles As String ' エラーが発生したファイル名を格納する変数
    Dim processedFiles As Collection ' 処理済みファイル名を格納するコレクション

    ' エラーファイルリストの初期化
    errorFiles = ""
    Set processedFiles = New Collection ' 処理済みファイルリストの初期化

    ' シート1（転記先のシート）
    Set ws = ThisWorkbook.Sheets(1)
    folderPath = ThisWorkbook.Path & Application.PathSeparator

    ' フォルダ内のすべてのCSVファイルを処理
    csvFile = Dir(folderPath & "*.csv")

    Do While csvFile <> ""
        ' すでに処理済みのファイルかを確認
        On Error Resume Next
        processedFiles.Add csvFile, csvFile ' ファイル名をキーに追加
        If Err.Number <> 0 Then
            ' ファイルが既にリストにある場合はスキップ
            Err.Clear
            csvFile = Dir
            On Error GoTo 0
            GoTo ContinueLoop
        End If
        On Error GoTo 0

        ' CSVファイルを開く
        On Error Resume Next ' エラーを一時的に無視して次に進む
        Set wb = Workbooks.Open(folderPath & csvFile)
        If Err.Number <> 0 Then
            ' ファイルを開けなかった場合、エラーファイルリストに追加
            errorFiles = errorFiles & vbCrLf & csvFile
            Err.Clear
            csvFile = Dir
            On Error GoTo 0
            GoTo ContinueLoop
        End If
        On Error GoTo 0

        Set csvSheet = wb.Sheets(1) ' CSVシートを取得

        ' ファイル形式を判定して対応する処理を実行
        If IsBillingConfirmation(csvSheet) Then
            On Error Resume Next
            ProcessBillingConfirmation csvSheet, ws
            If Err.Number <> 0 Then
                ' 処理中にエラーが発生した場合、エラーファイルリストに追加
                errorFiles = errorFiles & vbCrLf & csvFile
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf IsPaymentDetails(csvFile) Then
            On Error Resume Next
            ProcessPaymentDetails csvSheet, ws
            If Err.Number <> 0 Then
                errorFiles = errorFiles & vbCrLf & csvFile
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf IsDispensingFeeStatement(csvSheet) Then
            On Error Resume Next
            ProcessDispensingFeeStatement csvSheet, ws
            If Err.Number <> 0 Then
                errorFiles = errorFiles & vbCrLf & csvFile
                Err.Clear
            End If
            On Error GoTo 0
        Else
            ' ファイル形式が不明な場合もエラーファイルリストに追加
            errorFiles = errorFiles & vbCrLf & csvFile
        End If

        ' CSVファイルを閉じる
        wb.Close False

ContinueLoop:
        ' 次のCSVファイルを取得
        csvFile = Dir
    Loop

    ' エラーが発生したファイルがあればメッセージボックスで通知
    If Len(errorFiles) > 0 Then
        MsgBox "以下のファイルで処理中にエラーが発生しました:" & vbCrLf & errorFiles, vbExclamation, "エラー一覧"
    Else
        MsgBox "すべてのCSVファイルの処理が完了しました。"
    End If
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

' ===== 各形式に応じた処理 =====
Sub ProcessBillingConfirmation(csvSheet As Worksheet, ws As Worksheet)
    Dim lastRow As Long
    Dim searchMonth As String
    Dim i As Long
    Dim csvMonth As String
    Dim normalStartCol As Integer
    Dim reClaimStartCol As Integer
    Dim found As Boolean

    ' CSVシートの最終行を取得
    lastRow = csvSheet.Cells(csvSheet.Rows.Count, "A").End(xlUp).Row

    ' シート1のA5からA16の各月のラベルを検索
    For i = 5 To 16
        searchMonth = ws.Cells(i, 1).Value

        found = False ' 初期状態では見つかっていない

        ' CSVシートの該当するデータを検索
        csvMonth = Replace(csvSheet.Cells(1, 5).Value, "'", "") ' 'を削除
        csvMonth = Replace(csvMonth, " ", "") ' 半角スペースを削除
        csvMonth = ConvertZenkakuToHankaku(csvMonth) ' 全角数字を半角に変換

        ' 一致するか比較
        If csvMonth = searchMonth Then
            ' 一致した場合にデータを転記
            found = True

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
    Dim storeCode As String
    Dim processDate As String
    Dim uniqueKey As String
    Dim status As String

    ' "返戻管理"シートを取得
    Set returnManagementSheet = ThisWorkbook.Sheets("返戻管理")

    ' 次にデータを転記する行を取得（返戻管理シートの最終行+1）
    nextEmptyRow = returnManagementSheet.Cells(returnManagementSheet.Rows.Count, "A").End(xlUp).Row + 1

    ' 支払機関コードを取得（例：7桁目の値を使用）
    paymentAgencyCode = Mid(csvSheet.Name, 7, 1) ' ファイル名から7桁目を取得

    ' 診療年月を取得（例: Cells(1,2)の5桁の値）
    diagnosisDate = csvSheet.Cells(1, 2).Value

    ' 店番（仮に4桁の店番号を指定）
    storeCode = ThisWorkbook.Worksheets(1).Cells(3,2) ' 必要に応じて適切な値を取得または設定

    ' 返戻処理年月
    processDate = 5 & Format(Date, "yymm")

    ' 82列目の合計額を計算し、空欄を無視
    totalAmount = 0
    For i = 3 To csvSheet.Cells(csvSheet.Rows.Count, "A").End(xlUp).Row
        If IsNumeric(csvSheet.Cells(i, 82).Value) Then
            totalAmount = totalAmount + csvSheet.Cells(i, 82).Value
        Else
            ' 82列目が空欄の場合、返戻管理シートに転記（種別コード1）
            uniqueKey = paymentAgencyCode & diagnosisDate & "1" & processDate & storeCode
            returnManagementSheet.Cells(nextEmptyRow, 1).Value = uniqueKey
            returnManagementSheet.Cells(nextEmptyRow, 2).Value = paymentAgencyCode
            returnManagementSheet.Cells(nextEmptyRow, 3).Value = diagnosisDate
            returnManagementSheet.Cells(nextEmptyRow, 4).Value = csvSheet.Cells(i, 14).Value
            returnManagementSheet.Cells(nextEmptyRow, 5).Value = "振込なし"
            returnManagementSheet.Cells(nextEmptyRow, 6).Value = csvSheet.Cells(i, 22).Value
            returnManagementSheet.Cells(nextEmptyRow, 7).Value = csvSheet.Cells(i, 23).Value
            returnManagementSheet.Cells(nextEmptyRow, 8).Value = 0
            returnManagementSheet.Cells(nextEmptyRow, 9).Value = csvSheet.Cells(i, 22).Value
            returnManagementSheet.Cells(nextEmptyRow, 10).Value = "返戻"
            nextEmptyRow = nextEmptyRow + 1
        End If

        ' 22列目と23列目の差異をチェック
        If IsNumeric(csvSheet.Cells(i, 22).Value) And IsNumeric(csvSheet.Cells(i, 23).Value) Then
            requestPoints = csvSheet.Cells(i, 22).Value
            finalPoints = csvSheet.Cells(i, 23).Value
            difference = requestPoints - finalPoints

            If difference <> 0 Then
                If difference > 0 Then
                    uniqueKey = paymentAgencyCode & diagnosisDate & "2" & processDate & storeCode ' 加点の場合
                Else
                    uniqueKey = paymentAgencyCode & diagnosisDate & "3" & processDate & storeCode ' 減点の場合
                End If
                returnManagementSheet.Cells(nextEmptyRow, 1).Value = uniqueKey
                returnManagementSheet.Cells(nextEmptyRow, 2).Value = paymentAgencyCode
                returnManagementSheet.Cells(nextEmptyRow, 3).Value = diagnosisDate
                returnManagementSheet.Cells(nextEmptyRow, 4).Value = csvSheet.Cells(i, 14).Value
                returnManagementSheet.Cells(nextEmptyRow, 5).Value = Now
                returnManagementSheet.Cells(nextEmptyRow, 6).Value = requestPoints
                returnManagementSheet.Cells(nextEmptyRow, 7).Value = finalPoints
                returnManagementSheet.Cells(nextEmptyRow, 8).Value = csvSheet.Cells(i, 82).Value
                returnManagementSheet.Cells(nextEmptyRow, 9).Value = difference
                returnManagementSheet.Cells(nextEmptyRow, 10).Value = "差異あり"
                nextEmptyRow = nextEmptyRow + 1
            End If
        End If
    Next i

    ' 合計額を指定されたセルに転記
    ws.Cells(15, depositColumn).Value = totalAmount
End Sub

Sub ProcessDispensingFeeStatement(csvSheet As Worksheet, ws As Worksheet)
    Dim referenceAmount As Double
    Dim searchMonth As String
    Dim i As Long
    Dim csvMonth As String
    Dim found As Boolean

    ' シート1のA5からA16の各月のラベルを検索
    For i = 5 To 16
        searchMonth = ws.Cells(i, 1).Value

        found = False

        ' CSVシートの該当するデータを検索
        csvMonth = Replace(csvSheet.Cells(1, 5).Value, "'", "")
        csvMonth = ConvertZenkakuToHankaku(csvMonth)
        csvMonth = Format(Format(csvMonth, "@@@@/@@/@@"), "ggge年m月処理分")

        ' 一致するか比較
        If csvMonth = searchMonth Then
            found = True
            referenceAmount = csvSheet.Cells(1, 33).Value
            ws.Cells(i, 2).Value = referenceAmount
            Exit For
        End If
    Next i

    ' 該当する月が見つからなかった場合のエラーメッセージ
    If Not found Then
        MsgBox "対象年月日が見つかりません: " & csvMonth, vbExclamation, "エラー"
    End If
End Sub

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
            Case Else: result = result & currentChar
        End Select
    Next i

    ConvertZenkakuToHankaku = result
End Function

