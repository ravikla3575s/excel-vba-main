Sub TransferBillingDetails(newBook As Workbook)
    Dim wsBilling As Worksheet, wsDetails As Worksheet
    Dim lastRowBilling As Long, lastRowDetails As Long
    Dim i As Long, j As Long
    Dim dispensingMonth As String, convertedMonth As String
    Dim payerCode As String, payerType As String
    Dim receiptNo As String, claimPoints As Double, decisionPoints As Double
    Dim expectedPayment As Double, unpaidReceipts As Double
    Dim startRow As Long: startRow = 19 ' 転記開始行
    
    ' シート設定
    Set wsBilling = newBook.Sheets(1) ' メインシート
    Set wsDetails = newBook.Sheets(2) ' シート2: 詳細情報用

    ' 最終行取得
    lastRowBilling = wsBilling.Cells(Rows.Count, "D").End(xlUp).Row
    lastRowDetails = wsDetails.Cells(Rows.Count, "D").End(xlUp).Row

    ' 転記処理
    j = startRow ' 転記開始行
    For i = 2 To lastRowBilling
        ' レセプト番号と調剤年月取得
        dispensingMonth = wsBilling.Cells(i, 2).Value ' GYYMM形式
        receiptNo = wsBilling.Cells(i, 1).Value ' レセプト番号
        claimPoints = wsBilling.Cells(i, 6).Value ' 請求点数
        decisionPoints = wsBilling.Cells(i, 7).Value ' 決定点数
        expectedPayment = wsBilling.Cells(i, 9).Value ' 振込予定額
        unpaidReceipts = wsBilling.Cells(i, 10).Value ' 未請求レセプト金額

        ' 調剤年月を YY.MM 形式の西暦へ変換
        convertedMonth = ConvertToWesternDate(dispensingMonth)

        ' 支払期間番号取得
        payerCode = wsBilling.Cells(i, 3).Value ' 支払機関番号
        payerType = GetPayerType(payerCode)

        ' 転記処理
        wsDetails.Cells(j, 4).Value = wsBilling.Cells(i, 4).Value ' 患者氏名
        wsDetails.Cells(j, 5).Value = convertedMonth ' 調剤年月（西暦.月）
        wsDetails.Cells(j, 6).Value = wsBilling.Cells(i, 5).Value ' 処方元医療機関名
        wsDetails.Cells(j, 8).Value = payerType ' 請求先
        wsDetails.Cells(j, 10).Value = claimPoints ' 請求点数
        wsDetails.Cells(j, 11).Value = decisionPoints ' 決定点数
        wsDetails.Cells(j, 12).Value = expectedPayment ' 振込予定額
        wsDetails.Cells(j, 13).Value = unpaidReceipts ' 未請求レセプト

        ' 次の行へ
        j = j + 1
    Next i

    ' 処理完了メッセージ
    MsgBox "請求詳細データの転記が完了しました！", vbInformation, "処理完了"
End Sub