Sub TransferBillingDetails(newBook As Workbook)
    Dim wsBilling As Worksheet, wsDetails As Worksheet
    Dim lastRowBilling As Long, lastRowDetails As Long
    Dim i As Long, j As Long
    Dim dispensingMonth As String, convertedMonth As String
    Dim payerCode As String, payerType As String
    Dim receiptNo As String, claimPoints As Double
    Dim expectedPayment As Double, unpaidReceipts As Double
    Dim startRowDict As Object
    Dim category As String
    Dim startRow As Long
    
    ' シート設定
    Set wsBilling = newBook.Sheets(1) ' メインシート
    Set wsDetails = newBook.Sheets(2) ' 詳細用シート

    ' 最終行取得
    lastRowBilling = wsBilling.Cells(Rows.Count, "D").End(xlUp).Row
    lastRowDetails = wsDetails.Cells(Rows.Count, "D").End(xlUp).Row

    ' CSVデータの請求先分類
    payerCode = Trim(wsBilling.Cells(2, 3).Value) ' 空白除去
    Select Case payerCode
        Case "1": payerType = "社保"
        Case "2": payerType = "国保"
        Case Else: payerType = "労災"
    End Select

    ' 開始行管理用 Dictionary 作成
    Set startRowDict = CreateObject("Scripting.Dictionary")
    startRowDict.Add "社保返戻再請求", GetStartRow(wsDetails, "社保返戻再請求")
    startRowDict.Add "国保返戻再請求", GetStartRow(wsDetails, "国保返戻再請求")
    startRowDict.Add "社保月遅れ請求", GetStartRow(wsDetails, "社保月遅れ請求")
    startRowDict.Add "国保月遅れ請求", GetStartRow(wsDetails, "国保月遅れ請求")
    startRowDict.Add "労災", lastRowDetails + 1 ' 労災は常に最終行の次

    ' 請求タイプに応じて転記開始行を取得
    If payerType = "社保" Then
        If InStr(LCase(wsBilling.Cells(2, 4).Value), "返戻") > 0 Then
            category = "社保返戻再請求"
        Else
            category = "社保月遅れ請求"
        End If
    ElseIf payerType = "国保" Then
        If InStr(LCase(wsBilling.Cells(2, 4).Value), "返戻") > 0 Then
            category = "国保返戻再請求"
        Else
            category = "国保月遅れ請求"
        End If
    Else
        category = "労災"
    End If

    startRow = startRowDict(category) ' 選択されたカテゴリの開始行

    ' 転記処理
    j = startRow
    For i = 2 To lastRowBilling
        dispensingMonth = wsBilling.Cells(i, 2).Value ' GYYMM形式
        receiptNo = wsBilling.Cells(i, 1).Value ' レセプト番号
        claimPoints = wsBilling.Cells(i, 6).Value ' 請求点数

        ' 調剤年月を YY.MM 形式に変換
        convertedMonth = ConvertToWesternDate(dispensingMonth)

        ' 転記
        wsDetails.Cells(j, 4).Value = wsBilling.Cells(i, 4).Value ' 患者氏名
        wsDetails.Cells(j, 5).Value = convertedMonth ' 調剤年月
        wsDetails.Cells(j, 6).Value = wsBilling.Cells(i, 5).Value ' 処方元医療機関名
        wsDetails.Cells(j, 8).Value = payerType ' 請求先
        wsDetails.Cells(j, 10).Value = claimPoints ' 請求点数

        ' **転記後、次の行に移行 & 他の開始行と重なる場合 +1 する**
        j = j + 1
        If IsStartRowOverlap(startRowDict, j) Then
            IncreaseAllStartRows startRowDict
        End If
    Next i

    MsgBox "請求詳細データの転記が完了しました！", vbInformation, "処理完了"
End Sub

Function GetStartRow(ws As Worksheet, category As String) As Long
    Dim lastRow As Long
    Dim i As Long
    Dim keyword As String
    
    ' 最終行取得
    lastRow = ws.Cells(Rows.Count, "H").End(xlUp).Row

    ' キーワードに応じた開始行を決定
    Select Case category
        Case "社保返戻再請求"
            keyword = "国家→医本"
        Case "国保返戻再請求"
            keyword = "⑨返戻分再請求分（医保）"
        Case "社保月遅れ請求"
            keyword = "⑨返戻分再請求分"
        Case "国保月遅れ請求"
            keyword = "⑩月遅れ請求分（医保）"
        Case Else
            GetStartRow = lastRow + 1
            Exit Function
    End Select

    ' H列を検索（2行目から）
    For i = 2 To lastRow
        If ws.Cells(i, 8).Value = keyword Then
            GetStartRow = i + 1
            Exit Function
        End If
    Next i

    ' 該当なしの場合
    GetStartRow = lastRow + 1
End Function

Function IsStartRowOverlap(startRowDict As Object, newRow As Long) As Boolean
    Dim key As Variant
    For Each key In startRowDict.Keys
        If startRowDict(key) = newRow Then
            IsStartRowOverlap = True
            Exit Function
        End If
    Next key
    IsStartRowOverlap = False
End Function

Sub IncreaseAllStartRows(startRowDict As Object)
    Dim key As Variant
    For Each key In startRowDict.Keys
        startRowDict(key) = startRowDict(key) + 1
    Next key
End Sub