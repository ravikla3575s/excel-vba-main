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

Sub ProcessRebillSelection()
    Dim wsDetails As Worksheet
    Dim listBox As Object
    Dim selectedData As Object
    Dim i As Long
    Dim rowIndex As Long
    Dim startRowDict As Object
    Dim category As String
    Dim insertRows As Long

    ' 転記用ワークシート取得
    Set wsDetails = ThisWorkbook.Sheets(2) ' 詳細用シート

    ' 選択データを格納
    Set selectedData = CreateObject("Scripting.Dictionary")

    ' リストボックスのデータ取得
    Set listBox = UserForms("返戻再請求の選択").Controls("listBox")

    ' 開始行管理用 Dictionary 作成
    Set startRowDict = CreateObject("Scripting.Dictionary")
    startRowDict.Add "社保返戻再請求", GetStartRow(wsDetails, "社保返戻再請求")
    startRowDict.Add "国保返戻再請求", GetStartRow(wsDetails, "国保返戻再請求")
    startRowDict.Add "社保月遅れ請求", GetStartRow(wsDetails, "社保月遅れ請求")
    startRowDict.Add "国保月遅れ請求", GetStartRow(wsDetails, "国保月遅れ請求")

    ' **返戻再請求の転記**
    insertRows = 0 ' 追加行数カウント
    category = "社保返戻再請求" ' 返戻再請求で始める
    Dim rebillCount As Long: rebillCount = 0 ' 返戻再請求の件数カウント

    ' 選択された項目を取得
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) Then
            selectedData.Add i, listBox.List(i)
            rebillCount = rebillCount + 1
        End If
    Next i

    ' **5行以上ある場合は、行を追加**
    If rebillCount > 4 Then
        insertRows = rebillCount - 4
        wsDetails.Rows(startRowDict(category) + 1 & ":" & startRowDict(category) + insertRows).Insert Shift:=xlDown
    End If

    ' **返戻再請求の転記**
    For Each rowIndex In selectedData.Keys
        wsDetails.Cells(startRowDict(category), 5).Value = selectedData(rowIndex) ' 調剤年月
        wsDetails.Cells(startRowDict(category), 6).Value = selectedData(rowIndex) ' 患者氏名
        wsDetails.Cells(startRowDict(category), 7).Value = selectedData(rowIndex) ' 医療機関名
        wsDetails.Cells(startRowDict(category), 10).Value = selectedData(rowIndex) ' 請求点数

        ' 開始行を +1 して調整
        startRowDict(category) = startRowDict(category) + 1
    Next rowIndex

    ' **追加行数分を月遅れ請求の開始行に足す**
    If insertRows > 0 Then
        startRowDict("社保月遅れ請求") = startRowDict("社保月遅れ請求") + insertRows
        startRowDict("国保月遅れ請求") = startRowDict("国保月遅れ請求") + insertRows
    End If

    ' **月遅れ請求の転記**
    category = "社保月遅れ請求"
    For i = 0 To listBox.ListCount - 1
        If Not selectedData.Exists(i) Then
            wsDetails.Cells(startRowDict(category), 5).Value = listBox.List(i) ' 調剤年月
            wsDetails.Cells(startRowDict(category), 6).Value = listBox.List(i) ' 患者氏名
            wsDetails.Cells(startRowDict(category), 7).Value = listBox.List(i) ' 医療機関名
            wsDetails.Cells(startRowDict(category), 10).Value = listBox.List(i) ' 請求点数

            ' 開始行を +1 して調整
            startRowDict(category) = startRowDict(category) + 1
        End If
    Next i

    ' UserForm を閉じる
    Unload UserForms("返戻再請求の選択")

    MsgBox "転記が完了しました！", vbInformation, "処理完了"
End Sub

Function CreateRebillSelectionForm(listData As Object) As Object
    Dim uf As Object
    Dim listBox As Object
    Dim chkBox As Object
    Dim btnOK As Object
    Dim i As Long
    Dim rowData As Variant

    ' UserForm を作成
    Set uf = CreateObject("Forms.UserForm")
    uf.Caption = "返戻再請求の選択"
    uf.Width = 400
    uf.Height = 500

    ' ListBox を追加
    Set listBox = uf.Controls.Add("Forms.ListBox.1", "listBox", True)
    listBox.Left = 20
    listBox.Top = 20
    listBox.Width = 350
    listBox.Height = 350
    listBox.MultiSelect = 1

    ' リストデータ追加
    For i = 0 To listData.Count - 1
        rowData = listData.Items()(i)
        listBox.AddItem rowData(0) & " | " & rowData(1) & " | " & rowData(2) & " | " & rowData(3)
    Next i

    ' OKボタンを追加
    Set btnOK = uf.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    btnOK.Caption = "確定"
    btnOK.Left = 150
    btnOK.Top = 400
    btnOK.Width = 100
    btnOK.Height = 30

    ' OKボタンの処理
    btnOK.OnAction = "ProcessRebillSelection"

    ' UserForm を返す
    Set CreateRebillSelectionForm = uf
End Function

Sub ShowRebillSelectionForm(newBook As Workbook)
    Dim wsBilling As Worksheet
    Dim lastRow As Long, i As Long
    Dim userForm As Object
    Dim listData As Object
    Dim rowData As Variant
    
    ' メインシート取得
    Set wsBilling = newBook.Sheets(1)
    lastRow = wsBilling.Cells(Rows.Count, "D").End(xlUp).Row

    ' Dictionary でリストを管理
    Set listData = CreateObject("Scripting.Dictionary")

    ' 現在の請求月取得
    Dim currentBillingMonth As String
    currentBillingMonth = wsBilling.Cells(2, 2).Value ' GYYMM

    ' 該当調剤月以外のデータをリスト化
    For i = 2 To lastRow
        If wsBilling.Cells(i, 2).Value <> currentBillingMonth Then
            rowData = Array(wsBilling.Cells(i, 2).Value, wsBilling.Cells(i, 4).Value, wsBilling.Cells(i, 5).Value, wsBilling.Cells(i, 6).Value)
            listData.Add i, rowData
        End If
    Next i

    ' リストにデータがあればフォーム表示
    If listData.Count > 0 Then
        Set userForm = CreateRebillSelectionForm(listData)
        userForm.Show
    Else
        MsgBox "該当するデータはありません。", vbInformation, "確認"
    End If
End Sub