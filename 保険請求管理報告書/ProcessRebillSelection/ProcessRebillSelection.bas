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
