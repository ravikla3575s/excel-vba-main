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