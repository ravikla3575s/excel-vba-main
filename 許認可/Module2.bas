Sub UpdateMultiplePharmacists()
    Dim i As Integer
    Dim PharmacistName As String, strName As String
    Dim updWS As Worksheet
    Dim Kws As Worksheet
    Dim updateRow As Integer
    Dim startColumn As Integer
    
    Set updWS = Worksheets("届出一覧テーブル")
    Set Kws = Worksheets("所属変更")
    
    strName = Kws.Cells(2, 1).value
    
    ' 基準カラムを指定
    For startColumn = 1 To 215
    If updWS.Cells(1, startColumn) = "常勤薬剤師1" Then
        Exit For
    End If
    Next startColumn
    
    ' B13:B17 の範囲を順に処理
    For i = 13 To 17
        PharmacistName = Kws.Cells(i, 2).value
        
        If PharmacistName = "" Or "0" Then
            Exit For
        End If
        
        ' 更新する行を探す
        updateRow = findUpdateRow(updWS, strName)
        
        ' 薬剤師の登録の有無を確認
        If IsPharmacistRegistered(updWS, PharmacistName, updateRow, startColumn) Then
            MsgBox PharmacistName & "は既に登録されています。"
        Else
            ' 更新する列を探して、値を更新
            Call UpdatePharmacistAssignmentWithParams(updWS, PharmacistName, updateRow)
        End If
    Next i
End Sub

Function IsPharmacistRegistered(updWS As Worksheet, PharmacistName As String, updateRow As Integer, startColumn As Integer) As Boolean
    Dim i As Integer
    
    ' 基準カラムから右に向かって20個分のセルをチェック
    On Error GoTo ErrLabel
    For i = 0 To 19
        If updWS.Cells(updateRow, startColumn + i).value = PharmacistName Then
            IsPharmacistRegistered = True
            Exit Function
        End If
    Next i
    
    ' 見つからなかった場合、Falseを返す
    IsPharmacistRegistered = False
    Exit Function
      
ErrLabel:
    msg = "エラーが発生しました"
        IsPharmacistRegistered = False
    Resume Next
End Function

Sub UpdatePharmacistAssignmentWithParams(updWS As Worksheet, PharmacistName As String, updateRow As Integer)
    Dim updateColumn As Integer
    
    ' 更新する列を探す
    updateColumn = findUpdateColumn(updateRow)
    
    If updateColumn > 0 Then
        updWS.Cells(updateRow, updateColumn).value = PharmacistName
    End If
End Sub

Function findUpdateRow(updWS As Worksheet, strName As String) As Integer
    Dim updateRow As Integer
    
    ' 指定された店舗名がある行を検索
    For updateRow = 2 To 70
        If updWS.Cells(updateRow, 2).value = strName Then
            findUpdateRow = updateRow
            Exit Function
        End If
    Next updateRow
    
    ' 見つからなかった場合、0を返す
    findUpdateRow = 0
End Function

Function findUpdateColumn(updateRow As Integer) As Integer
    Dim updateColumn As Integer
    Dim updWS As Worksheet
    Dim itemNames As Variant
    Dim i As Integer
    
    Set updWS = Worksheets("届出一覧テーブル")
    
    ' 非常勤薬剤師のスロットを定義
    itemNames = Array("非常勤薬剤師6", "非常勤薬剤師7", "非常勤薬剤師8", "非常勤薬剤師9", "非常勤薬剤師10")
    
    ' itemNames 配列を順に処理し、空欄の列を探す
    For i = LBound(itemNames) To UBound(itemNames)
        For updateColumn = 121 To 188
            If updWS.Cells(1, updateColumn).value = itemNames(i) Then
                ' 該当するセルが空欄かをチェック
                If IsEmpty(updWS.Cells(updateRow, updateColumn)) Then
                    findUpdateColumn = updateColumn
                    Exit Function
                End If
                Exit For
            End If
        Next updateColumn
    Next i
    
    ' 見つからなかった場合、0を返す
    findUpdateColumn = 0
End Function
