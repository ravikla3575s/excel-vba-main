Sub 厚生局異動届作成()
    Dim chkRow As Integer, KRow As Integer
    Dim upDate As Date
    Dim orgValue As String, chgValue As String
    Dim orgHours As Double, chgHours As Double
    Dim ws As Worksheet, Kws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("検索")
    Set Kws = ThisWorkbook.Worksheets("所属変更")
    
    upDate = ws.Cells(2,1)

    ' 所属変更シートのデータをクリア
    Kws.Range("B3:E11").ClearContents
    
    If ws.Cells(2, 2) <> Kws.Cells(2, 1) Then
        Debug.Print ws.Cells(2, 2).value
        Kws.Cells(2, 1).value = ws.Cells(2, 2).value
    End If
    
    ' 異動対象の確認と転記処理
    For chkRow = 3 To 12
        orgValue = ws.Cells(chkRow, 5).Value  ' 変更前氏名
        chgValue = ws.Cells(chkRow, 9).Value  ' 変更後氏名
        orgHours = ws.Cells(chkRow, 6).Value  ' 変更前勤務時間
        chgHours = ws.Cells(chkRow, 10).Value ' 変更後勤務時間
        KRow = Kws.Cells(7, 2).End(xlUp).Row + 1
    ' 重複チェック用フラグ初期化
    isDuplicate = False

        ' 重複チェック: B列の既存データから一致する名前を探す
        For existingRow = 3 To 11
            If Kws.Cells(existingRow, 2).Value = chgValue And chgValue <> "" Then
                isDuplicate = True
                Exit For
            End If
        Next existingRow
        
        ' 重複が見つかった場合は次のデータへ
        If isDuplicate Then
            Debug.Print "重複データ: " & chgValue & " (行 " & chkRow & ")"
            GoTo NextRow
        End If        ' 重複チェック用フラグ初期化
        isDuplicate = False

        ' 重複チェック: B列の既存データから一致する名前を探す
        For existingRow = 3 To 11
            If Kws.Cells(existingRow, 2).Value = chgValue And chgValue <> "" Then
                isDuplicate = True
                Exit For
            End If
        Next existingRow
        
        ' 重複が見つかった場合は次のデータへ
        If isDuplicate Then
            Debug.Print "重複データ: " & chgValue & " (行 " & chkRow & ")"
            GoTo NextRow
        End If
        
        ' 人員が増える場合の処理
        If orgValue <> chgValue And orgValue = "" Then
            Kws.Cells(KRow, 2).value = ws.Cells(chkRow, 9).value
            If Kws.Cells(21, 2).value <> "" Then '新人薬剤師の場合
                Kws.Cells(KRow, 3).value = Kws.Cells(21, 2).value
                Kws.Cells(KRow, 5).value = "常勤"
            Else
                Kws.Cells(KRow, 3).value = upDate
                If chgHours < 32 then
                    Kws.Cells(KRow, 5).value = "非常勤"
                Else 
                    Kws.Cells(KRow,5).value = "常勤"
                End if
            End If
        ' 人員が減る場合の処理
        ElseIf orgValue <> chgValue And chgValue = "" Then
            If Kws.Cells(KRow - 1, 2).value = ws.Cells(chkRow, 5).value Then
                Kws.Cells(KRow - 1, 4).value = upDate - 1
            Else
                Kws.Cells(KRow, 2).value = ws.Cells(chkRow, 5).value
                Kws.Cells(KRow, 4).value = upDate - 1
            End If

        ' 勤務時間に変更がある場合
        ElseIf orgValue = chgValue And orgHours <> chgHours Then
            If chgHours < 32 And orgHours >= 32 Then
                Kws.Cells(KRow, 2).Value = chgValue
                Kws.Cells(KRow, 3).value = upDate
                Kws.Cells(KRow, 5).Value = "非常勤"
            ElseIf chgHours >= 32 And orgHours < 32 Then
                Kws.Cells(KRow, 2).Value = chgValue
                Kws.Cells(KRow, 3).value = upDate
                Kws.Cells(KRow, 5).Value = "常勤"
            End If
        ElseIf orgValue = "" And chgValue = "" Then
            Exit For
        Else
            ' do nothing
        End If
NextRow:
    Next chkRow

    ' シェイプ設定呼び出し
    Call Shapes("pharmacy")
End Sub
