Sub PrintToPdf(sheet As Worksheet, fileName As String)
    Dim saveName As String
    Dim d As String
    Dim strName As String
    Dim updateDate As String
    Dim pName As String
    Dim updateContent As String
    Dim KupdateContent As String
    Dim i As Integer

    ' Determine which subroutine called this one
    Select Case fileName
        Case "government"
            saveName = "【厚生局】異動届"
            strName = ThisWorkbook.Sheets("所属変更").Cells(2, 1).Value & Format(ThisWorkbook.Sheets("所属変更").Cells(19, 2).Value, "0000")
            updateDate = Format(ThisWorkbook.Sheets("所属変更").Cells(3, 3).Value, "yyyymmdd")
            updateContent = "_" & ThisWorkbook.Sheets("所属変更").Cells(3, 2).Value

            For i = 1 To 9
                If ThisWorkbook.Sheets("所属変更").Cells(i + 2, 2).Value <> "" Then
                    updateContent = updateContent & ThisWorkbook.Sheets("所属変更").Cells(i + 2, 2).Value
                    Select Case True
                        Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 3) <> "" And ThisWorkbook.Sheets("所属変更").Cells(i + 2, 4) <> ""
                            updateContent = updateContent & "(±非)"
                        Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 3) <> ""
                            updateContent = updateContent & "(+非)"
                        Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 4) <> ""
                            updateContent = updateContent & "(-非)"
                    End Select
                Else
                    Exit For
                End If
            Next i

        Case "pharmacy"
            saveName = "【保健所】その他薬剤師変更"
            strName = ThisWorkbook.Sheets("検索").Cells(2, 2).Value & Format(ThisWorkbook.Sheets("検索").Cells(19, 3).Value, "0000")
            updateDate = Format(ThisWorkbook.Sheets("検索").Cells(2, 1).Value, "yyyymmdd")
            updateContent = "_" & ThisWorkbook.Sheets("検索").Cells(11, 2).Value & "(+" & ThisWorkbook.Sheets("検索").Cells(11, 3).Value & "hr)"
            If ThisWorkbook.Sheets("検索").Cells(12, 2).Value <> "" Then
                updateContent = updateContent & ThisWorkbook.Sheets("検索").Cells(12, 2).Value & "(+" & ThisWorkbook.Sheets("検索").Cells(12, 3).Value & "hr)"
            End If
            KupdateContent = "_" & ThisWorkbook.Sheets("所属変更").Cells(3, 2).Value

            For i = 1 To 9
                Select Case True
                    Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 3).Value <> "" And ThisWorkbook.Sheets("所属変更").Cells(i + 2, 4).Value <> ""
                        KupdateContent = KupdateContent & ThisWorkbook.Sheets("所属変更").Cells(i + 2, 2).Value & "(±常)"
                    Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 3).Value <> ""
                        KupdateContent = KupdateContent & ThisWorkbook.Sheets("所属変更").Cells(i + 2, 2).Value & "(+常)"
                    Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 4).Value <> ""
                        KupdateContent = KupdateContent & ThisWorkbook.Sheets("所属変更").Cells(i + 2, 2).Value & "(-常)"
                End Select
            Next i

        Case "admin"
            saveName = "【厚生局・保健所・振興局・労働局】管理薬剤師変更"
            strName = ThisWorkbook.Sheets("検索").Cells(2, 2).Value & Format(ThisWorkbook.Sheets("検索").Cells(19, 3).Value, "0000")
            updateDate = Format(ThisWorkbook.Sheets("検索").Cells(2, 1).Value, "yyyymmdd")
            updateContent = "_" & ThisWorkbook.Sheets("検索").Cells(7, 1).Value & "→" & ThisWorkbook.Sheets("検索").Cells(9, 1).Value
            KupdateContent = "_" & ThisWorkbook.Sheets("所属変更").Cells(3, 2).Value

            For i = 1 To 9
                Select Case True
                    Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 3).Value <> "" And ThisWorkbook.Sheets("所属変更").Cells(i + 2, 4).Value <> ""
                        KupdateContent = KupdateContent & "(±常)"
                    Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 3).Value <> ""
                        KupdateContent = KupdateContent & "(+常)"
                    Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 4).Value <> ""
                        KupdateContent = KupdateContent & "(-常)"
                End Select
            Next i
    End Select

    ' Get valid PDF folder path
    'On Error GoTo ErrorHandler
    pName = ThisWorkbook.Path & Application.PathSeparator & "PDFs" & Application.PathSeparator & updateDate & strName & saveName & updateContent & d & ".pdf"

    Set RWs = ThisWorkbook.Sheets("作成書類リネーム用")

    Debug.Print pName
    Select Case fileName
        Case "government"
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【厚生局】異動届" & updateContent
        Case "pharmacy"
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【厚生局】異動届" & KupdateContent
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【保健所】その他薬剤師変更届" & updateContent
        Case "admin"
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【厚生局】異動届" & updateContent & KupdateContent
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【保健所】管理薬剤師変更届" & updateContent
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【保健所】高度管理機器管理者変更届" & updateContent
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【保健所】自立支援(育生更生)管理薬剤師変更届" & updateContent
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【振興局】自立支援(精神通院)管理薬剤師変更届" & updateContent
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【労働局】管理薬剤師変更届" & updateContent
    End Select
    Call makePdfs(sheet, pName)
Exit Sub
ErrorHandler:
MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

Sub Shapes(fileName As String)
    Dim Kws As Worksheet
    Dim j As Integer
    Dim sheetProtected1 As Boolean
    Dim sheetProtected2 As Boolean
    
    On Error GoTo ErrorHandler

    Set Kws = ThisWorkbook.Worksheets("所属変更")
    
    ' シートの保護を解除
    sheetProtected1 = ThisWorkbook.Worksheets("新<厚>異動届").ProtectContents
    If sheetProtected1 Then ThisWorkbook.Worksheets("新<厚>異動届").Unprotect
    
    ' "新<厚>異動届"シートのシェイプを設定
    ThisWorkbook.Worksheets("新<厚>異動届").Shapes("管薬").Visible = False
    ThisWorkbook.Worksheets("新<厚>異動届").Shapes("チェック1").Visible = False
    ThisWorkbook.Worksheets("新<厚>異動届").Shapes("チェック2").Visible = False

    If fileName = "admin" Then
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("管薬").Visible = True
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("チェック1").Visible = True
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("チェック2").Visible = True
    End If
    
    If Kws.Cells(3, 5).Value = "常勤" Then
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("常勤").Visible = True
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("非常勤").Visible = False
    ElseIf Kws.Cells(3, 5).Value = "非常勤" Then
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("常勤").Visible = False
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("非常勤").Visible = True
    Else
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("常勤").Visible = False
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("非常勤").Visible = False
    End If

    If Kws.Cells(3, 1).Value = "転入" Then
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("転入").Visible = True
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("転出").Visible = False
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("入薬").Visible = True
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("出薬").Visible = False
    ElseIf Kws.Cells(3, 1).Value = "転出" Then
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("転入").Visible = False
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("転出").Visible = True
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("入薬").Visible = False
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("出薬").Visible = True
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("常勤").Visible = False
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("非常勤").Visible = False
    Else
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("転入").Visible = False
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("転出").Visible = False
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("入薬").Visible = False
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("出薬").Visible = False
    End If
    
    ' 必要に応じて別のシートの保護も解除
    sheetProtected2 = ThisWorkbook.Worksheets("新<厚>別紙").ProtectContents
    If sheetProtected2 Then ThisWorkbook.Worksheets("新<厚>別紙").Unprotect
    
    ' "新<厚>別紙"シートのシェイプを設定
    For j = 1 To 4
        If Kws.Cells(j + 3, 1).Value = "転入" Then
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("転入" & j).Visible = True
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("転出" & j).Visible = False
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("入薬" & j).Visible = True
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("出薬" & j).Visible = False
        ElseIf Kws.Cells(j + 3, 1).Value = "転出" Then
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("転入" & j).Visible = False
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("転出" & j).Visible = True
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("入薬" & j).Visible = False
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("出薬" & j).Visible = True
        Else
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("転入" & j).Visible = False
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("転出" & j).Visible = False
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("入薬" & j).Visible = False
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("出薬" & j).Visible = False
        End If

        If Kws.Cells(j + 3, 5).Value = "常勤" Then
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("常勤" & j).Visible = True
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("非常勤" & j).Visible = False
        ElseIf Kws.Cells(j + 3, 5).Value = "非常勤" Then
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("常勤" & j).Visible = False
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("非常勤" & j).Visible = True
        Else
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("常勤" & j).Visible = False
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("非常勤" & j).Visible = False
        End If
    Next j

    ' シートの保護を再度有効化
    If sheetProtected1 Then ThisWorkbook.Worksheets("新<厚>異動届").Protect
    If sheetProtected2 Then ThisWorkbook.Worksheets("新<厚>別紙").Protect

Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
End Sub
Sub 一般薬剤師異動時必要書類印刷()
    Dim n As Integer, chkRow As Integer
    Dim upDate As Date
    Dim orgValue As String, chgValue As String
    Dim ws As Worksheet, cover As Worksheet, Kws As Worksheet
    Dim sheetArray() As String
    Dim sheetCount As Integer

    Set ws = ThisWorkbook.Worksheets("検索")
    Set Kws = ThisWorkbook.Worksheets("所属変更")
    
    If Application.Caller = "その他薬剤師変更" Then
        Kws.Cells(2, 5).Value = "常勤"
    End If
    
    Debug.Print Application.Caller
    
    upDate = DateAdd("d", 29, ws.Cells(2, 1).Value)
    ws.Cells(10, 3).ClearContents
    ws.Cells(11, 2).ClearContents
    ws.Cells(12, 2).ClearContents
    
    Kws.Cells(15, 2).Value = ws.Cells(4, 1).Value
    
    ' Initialize the array and counter
    sheetCount = 1
    ReDim sheetArray(1 To sheetCount)
    
    For chkRow = 3 To 12
        orgValue = ws.Cells(chkRow, 5).Value
        chgValue = ws.Cells(chkRow, 9).Value

        If orgValue <> chgValue And orgValue = "" Then
            If ws.Cells(11, 2).Value = "" Then
                Set cover = ThisWorkbook.Worksheets("<保>変更届(その他薬剤師)")
                ws.Cells(10, 3).Value = ws.Cells(chkRow, 11).Value
                ws.Cells(11, 2).Value = ws.Cells(chkRow, 9).Value
            Else
                ws.Cells(12, 2).Value = ws.Cells(chkRow, 9).Value
            End If
            If ws.Cells(11, 2).Value = "" And ws.Cells(chkRow, 11).Value = "薬剤師" Then
                Set cover = ThisWorkbook.Worksheets("<保>変更届(その他薬剤師)")
                ws.Cells(10, 3).Value = ws.Cells(chkRow, 7).Value
            ElseIf ws.Cells(11, 2).Value = "" And ws.Cells(chkRow, 11).Value = "登録販売者" Then
                Set cover = ThisWorkbook.Worksheets("<保>変更届(その他登録販売者)")
                ws.Cells(10, 3).Value = ws.Cells(chkRow, 7).Value
            End If
        ElseIf ws.Cells(11, 2).Value = "" And orgValue <> chgValue And chgValue = "" Then
            Set cover = ThisWorkbook.Worksheets("<保>変更届(その他薬剤師)")
            ws.Cells(10, 3).Value = ws.Cells(chkRow, 7).Value
        ElseIf ws.Cells(11, 2).Value = "" And orgValue = chgValue And ws.Cells(chkRow, 6).Value <> ws.Cells(chkRow, 10).Value Then
            Set cover = ThisWorkbook.Worksheets("<保>変更届(その他薬剤師)")
            ws.Cells(10, 3).Value = "薬剤師勤務時間"
        ElseIf orgValue = "" And chgValue = "" Then
            Exit For
        End If
    Next chkRow

    On Error GoTo ErrorHandler
    ' Add sheets to the array
    sheetArray(sheetCount) = cover.name
    sheetCount = sheetCount + 1
    ReDim Preserve sheetArray(1 To sheetCount)
    sheetArray(sheetCount) = "<保>変更届別紙1"
    
    If ws.Cells(11, 2).Value <> "" And ws.Cells(10, 3).Value = "薬剤師" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "雇用証明書（その他）"
    ElseIf ws.Cells(11, 2).Value <> "" And ws.Cells(10, 3).Value = "登録販売者" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "雇用証明書（登録販売者）"
    End If
    
    If ws.Cells(12, 2).Value <> "" And ws.Cells(10, 3).Value = "薬剤師" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "雇用証明書（その他） (2)"
    End If
    
    If ws.Cells(8, 9).Value <> "" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "<保>変更届別紙2"
    End If
    
    If ws.Cells(4, 2).Value > upDate Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "遅延理由書"
    End If
    
    ' Prompt for 厚生局異動届 creation
    Dim rc As VbMsgBoxResult
    rc = MsgBox("厚生局異動届に変更を加えますか？", vbYesNo + vbQuestion)
    If rc = vbYes Then
        Call 厚生局異動届作成
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "新<厚>異動届"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "新<厚>別紙"
    End If
    
    ' Select all sheets in the array and export to PDF
    Sheets(sheetArray).Select
    Call PrintToPdf(ActiveSheet, "pharmacy")
    Call 変更ログ記録
    
    ws.Cells(10, 3).ClearContents
    Kws.Cells(15, 2).ClearContents
    Exit Sub

ErrorHandler:
    MsgBox " 変更箇所がみつかりません"
End Sub
Sub 厚生局異動届作成()
    Dim chkRow As Integer, KRow As Integer
    Dim upDate As Date
    Dim orgValue As String, chgValue As String
    Dim ws As Worksheet, cover As Worksheet, Kws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("検索")
    Set Kws = ThisWorkbook.Worksheets("所属変更")
    
    Kws.Range("A3:D11").ClearContents
    
    If ws.Cells(2, 2) <> Kws.Cells(2, 1) Then
        Debug.Print ws.Cells(2, 2).Value
        Kws.Cells(2, 1).Value = ws.Cells(2, 2).Value
    End If

    For chkRow = 3 To 12
        orgValue = ws.Cells(chkRow, 5).Value
        chgValue = ws.Cells(chkRow, 9).Value
        KRow = Kws.Cells(7, 2).End(xlUp).Row + 1
        Debug.Print KRow
        
        If orgValue <> chgValue And orgValue = "" Then ' 人員が増える場合
            If Kws.Cells(KRow - 1, 2).Value = ws.Cells(chkRow, 9).Value Then '対象者がすでに表上に記載されている場合
                ' do nothing
            Else
                Kws.Cells(KRow, 2).Value = ws.Cells(chkRow, 9).Value
                If Kws.Cells(16, 2).Value <> "" Then '新人薬剤師の場合
                    Kws.Cells(KRow, 3).Value = Kws.Cells(16, 2).Value
                Else
                    Kws.Cells(KRow, 3).Value = ws.Cells(2, 1).Value
                End If
            End If

        ElseIf orgValue <> chgValue And chgValue = "" Then '人員が減る場合
            If Kws.Cells(KRow - 1, 2).Value = ws.Cells(chkRow, 5).Value Then
                Kws.Cells(KRow - 1, 4).Value = ws.Cells(2, 1).Value - 1
            Else
                Kws.Cells(KRow, 2).Value = ws.Cells(chkRow, 5).Value
                Kws.Cells(KRow, 4).Value = ws.Cells(2, 1).Value - 1
            End If

        ElseIf orgValue = "" And chgValue = "" Then
            Exit For
        Else
            ' do nothing
        End If
    Next chkRow

    Call Shapes("pharmacy")
End Sub
Sub 管理薬剤師変更()
    Dim i As Integer, n As Integer, j As Integer, a As Integer, anyChanges As Boolean
    Dim ws As Worksheet, Kws As Worksheet
    Dim upDate As Date, changeDate As Date
    Dim t As String, F As String
    Dim rc As VbMsgBoxResult
    Dim sheetArray() As String
    Dim sheetCount As Integer
    
    t = "常勤"
    F = "非常勤"
    anyChanges = True

    ThisWorkbook.Worksheets("遅延理由書").Cells(10, 17).Value = "管理薬剤師"
    ThisWorkbook.Worksheets("遅延理由書").Cells(18, 1).Value = "保健所長"

    Set ws = Worksheets("検索")
    Set Kws = Worksheets("所属変更")
    
    If Kws.Cells(2, 1).Value = ws.Cells(2, 2).Value Then
        rc = MsgBox("入力内容を一旦リセットしますか？", vbYesNo + vbQuestion)
        If rc = vbYes Then
            Kws.Range("A3:D7").ClearContents
        End If
    Else
        Kws.Cells(2, 1).Value = ws.Cells(2, 2).Value
    End If
    
    upDate = DateAdd("d", 29, ws.Cells(2, 1).Value)

    If ws.Cells(9, 1).Value <> "" Then
        ws.Cells(11, 2).Value = ws.Cells(9, 1).Value
        changeDate = ws.Cells(2, 1).Value
        upDate = DateAdd("d", 29, changeDate)
        Kws.Range("E2").Value = t
        
        For a = 3 To 12
            If ws.Cells(9, 1).Value = ws.Cells(a, 5).Value Then
                anyChanges = False
            End If
        Next a
        
        If anyChanges = True Then
            cRow = Kws.Cells(12, 2).End(xlUp).Row + 1
            Kws.Cells(cRow, 2).Value = ws.Range("I2").Value
            Kws.Cells(cRow, 3).Value = DateAdd("d", 0, ws.Range("A2").Value)
        End If
        
        rc = MsgBox("旧管理者の異動はありますか？", vbYesNo + vbQuestion)

        If rc = vbYes Then
            cRow = Kws.Cells(12, 2).End(xlUp).Row + 1
            Kws.Cells(cRow, 2).Value = ws.Range("E2").Value
            Kws.Cells(cRow, 4).Value = DateAdd("d", -1, ws.Range("A2").Value)
        End If

        Call Shapes("admin")
        
        ' Initialize the array and counter
        sheetCount = 1
        ReDim sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "新<厚>異動届"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "新<厚>別紙"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "<保>変更届 (管理者)"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "<保>変更届別紙1"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "雇用証明書（管理者）"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "<保>高度管理（管理者）"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "雇用証明書（高度管理）"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "<振>自立支援（精神）変更届（旭川も含む）"
        
        If ws.Cells(248, 3).Value = "旭川市" Then
            sheetCount = sheetCount + 1
            ReDim Preserve sheetArray(1 To sheetCount)
            sheetArray(sheetCount) = "<保>旭川市・自立支援（育成更生）変更届"

            If ws.Cells(71, 3).Value <> "" Then
                sheetCount = sheetCount + 1
                ReDim Preserve sheetArray(1 To sheetCount)
                sheetArray(sheetCount) = "<旭>毒劇物取扱責任者"
                sheetCount = sheetCount + 1
                ReDim Preserve sheetArray(1 To sheetCount)
                sheetArray(sheetCount) = "雇用証明書（毒劇物）"
            End If
        Else
            sheetCount = sheetCount + 1
            ReDim Preserve sheetArray(1 To sheetCount)
            sheetArray(sheetCount) = "<保>自立支援（育成更生）変更届(旭川以外・北海道)"

            If ws.Cells(71, 3).Value <> "" Then
                sheetCount = sheetCount + 1
                ReDim Preserve sheetArray(1 To sheetCount)
                sheetArray(sheetCount) = "<北>毒劇物取扱責任者"
                sheetCount = sheetCount + 1
                ReDim Preserve sheetArray(1 To sheetCount)
                sheetArray(sheetCount) = "雇用証明書（毒劇物）"
            End If
        End If
        
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "<労>管理薬剤師変更届"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "許認可提出状況確認票"
        
        ' Select all sheets in the array and export to PDF
        Sheets(sheetArray).Select
        Call PrintToPdf(ActiveSheet, "admin")
        
        Call 変更ログ記録
        Call FindMatchingStrings
    Else
        MsgBox "A9 に変更後の管理薬剤師名を入力してください"
    End If
End Sub
Sub 厚生局所属変更書類PDF()
    Dim rc As VbMsgBoxResult
    Dim j As Integer
    Dim upDate As Date
    Dim Kws As Worksheet
    Dim ArrayShName() As String
    Dim sheetCount As Integer
    
    Set Kws = ThisWorkbook.Worksheets("所属変更")
    
    If Application.Caller = "応援者登録" Then
        Kws.Cells(2, 5).Value = "非常勤"
    End If
    
    If ThisWorkbook.Worksheets("検索").Cells(9, 1).Value <> "" Then
        rc = MsgBox("厚生局異動に関して、管理者変更も同時に行いますか？", vbYesNo + vbQuestion)
        If rc = vbNo Then
            ThisWorkbook.Worksheets("検索").Cells(9, 1).ClearContents
        End If
    End If
    
    ' Initialize the array and counter
    sheetCount = 1
    ReDim sheetArray(1 To sheetCount)
    sheetArray(sheetCount) = "新<厚>異動届"
    
    If Kws.Cells(4, 2).Value <> "" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "新<厚>別紙"
    End If
    
    If Kws.Cells(8, 2).Value <> "" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "新<厚>別紙 (2)"
    End If
    
    Call Shapes("pharmacy")
    ' Select all sheets in the array and export to PDF
    Sheets(sheetArray).Select
    Call PrintToPdf(ActiveSheet, "government")
    
    ThisWorkbook.Worksheets("所属変更").Select
End Sub
Sub 変更ログ記録()
    Dim strName As String, logNo As String
    Dim logRow As Long, updRow As Long
    Dim ws As Worksheet, updWS As Worksheet, logWS As Worksheet

    Set ws = ThisWorkbook.Worksheets("検索")
    Set logWS = ThisWorkbook.Worksheets("LOG")

    strName = ws.Cells(2, 2).Value
    logNo = ws.Cells(18, 3).Value
    logRow = logWS.Cells(Rows.Count, 1).End(xlUp).Row + 1

    Debug.Print "strName =" & strName
    Debug.Print logNo
    Debug.Print logRow

    logWS.Cells(logRow, 1) = logNo
    logWS.Cells(logRow, 2) = strName
    For i = 1 To 11
    logWS.Cells(logRow, i * 2 + 1) = ws.Cells(i + 1, 5)
    logWS.Cells(logRow, i * 2 + 2) = ws.Cells(i + 1, 6)
    Next i

    logWS.Cells(logRow + 1, 1) = logNo & "!"
    logWS.Cells(logRow + 1, 2) = strName
    For i = 1 To 11
    logWS.Cells(logRow + 1, i * 2 + 1) = ws.Cells(i + 1, 9)
    logWS.Cells(logRow + 1, i * 2 + 2) = ws.Cells(i + 1, 10)
    Next i
    
End Sub
Sub 管薬所属変更転記()
    Dim i As Integer, trsRow As Integer, n As Integer, trsColumn As Integer, c As Integer
    Dim wkHour As Single
    Dim updName As String, updNumber As String, strName As String
    Dim ws As Worksheet, updWS As Worksheet
    Dim contentsArray() As Variant
    
    contentsArray = Array("保健所管理薬剤師", "高度管理医療機器等販売管理者", "自立支援（精神通院医療）担当者", "自立支援（更生育成医療）担当者", "生活保護法管理薬剤師", "労災保険（管理薬剤師）", "薬局機能情報（管理薬剤師）", "リタリン管理者", "ADHD流通管理 責任者", "厚生局管理薬剤師", "毒劇物取扱責任者")
    
    Set ws = ThisWorkbook.Worksheets("検索")
    Set Kws = ThisWorkbook.Worksheets("所属変更")
    Set updWS = ThisWorkbook.Worksheets("届出一覧テーブル")
    strName = ws.Cells(2, 2)
    
    If Application.Caller = "管理薬剤師更新" Then
        updName = Kws.Cells(21, 1)
        updNumber = Kws.Cells(21, 3)
        strName = Kws.Cells(2, 1)
    Else
        updName = ws.Cells(9, 1)
        updNumber = ws.Cells(9, 2)
        strName = ws.Cells(2, 2)
    End If
    
    For n = 2 To 150

    If strName = updWS.Cells(n, 2) Then
        trsRow = n
    Exit For
    End If

    Next n

    For i = 0 To 9
        For c = 50 To 150
        If updWS.Cells(1, c) = contentsArray(i) Then
            trsColumn = c
            updWS.Cells(trsRow, trsColumn) = updName
            updWS.Cells(trsRow, trsColumn - 1) = updNumber
            Exit For
        End If
        Next c
    Next i
    
    If ws.Cells(68, 3) <> "" Then
        For c = 50 To 150
        If updWS.Cells(1, c) = contentsArray(10) Then
            trsColumn = c
            updWS.Cells(trsRow, trsColumn) = updName
            updWS.Cells(trsRow, trsColumn - 1) = updNumber
            Exit For
        End If
        Next c
    End If
End Sub
Sub makePdfs(sheet As Worksheet, pName As String)
    On Error GoTo ErrorHandler
    sheet.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    fileName:=pName, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, _
    OpenAfterPublish:=True
    Exit Sub

ErrorHandler:
    MsgBox "PDFの作成中にエラーが発生しました"
End Sub
Sub 所属変更転記()
    Dim updRow As Integer, trsRow As Integer, n As Integer
    Dim wkHour As Variant
    Dim updName As String, updNumber As String, strName As String
    Dim ws As Worksheet, updWS As Worksheet

    Set ws = ThisWorkbook.Worksheets("検索")
    Set updWS = ThisWorkbook.Worksheets("届出一覧テーブル")

    strName = ws.Cells(2, 2)

    For n = 2 To 70

    If strName = updWS.Cells(n, 2) Then
        trsRow = n
        Exit For
    End If

    Next n

    For updRow = 3 To 12

    If ws.Cells(updRow, 5) = "" And ws.Cells(updRow, 9) = "" Then
        Exit For
    Else
        updNumber = ws.Cells(updRow, 13)
        updName = ws.Cells(updRow, 9)
        wkHour = ws.Cells(updRow, 10)

        updWS.Cells(trsRow, (updRow * 3) + 151) = updNumber
        updWS.Cells(trsRow, (updRow * 3) + 152) = updName
        updWS.Cells(trsRow, (updRow * 3) + 153) = wkHour
    End If

    Next updRow
End Sub
Sub 厚生局異動届印刷確認()
    Dim rc As VbMsgBoxResult
    rc = MsgBox("厚生局異動届を印刷しますか？", vbYesNo + vbQuestion)

    If rc = vbYes Then
        Call 厚生局所属変更書類PDF
    End If

End Sub
Sub FindMatchingStrings()
    Dim rng As Range
    Dim cell As Range
    Dim targetStr As String
    Dim ws As Worksheet
    Dim resultWs As Worksheet
    Dim resultRow As Integer

    Set ws = ThisWorkbook.Sheets("検索") ' シート名を適切に設定してください
    
    ws.Range("L2:L30").Clear
    ' 検索したい文字列を定義します
    targetStr = ws.Cells(16, 2)

    ' 検索対象の範囲を設定します
    Set rng = Sheets("届出一覧テーブル").Range("CU2:GX70")  ' B2:B100 を検索対象の範囲に置き換えてください

    resultRow = 2 ' 結果を出力する開始行（ここではSheet2の1行目から開始）
    ' 各セルをループで調べます
    For Each cell In rng

    If cell.Value = targetStr Then
        ' マッチする文字列が見つかったら対応するA列のセルの値を出力します
        ws.Cells(resultRow, 12).Value = Sheets("届出一覧テーブル").Range("B" & cell.Row).Value
        resultRow = resultRow + 1 ' 結果出力行を次に進めます
    End If

    Next cell

End Sub
Sub 薬剤師変更マクロ()
    Dim strName As String, sptName As String, apdateItem As String
    Dim enployeeNumber As Long, workTime As Single
    Dim apdateRow As Integer, apdateColumn As Integer
    Dim apdWS As Worksheet, ws As Worksheet
    
    Set apdWS = Worksheets("届出一覧テーブル")
    Set ws = Worksheets("所属変更")
    
    
    
    If ws.Cells(3, 13) <> "" Then
        strName = ws.Cells(2, 1).Value
        apdateItem = ws.Cells(3, 15).Value
        sptName = ws.Cells(3, 13).Value
        enployeeNumber = ws.Cells(3, 12).Value
        workTime = ws.Cells(3, 14).Value
        
        For apdateRow = 2 To 70

        If strName = apdWS.Cells(apdateRow, 2).Value Then
            For apdateColumn = 121 To 188

            If apdateItem = apdWS.Cells(1, apdateColumn).Value Then
                apdWS.Cells(apdateRow, apdateColumn - 1).Value = Format(enployeeNumber, "0000000")
                apdWS.Cells(apdateRow, apdateColumn).Value = sptName
                    If apdateColumn > 160 Then
                        apdWS.Cells(apdateRow, apdateColumn + 1).Value = workTime
                    End If
                Exit For
            End If

            Next apdateColumn
        End If

        Next apdateRow
            
        Else
        strName = ws.Cells(2, 1).Value
        apdateItem = ws.Cells(3, 15).Value
        For apdateRow = 2 To 50

        If strName = apdWS.Cells(apdateRow, 2).Value Then
            For apdateColumn = 121 To 188

            If apdateItem = apdWS.Cells(1, apdateColumn).Value Then
                apdWS.Cells(apdateRow, apdateColumn - 1).Value = ""
                apdWS.Cells(apdateRow, apdateColumn).Value = ""

                    If apdateColumn > 160 Then
                    apdWS.Cells(apdateRow, apdateColumn + 1).Value = ""
                    End If

            Exit For
            End If

            Next apdateColumn
        End If

        Next apdateRow
    End If

    MsgBox "所属薬剤師のデータベース更新が完了しました！"
    Call 変更ログ記録
End Sub
Sub 新店許認可登録()
    Dim i As Integer, n As Integer, updColumn As Integer
    Dim updDate As Date
    Dim updContents As String
    Dim ws As Worksheet, updWS As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("検索")
    Set updWS = ThisWorkbook.Worksheets("届出一覧テーブル")
    strName = ws.Cells(14, 9)
    
    For n = 2 To 100
        If strName = updWS.Cells(n, 2) Then
            trsRow = n
            Exit For
        End If
    Next n
    
    For i = 16 To 39
    
    If ws.Cells(i, 9) <> "" Then
        updContents = ws.Cells(i, 9)
        For updColumn = 1 To 74
            If ws.Cells(i, 8) = updWS.Cells(1, updColumn) Then
                updWS.Cells(n, updColumn) = updContents
            End If
        Next updColumn
    End If
    Next i
    
End Sub
Sub ProcessSupporterFilesAndCreatePDFs()
    Dim folderPath As String
    Dim exfileName As String
    Dim wb As Workbook
    Dim Kws As Worksheet
    Dim supporterData As Variant
    Dim i As Long
    Dim supporterName As String
    Dim startDate As Variant
    Dim endDate As Variant
    
    ' Excelファイルのあるフォルダパスを指定
    folderPath = "/Users/yoshipc/Desktop/令和6年3月応援者リスト/" ' <-- フォルダパスを適宜変更してください
    
    ' 所属変更シートの設定
    Set Kws = ThisWorkbook.Sheets("所属変更")
    
    ' フォルダ内の最初のファイルを取得
    exfileName = Dir(folderPath & "*.xlsx")
    
    ' フォルダ内の全てのファイルをループ
    Do While exfileName <> ""
        ' 各ファイルを開く
        Set wb = Workbooks.Open(folderPath & exfileName)
        
        ' 店舗名をA2セルにセット
        Kws.Cells(2, 1).Value = wb.Sheets(1).Range("A1").Value
        
        ' 必要なデータ（B列, C列, D列）を取得
        supporterData = GetSupporterDataFromSheet(wb.Sheets(1))
        
        ' データを所属変更シートに反映
        For i = LBound(supporterData, 1) + 2 To UBound(supporterData, 1)
            supporterName = supporterData(i, 1)
            startDate = supporterData(i, 2)
            endDate = supporterData(i, 3)
            
            If supporterName = "" Then
                Exit For
            End If
            
            ' 日付を文字列から日付型に変換
            If Not IsDate(startDate) Then
                startDate = CDate(startDate)
            End If
            If Not IsDate(endDate) Then
                endDate = CDate(endDate)
            End If
            
            ' 所属変更シートにデータを更新
            UpdateSupporterInSheet Kws, supporterName, startDate, endDate
        Next i
        
        ' 処理が終わったらファイルを閉じる
        wb.Close False
        
        ' PDFを作成（必要に応じて）
        ThisWorkbook.Activate
        Call 厚生局所属変更書類PDF
        
        ' 次のファイルを取得
        exfileName = Dir
    Loop
End Sub
'
'Function getPdfFolder() As String
'    Dim pdfFolderPath As String
'    ' PDFフォルダのパスを設定します
'    pdfFolderPath = ThisWorkbook.Path & Application.PathSeparator & "PDFs"
'
'    ' PDFフォルダが存在しない場合は作成します
'    If Dir(pdfFolderPath, vbDirectory) = "" Then
'        MkDir pdfFolderPath
'    End If
'
'    getPdfFolder = pdfFolderPath
'End Function

Function GetSupporterDataFromSheet(ws As Worksheet) As Variant
    Dim lastRow As Long
    Dim dataRange As Range
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' データ範囲を設定
    Set dataRange = ws.Range("B2:D" & lastRow)
    
    ' データを配列に変換して返す
    GetSupporterDataFromSheet = dataRange.Value
End Function

Sub UpdateSupporterInSheet(Kws As Worksheet, name As String, startDate As Variant, endDate As Variant)
    Dim lastRow As Long
    Dim found As Range
    
    ' 名前がすでにシートにあるかを確認
    Set found = Kws.Columns("B").Find(name, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' 日付データがDate型でない場合、変換
    If Not IsDate(startDate) Then
        startDate = CDate(startDate)
    End If
    If Not IsDate(endDate) Then
        endDate = CDate(endDate)
    End If
    
    If Not found Is Nothing Then
        ' 名前が見つかった場合、最終日付を更新
        If found.Offset(0, 2).Value < endDate Then
            found.Offset(0, 2).Value = endDate
        End If
    Else
        ' 新しい行に追加
        lastRow = Kws.Cells(11, "B").End(xlUp).Row + 1
        Kws.Cells(lastRow, 2).Value = name
        Kws.Cells(lastRow, 3).Value = startDate
        Kws.Cells(lastRow, 4).Value = endDate
    End If
End Sub
Sub ExampleDir()
    Dim folderPath As String
    Dim fileName As String
    
    ' 検索するフォルダを指定
    folderPath = "/Users/yoshipc/Desktop/令和6年3月応援者リスト/"
    
    ' 最初のファイルを取得
    fileName = Dir(folderPath & "*.xlsx")
    
    ' ループでフォルダ内のすべての .xlsx ファイルを取得
    Do While fileName <> ""
        ' ファイル名を出力
        Debug.Print "Found file: " & fileName
        
        ' 次のファイルを取得
        fileName = Dir
    Loop
End Sub
