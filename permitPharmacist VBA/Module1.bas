Attribute VB_Name = "Module1"
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
        Kws.Cells(2, 5).value = "常勤"
    End If
    
    Debug.Print Application.Caller
    
    upDate = DateAdd("d", 29, ws.Cells(2, 1).value)
    ws.Cells(10, 3).ClearContents
    ws.Cells(11, 2).ClearContents
    ws.Cells(12, 2).ClearContents
    
    Kws.Cells(15, 2).value = ws.Cells(4, 1).value
    
    ' Initialize the array and counter
    sheetCount = 1
    ReDim sheetArray(1 To sheetCount)
    
    For chkRow = 3 To 12
        orgValue = ws.Cells(chkRow, 5).value
        chgValue = ws.Cells(chkRow, 9).value

        If orgValue <> chgValue And orgValue = "" Then
            If ws.Cells(11, 2).value = "" Then
                Set cover = ThisWorkbook.Worksheets("<保>変更届(その他薬剤師)")
                ws.Cells(10, 3).value = ws.Cells(chkRow, 11).value
                ws.Cells(11, 2).value = ws.Cells(chkRow, 9).value
            Else
                ws.Cells(12, 2).value = ws.Cells(chkRow, 9).value
            End If
            If ws.Cells(11, 2).value = "" And ws.Cells(chkRow, 11).value = "薬剤師" Then
                Set cover = ThisWorkbook.Worksheets("<保>変更届(その他薬剤師)")
                ws.Cells(10, 3).value = ws.Cells(chkRow, 7).value
            ElseIf ws.Cells(11, 2).value = "" And ws.Cells(chkRow, 11).value = "登録販売者" Then
                Set cover = ThisWorkbook.Worksheets("<保>変更届(その他登録販売者)")
                ws.Cells(10, 3).value = ws.Cells(chkRow, 7).value
            End If
        ElseIf ws.Cells(11, 2).value = "" And orgValue <> chgValue And chgValue = "" Then
            Set cover = ThisWorkbook.Worksheets("<保>変更届(その他薬剤師)")
            ws.Cells(10, 3).value = ws.Cells(chkRow, 7).value
        ElseIf ws.Cells(11, 2).value = "" And orgValue = chgValue And ws.Cells(chkRow, 6).value <> ws.Cells(chkRow, 10).value Then
            Set cover = ThisWorkbook.Worksheets("<保>変更届(その他薬剤師)")
            ws.Cells(10, 3).value = "薬剤師勤務時間"
        ElseIf orgValue = "" And chgValue = "" Then
            Exit For
        End If
    Next chkRow

    'On Error GoTo ErrorHandler
    ' Add sheets to the array
    sheetArray(sheetCount) = cover.Name
    sheetCount = sheetCount + 1
    ReDim Preserve sheetArray(1 To sheetCount)
    sheetArray(sheetCount) = "<保>変更届別紙1"
    
    If ws.Cells(11, 2).value <> "" And ws.Cells(10, 3).value = "薬剤師" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "雇用証明書（その他）"
    ElseIf ws.Cells(11, 2).value <> "" And ws.Cells(10, 3).value = "登録販売者" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "雇用証明書（登録販売者）"
    End If
    
    If ws.Cells(12, 2).value <> "" And ws.Cells(10, 3).value = "薬剤師" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "雇用証明書（その他） (2)"
    End If
    
    If ws.Cells(8, 9).value <> "" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "<保>変更届別紙2"
    End If
    
    If ws.Cells(4, 2).value > upDate Then
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
    
    Kws.Range("B3:D11").ClearContents
    
    If ws.Cells(2, 2) <> Kws.Cells(2, 1) Then
        Debug.Print ws.Cells(2, 2).value
        Kws.Cells(2, 1).value = ws.Cells(2, 2).value
    End If

    For chkRow = 3 To 12
        orgValue = ws.Cells(chkRow, 5).value
        chgValue = ws.Cells(chkRow, 9).value
        KRow = Kws.Cells(7, 2).End(xlUp).Row + 1
        Debug.Print KRow
        
        If orgValue <> chgValue And orgValue = "" Then ' 人員が増える場合
            If Kws.Cells(KRow - 1, 2).value = ws.Cells(chkRow, 9).value Then '対象者がすでに表上に記載されている場合
                ' do nothing
            Else
                Kws.Cells(KRow, 2).value = ws.Cells(chkRow, 9).value
                If Kws.Cells(16, 2).value <> "" Then '新人薬剤師の場合
                    Kws.Cells(KRow, 3).value = Kws.Cells(16, 2).value
                Else
                    Kws.Cells(KRow, 3).value = ws.Cells(2, 1).value
                End If
            End If

        ElseIf orgValue <> chgValue And chgValue = "" Then '人員が減る場合
            If Kws.Cells(KRow - 1, 2).value = ws.Cells(chkRow, 5).value Then
                Kws.Cells(KRow - 1, 4).value = ws.Cells(2, 1).value - 1
            Else
                Kws.Cells(KRow, 2).value = ws.Cells(chkRow, 5).value
                Kws.Cells(KRow, 4).value = ws.Cells(2, 1).value - 1
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

    ThisWorkbook.Worksheets("遅延理由書").Cells(10, 17).value = "管理薬剤師"
    ThisWorkbook.Worksheets("遅延理由書").Cells(18, 1).value = "保健所長"

    Set ws = Worksheets("検索")
    Set Kws = Worksheets("所属変更")
    
    If Kws.Cells(2, 1).value = ws.Cells(2, 2).value Then
        rc = MsgBox("入力内容を一旦リセットしますか？", vbYesNo + vbQuestion)
        If rc = vbYes Then
            Kws.Range("B3:D11").ClearContents
        End If
    Else
        Kws.Range("B3:D11").ClearContents
        Kws.Cells(2, 1).value = ws.Cells(2, 2).value
    End If
    
    upDate = DateAdd("d", 29, ws.Cells(2, 1).value)

    If ws.Cells(9, 1).value <> "" Then
        ws.Cells(11, 2).value = ws.Cells(9, 1).value
        changeDate = ws.Cells(2, 1).value
        upDate = DateAdd("d", 29, changeDate)
        Kws.Range("E2").value = t
        
        For a = 3 To 12
            If ws.Cells(9, 1).value = ws.Cells(a, 5).value Then
                anyChanges = False
            End If
        Next a
        
        If anyChanges = True Then
            cRow = Kws.Cells(12, 2).End(xlUp).Row + 1
            Kws.Cells(cRow, 2).value = ws.Range("I2").value
            Kws.Cells(cRow, 3).value = DateAdd("d", 0, ws.Range("A2").value)
        End If
        
        rc = MsgBox("旧管理者の異動はありますか？", vbYesNo + vbQuestion)

        If rc = vbYes Then
            cRow = Kws.Cells(12, 2).End(xlUp).Row + 1
            Kws.Cells(cRow, 2).value = ws.Range("E2").value
            Kws.Cells(cRow, 4).value = DateAdd("d", -1, ws.Range("A2").value)
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
        
        If ws.Cells(248, 3).value = "旭川市" Then
            sheetCount = sheetCount + 1
            ReDim Preserve sheetArray(1 To sheetCount)
            sheetArray(sheetCount) = "<保>旭川市・自立支援（育成更生）変更届"

            If ws.Cells(71, 3).value <> "" Then
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

            If ws.Cells(71, 3).value <> "" Then
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
        Kws.Cells(2, 5).value = "非常勤"
    End If
    
    If ThisWorkbook.Worksheets("検索").Cells(9, 1).value <> "" Then
        rc = MsgBox("厚生局異動に関して、管理者変更も同時に行いますか？", vbYesNo + vbQuestion)
        If rc = vbNo Then
            ThisWorkbook.Worksheets("検索").Cells(9, 1).ClearContents
        End If
    End If
    
    ' Initialize the array and counter
    sheetCount = 1
    ReDim sheetArray(1 To sheetCount)
    sheetArray(sheetCount) = "新<厚>異動届"
    
    If Kws.Cells(4, 2).value <> "" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "新<厚>別紙"
    End If
    
    If Kws.Cells(8, 2).value <> "" Then
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

    strName = ws.Cells(2, 2).value
    If logNo <> "" Then
    logNo = ws.Cells(18, 3).value
    End If
    logRow = logWS.Cells(Rows.count, 1).End(xlUp).Row + 1

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
    
    If Application.Caller = "管理薬剤師更新" Then
        updName = Kws.Cells(29, 1)
        updNumber = Kws.Cells(29, 3)
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

Sub 所属変更転記()
    Dim updRow As Integer, trsRow As Integer, n As Integer, startColumn As Integer
    Dim wkHour As Variant
    Dim updName As String, updNumber As String, strName As String
    Dim ws As Worksheet, updWS As Worksheet, Kws As Worksheet
    

    Set ws = ThisWorkbook.Worksheets("検索")
    Set updWS = ThisWorkbook.Worksheets("届出一覧テーブル")
    Set Kws = ThisWorkbook.Worksheets("所属変更")

    strName = ws.Cells(2, 2)
    
    ' 基準カラム（「非常勤薬剤師10」）の検索
    For startColumn = 1 To updWS.Columns.count
        If updWS.Cells(1, startColumn).value = "非常勤薬剤師10" Then
            startColumn = startColumn + 1
            Exit For
        End If
    Next startColumn

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

        updWS.Cells(trsRow, (updRow - 3) * 3 + startColumn) = updNumber
        updWS.Cells(trsRow, (updRow - 3) * 3 + startColumn + 1) = updName
        updWS.Cells(trsRow, (updRow - 3) * 3 + startColumn + 2) = wkHour
    End If

    Next updRow
    Kws.Cells(2, 1) = ws.Cells(2, 2)
    Call 初期化(ws)
    Call UpdatePharmacistInfoWithClass
End Sub
Sub 初期化(ws As Worksheet)
    ws.Range("E2").Formula = "=IF($C$4="""",C138,VLOOKUP(検索!$C$4,LOG!$A$2:$Y$1048507,3,0))"
    ws.Range("F2").Formula = "=VLOOKUP(E2,薬剤師マスタ!C:V,17,0)"
    ws.Range("E3:E12").Formula = "=IF($C$4="""",INDIRECT(""C""&ROW($A$171)+ROW()*3),VLOOKUP(検索!$C$4,LOG!$A$1:$Y$1048507,ROW()*2-1,0))"
    ws.Range("F3:F12").Formula = "=IFERROR(IF($C$4="""",IFERROR(INDIRECT(""C""&ROW($A$172)+ROW()*3),""""),IF(E3<>"""",VLOOKUP(検索!$C$4,LOG!$A$1:$Y$1048507,ROW()*2,0),"""")),J3)"
    ws.Range("I2").Formula = "=IF(A9<>"""",A9,E2)"
    'ws.Range("J2").Formula = "=DGET(薬剤師マスタ[#すべて],""週労働時間"",検索!I1:I2)"
    ws.Range("I3:I12").Formula = "=INDIRECT(""C""&ROW($A$171)+ROW()*3)"
    ws.Range("J3:J12").Formula = "=INDIRECT(""C""&ROW($A$172)+ROW()*3)"
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

    If cell.value = targetStr Then
        ' マッチする文字列が見つかったら対応するA列のセルの値を出力します
        ws.Cells(resultRow, 12).value = Sheets("届出一覧テーブル").Range("B" & cell.Row).value
        resultRow = resultRow + 1 ' 結果出力行を次に進めます
    End If

    Next cell

End Sub

Sub 薬剤師変更マクロ()
    Dim strName As String, PharmacistName As String, apdateItem As String
    Dim updateRow As Integer, updateColumn As Integer
    Dim updWS As Worksheet, ws As Worksheet

    Set updWS = Worksheets("届出一覧テーブル")
    Set Kws = Worksheets("所属変更")

    strName = Kws.Cells(2, 1).value
    apdateItem = Kws.Cells(27, 2).value
    PharmacistName = Kws.Cells(27, 1).value
    strColumn = 119
    For updateRow = 2 To 70
        If strName = updWS.Cells(updateRow, 2).value Then
            For updateColumn = strColumn To strColumn + 20
                If apdateItem = updWS.Cells(1, updateColumn).value Then
                    If Kws.Cells(27, 1) <> "" Then
                        updWS.Cells(updateRow, updateColumn).value = PharmacistName
                    Else: updWS.Cells(updateRow, updateColumn).value = ""
                    End If
                End If
            Next updateColumn
        End If
    Next updateRow

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
