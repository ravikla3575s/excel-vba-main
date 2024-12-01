Sub PrintToPdf(ByVal sheet As Worksheet, ByVal fileName As String)
    Dim saveName As String
    Dim d As String
    Dim strName As String
    Dim updateDate As String
    Dim pName As String
    Dim updateContent As String
    Dim KupdateContent As String
    Dim i As Integer
    Dim pdfFolder As String

    Debug.Print fileName

    ' Determine which subroutine called this one
    Select Case fileName
        Case "government"
            saveName = "【厚生局】異動届"
            strName = ThisWorkbook.Sheets("所属変更").Cells(2, 1).value & Format(ThisWorkbook.Sheets("所属変更").Cells(25, 2).value, "0000")
            updateDate = Format(ThisWorkbook.Sheets("所属変更").Cells(3, 3).value, "yyyymmdd")
            updateContent = "_" & ThisWorkbook.Sheets("所属変更").Cells(3, 2).value

            For i = 1 To 9
                If ThisWorkbook.Sheets("所属変更").Cells(i + 2, 2).value <> "" Then
                    updateContent = updateContent & ThisWorkbook.Sheets("所属変更").Cells(i + 2, 2).value
                    Select Case True
                        Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 3).value <> "" And ThisWorkbook.Sheets("所属変更").Cells(i + 2, 4).value <> ""
                            updateContent = updateContent & "(±非)"
                        Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 3).value <> ""
                            updateContent = updateContent & "(+非)"
                        Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 4).value <> ""
                            updateContent = updateContent & "(-非)"
                    End Select
                Else
                    Exit For
                End If
            Next i

        Case "pharmacy"
            saveName = "【保健所】その他薬剤師変更"
            strName = ThisWorkbook.Sheets("検索").Cells(2, 2).value & Format(ThisWorkbook.Sheets("検索").Cells(19, 3).value, "0000")
            updateDate = Format(ThisWorkbook.Sheets("検索").Cells(2, 1).value, "yyyymmdd")
            If ThisWorkbook.Sheets("検索").Cells(11, 2).value <> "" Then
                updateContent = "_" & ThisWorkbook.Sheets("検索").Cells(11, 2).value & "(+" & ThisWorkbook.Sheets("検索").Cells(11, 3).value & "hr)"
            Else
                updateContent = "_(-hr)"
            End If
            If ThisWorkbook.Sheets("検索").Cells(12, 2).value <> "" Then
                updateContent = updateContent & ThisWorkbook.Sheets("検索").Cells(12, 2).value & "(+" & ThisWorkbook.Sheets("検索").Cells(12, 3).value & "hr)"
            End If
            KupdateContent = "_" & ThisWorkbook.Sheets("所属変更").Cells(3, 2).value

            For i = 1 To 9
                Select Case True
                    Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 3).value <> "" And ThisWorkbook.Sheets("所属変更").Cells(i + 2, 4).value <> ""
                        KupdateContent = KupdateContent & ThisWorkbook.Sheets("所属変更").Cells(i + 2, 2).value & "(±常)"
                    Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 3).value <> ""
                        KupdateContent = KupdateContent & ThisWorkbook.Sheets("所属変更").Cells(i + 2, 2).value & "(+常)"
                    Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 4).value <> ""
                        KupdateContent = KupdateContent & ThisWorkbook.Sheets("所属変更").Cells(i + 2, 2).value & "(-常)"
                End Select
            Next i

        Case "admin"
            saveName = "【厚生局・保健所・振興局・労働局】管理薬剤師変更"
            strName = ThisWorkbook.Sheets("検索").Cells(2, 2).value & Format(ThisWorkbook.Sheets("検索").Cells(19, 3).value, "0000")
            updateDate = Format(ThisWorkbook.Sheets("検索").Cells(2, 1).value, "yyyymmdd")
            updateContent = "_" & ThisWorkbook.Sheets("検索").Cells(7, 1).value & "→" & ThisWorkbook.Sheets("検索").Cells(9, 1).value
            KupdateContent = "_" & ThisWorkbook.Sheets("所属変更").Cells(3, 2).value

            For i = 1 To 9
                Select Case True
                    Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 3).value <> "" And ThisWorkbook.Sheets("所属変更").Cells(i + 2, 4).value <> ""
                        KupdateContent = KupdateContent & "(±常)"
                    Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 3).value <> ""
                        KupdateContent = KupdateContent & "(+常)"
                    Case ThisWorkbook.Sheets("所属変更").Cells(i + 2, 4).value <> ""
                        KupdateContent = KupdateContent & "(-常)"
                End Select
            Next i
    End Select

    ' Create PDFs folder if it doesn't exist
    pdfFolder = ThisWorkbook.Path & Application.PathSeparator & "PDFs"

    ' Set full path for PDF file
    On Error GoTo ErrorHandler
    pName = pdfFolder & Application.PathSeparator & updateDate & strName & saveName & updateContent & d & ".pdf"

    Debug.Print pName

    ' Log information based on file type
    Set RWs = ThisWorkbook.Sheets("作成書類リネーム用")
    Select Case fileName
        Case "government"
            RWs.Cells(RWs.Cells(1000, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【厚生局】異動届" & updateContent
        Case "pharmacy"
            RWs.Cells(RWs.Cells(1000, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【厚生局】異動届" & KupdateContent
            RWs.Cells(RWs.Cells(1000, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【保健所】その他薬剤師変更届" & updateContent
        Case "admin"
            RWs.Cells(RWs.Cells(1000, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【厚生局】異動届" & updateContent & KupdateContent
            RWs.Cells(RWs.Cells(1000, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【保健所】管理薬剤師変更届" & updateContent
            RWs.Cells(RWs.Cells(1000, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【保健所】高度管理機器管理者変更届" & updateContent
            RWs.Cells(RWs.Cells(1000, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【保健所】自立支援(育生更生)管理薬剤師変更届" & updateContent
            RWs.Cells(RWs.Cells(1000, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【振興局】自立支援(精神通院)管理薬剤師変更届" & updateContent
            RWs.Cells(RWs.Cells(1000, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " 【労働局】管理薬剤師変更届" & updateContent
        End Select
        ' Call the function to create PDF
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
    
    If Kws.Cells(3, 5).value = "常勤" Then
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("常勤").Visible = True
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("非常勤").Visible = False
    ElseIf Kws.Cells(3, 5).value = "非常勤" Then
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("常勤").Visible = False
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("非常勤").Visible = True
    Else
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("常勤").Visible = False
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("非常勤").Visible = False
    End If

    If Kws.Cells(3, 1).value = "転入" Then
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("転入").Visible = True
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("転出").Visible = False
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("入薬").Visible = True
        ThisWorkbook.Worksheets("新<厚>異動届").Shapes("出薬").Visible = False
    ElseIf Kws.Cells(3, 1).value = "転出" Then
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
        If Kws.Cells(j + 3, 1).value = "転入" Then
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("転入" & j).Visible = True
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("転出" & j).Visible = False
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("入薬" & j).Visible = True
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("出薬" & j).Visible = False
        ElseIf Kws.Cells(j + 3, 1).value = "転出" Then
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

        If Kws.Cells(j + 3, 5).value = "常勤" Then
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("常勤" & j).Visible = True
            ThisWorkbook.Worksheets("新<厚>別紙").Shapes("非常勤" & j).Visible = False
        ElseIf Kws.Cells(j + 3, 5).value = "非常勤" Then
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

