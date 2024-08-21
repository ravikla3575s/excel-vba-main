Attribute VB_Name = "Module1"
Sub ��ʖ�܎t�ٓ����K�v���ވ��()
    Dim n As Integer, chkRow As Integer
    Dim upDate As Date
    Dim orgValue As String, chgValue As String
    Dim ws As Worksheet, cover As Worksheet, Kws As Worksheet
    Dim sheetArray() As String
    Dim sheetCount As Integer

    Set ws = ThisWorkbook.Worksheets("����")
    Set Kws = ThisWorkbook.Worksheets("�����ύX")
    
    If Application.Caller = "���̑���܎t�ύX" Then
        Kws.Cells(2, 5).value = "���"
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
                Set cover = ThisWorkbook.Worksheets("<��>�ύX��(���̑���܎t)")
                ws.Cells(10, 3).value = ws.Cells(chkRow, 11).value
                ws.Cells(11, 2).value = ws.Cells(chkRow, 9).value
            Else
                ws.Cells(12, 2).value = ws.Cells(chkRow, 9).value
            End If
            If ws.Cells(11, 2).value = "" And ws.Cells(chkRow, 11).value = "��܎t" Then
                Set cover = ThisWorkbook.Worksheets("<��>�ύX��(���̑���܎t)")
                ws.Cells(10, 3).value = ws.Cells(chkRow, 7).value
            ElseIf ws.Cells(11, 2).value = "" And ws.Cells(chkRow, 11).value = "�o�^�̔���" Then
                Set cover = ThisWorkbook.Worksheets("<��>�ύX��(���̑��o�^�̔���)")
                ws.Cells(10, 3).value = ws.Cells(chkRow, 7).value
            End If
        ElseIf ws.Cells(11, 2).value = "" And orgValue <> chgValue And chgValue = "" Then
            Set cover = ThisWorkbook.Worksheets("<��>�ύX��(���̑���܎t)")
            ws.Cells(10, 3).value = ws.Cells(chkRow, 7).value
        ElseIf ws.Cells(11, 2).value = "" And orgValue = chgValue And ws.Cells(chkRow, 6).value <> ws.Cells(chkRow, 10).value Then
            Set cover = ThisWorkbook.Worksheets("<��>�ύX��(���̑���܎t)")
            ws.Cells(10, 3).value = "��܎t�Ζ�����"
        ElseIf orgValue = "" And chgValue = "" Then
            Exit For
        End If
    Next chkRow

    'On Error GoTo ErrorHandler
    ' Add sheets to the array
    sheetArray(sheetCount) = cover.Name
    sheetCount = sheetCount + 1
    ReDim Preserve sheetArray(1 To sheetCount)
    sheetArray(sheetCount) = "<��>�ύX�͕ʎ�1"
    
    If ws.Cells(11, 2).value <> "" And ws.Cells(10, 3).value = "��܎t" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "�ٗp�ؖ����i���̑��j"
    ElseIf ws.Cells(11, 2).value <> "" And ws.Cells(10, 3).value = "�o�^�̔���" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "�ٗp�ؖ����i�o�^�̔��ҁj"
    End If
    
    If ws.Cells(12, 2).value <> "" And ws.Cells(10, 3).value = "��܎t" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "�ٗp�ؖ����i���̑��j (2)"
    End If
    
    If ws.Cells(8, 9).value <> "" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "<��>�ύX�͕ʎ�2"
    End If
    
    If ws.Cells(4, 2).value > upDate Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "�x�����R��"
    End If
    
    ' Prompt for �����ǈٓ��� creation
    Dim rc As VbMsgBoxResult
    rc = MsgBox("�����ǈٓ��͂ɕύX�������܂����H", vbYesNo + vbQuestion)
    If rc = vbYes Then
        Call �����ǈٓ��͍쐬
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "�V<��>�ٓ���"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "�V<��>�ʎ�"
    End If
    
    ' Select all sheets in the array and export to PDF
    Sheets(sheetArray).Select
    Call PrintToPdf(ActiveSheet, "pharmacy")
    Call �ύX���O�L�^
    
    ws.Cells(10, 3).ClearContents
    Kws.Cells(15, 2).ClearContents
    Exit Sub

ErrorHandler:
    MsgBox " �ύX�ӏ����݂���܂���"
End Sub
Sub �����ǈٓ��͍쐬()
    Dim chkRow As Integer, KRow As Integer
    Dim upDate As Date
    Dim orgValue As String, chgValue As String
    Dim ws As Worksheet, cover As Worksheet, Kws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("����")
    Set Kws = ThisWorkbook.Worksheets("�����ύX")
    
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
        
        If orgValue <> chgValue And orgValue = "" Then ' �l����������ꍇ
            If Kws.Cells(KRow - 1, 2).value = ws.Cells(chkRow, 9).value Then '�Ώێ҂����łɕ\��ɋL�ڂ���Ă���ꍇ
                ' do nothing
            Else
                Kws.Cells(KRow, 2).value = ws.Cells(chkRow, 9).value
                If Kws.Cells(16, 2).value <> "" Then '�V�l��܎t�̏ꍇ
                    Kws.Cells(KRow, 3).value = Kws.Cells(16, 2).value
                Else
                    Kws.Cells(KRow, 3).value = ws.Cells(2, 1).value
                End If
            End If

        ElseIf orgValue <> chgValue And chgValue = "" Then '�l��������ꍇ
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
Sub �Ǘ���܎t�ύX()
    Dim i As Integer, n As Integer, j As Integer, a As Integer, anyChanges As Boolean
    Dim ws As Worksheet, Kws As Worksheet
    Dim upDate As Date, changeDate As Date
    Dim t As String, F As String
    Dim rc As VbMsgBoxResult
    Dim sheetArray() As String
    Dim sheetCount As Integer
    
    t = "���"
    F = "����"
    anyChanges = True

    ThisWorkbook.Worksheets("�x�����R��").Cells(10, 17).value = "�Ǘ���܎t"
    ThisWorkbook.Worksheets("�x�����R��").Cells(18, 1).value = "�ی�����"

    Set ws = Worksheets("����")
    Set Kws = Worksheets("�����ύX")
    
    If Kws.Cells(2, 1).value = ws.Cells(2, 2).value Then
        rc = MsgBox("���͓��e����U���Z�b�g���܂����H", vbYesNo + vbQuestion)
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
        
        rc = MsgBox("���Ǘ��҂̈ٓ��͂���܂����H", vbYesNo + vbQuestion)

        If rc = vbYes Then
            cRow = Kws.Cells(12, 2).End(xlUp).Row + 1
            Kws.Cells(cRow, 2).value = ws.Range("E2").value
            Kws.Cells(cRow, 4).value = DateAdd("d", -1, ws.Range("A2").value)
        End If

        Call Shapes("admin")
        
        ' Initialize the array and counter
        sheetCount = 1
        ReDim sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "�V<��>�ٓ���"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "�V<��>�ʎ�"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "<��>�ύX�� (�Ǘ���)"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "<��>�ύX�͕ʎ�1"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "�ٗp�ؖ����i�Ǘ��ҁj"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "<��>���x�Ǘ��i�Ǘ��ҁj"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "�ٗp�ؖ����i���x�Ǘ��j"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "<�U>�����x���i���_�j�ύX�́i������܂ށj"
        
        If ws.Cells(248, 3).value = "����s" Then
            sheetCount = sheetCount + 1
            ReDim Preserve sheetArray(1 To sheetCount)
            sheetArray(sheetCount) = "<��>����s�E�����x���i�琬�X���j�ύX��"

            If ws.Cells(71, 3).value <> "" Then
                sheetCount = sheetCount + 1
                ReDim Preserve sheetArray(1 To sheetCount)
                sheetArray(sheetCount) = "<��>�Ō����戵�ӔC��"
                sheetCount = sheetCount + 1
                ReDim Preserve sheetArray(1 To sheetCount)
                sheetArray(sheetCount) = "�ٗp�ؖ����i�Ō����j"
            End If
        Else
            sheetCount = sheetCount + 1
            ReDim Preserve sheetArray(1 To sheetCount)
            sheetArray(sheetCount) = "<��>�����x���i�琬�X���j�ύX��(����ȊO�E�k�C��)"

            If ws.Cells(71, 3).value <> "" Then
                sheetCount = sheetCount + 1
                ReDim Preserve sheetArray(1 To sheetCount)
                sheetArray(sheetCount) = "<�k>�Ō����戵�ӔC��"
                sheetCount = sheetCount + 1
                ReDim Preserve sheetArray(1 To sheetCount)
                sheetArray(sheetCount) = "�ٗp�ؖ����i�Ō����j"
            End If
        End If
        
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "<�J>�Ǘ���܎t�ύX��"
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "���F��o�󋵊m�F�["
        
        ' Select all sheets in the array and export to PDF
        Sheets(sheetArray).Select
        Call PrintToPdf(ActiveSheet, "admin")
        
        Call �ύX���O�L�^
        Call FindMatchingStrings
    Else
        MsgBox "A9 �ɕύX��̊Ǘ���܎t������͂��Ă�������"
    End If
End Sub
Sub �����Ǐ����ύX����PDF()
    Dim rc As VbMsgBoxResult
    Dim j As Integer
    Dim upDate As Date
    Dim Kws As Worksheet
    Dim ArrayShName() As String
    Dim sheetCount As Integer
    
    Set Kws = ThisWorkbook.Worksheets("�����ύX")
    
    If Application.Caller = "�����ғo�^" Then
        Kws.Cells(2, 5).value = "����"
    End If
    
    If ThisWorkbook.Worksheets("����").Cells(9, 1).value <> "" Then
        rc = MsgBox("�����ǈٓ��Ɋւ��āA�Ǘ��ҕύX�������ɍs���܂����H", vbYesNo + vbQuestion)
        If rc = vbNo Then
            ThisWorkbook.Worksheets("����").Cells(9, 1).ClearContents
        End If
    End If
    
    ' Initialize the array and counter
    sheetCount = 1
    ReDim sheetArray(1 To sheetCount)
    sheetArray(sheetCount) = "�V<��>�ٓ���"
    
    If Kws.Cells(4, 2).value <> "" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "�V<��>�ʎ�"
    End If
    
    If Kws.Cells(8, 2).value <> "" Then
        sheetCount = sheetCount + 1
        ReDim Preserve sheetArray(1 To sheetCount)
        sheetArray(sheetCount) = "�V<��>�ʎ� (2)"
    End If
    
    Call Shapes("pharmacy")
    ' Select all sheets in the array and export to PDF
    Sheets(sheetArray).Select
    Call PrintToPdf(ActiveSheet, "government")
    
    ThisWorkbook.Worksheets("�����ύX").Select
End Sub
Sub �ύX���O�L�^()
    Dim strName As String, logNo As String
    Dim logRow As Long, updRow As Long
    Dim ws As Worksheet, updWS As Worksheet, logWS As Worksheet

    Set ws = ThisWorkbook.Worksheets("����")
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
Sub �ǖ򏊑��ύX�]�L()
    Dim i As Integer, trsRow As Integer, n As Integer, trsColumn As Integer, c As Integer
    Dim wkHour As Single
    Dim updName As String, updNumber As String, strName As String
    Dim ws As Worksheet, updWS As Worksheet
    Dim contentsArray() As Variant
    
    contentsArray = Array("�ی����Ǘ���܎t", "���x�Ǘ���Ë@�퓙�̔��Ǘ���", "�����x���i���_�ʉ@��Áj�S����", "�����x���i�X���琬��Áj�S����", "�����ی�@�Ǘ���܎t", "�J�Еی��i�Ǘ���܎t�j", "��ǋ@�\���i�Ǘ���܎t�j", "���^�����Ǘ���", "ADHD���ʊǗ� �ӔC��", "�����ǊǗ���܎t", "�Ō����戵�ӔC��")
    
    Set ws = ThisWorkbook.Worksheets("����")
    Set Kws = ThisWorkbook.Worksheets("�����ύX")
    Set updWS = ThisWorkbook.Worksheets("�͏o�ꗗ�e�[�u��")
    
    If Application.Caller = "�Ǘ���܎t�X�V" Then
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

Sub �����ύX�]�L()
    Dim updRow As Integer, trsRow As Integer, n As Integer, startColumn As Integer
    Dim wkHour As Variant
    Dim updName As String, updNumber As String, strName As String
    Dim ws As Worksheet, updWS As Worksheet, Kws As Worksheet
    

    Set ws = ThisWorkbook.Worksheets("����")
    Set updWS = ThisWorkbook.Worksheets("�͏o�ꗗ�e�[�u��")
    Set Kws = ThisWorkbook.Worksheets("�����ύX")

    strName = ws.Cells(2, 2)
    
    ' ��J�����i�u���Ζ�܎t10�v�j�̌���
    For startColumn = 1 To updWS.Columns.count
        If updWS.Cells(1, startColumn).value = "���Ζ�܎t10" Then
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
    Call ������(ws)
    Call UpdatePharmacistInfoWithClass
End Sub
Sub ������(ws As Worksheet)
    ws.Range("E2").Formula = "=IF($C$4="""",C138,VLOOKUP(����!$C$4,LOG!$A$2:$Y$1048507,3,0))"
    ws.Range("F2").Formula = "=VLOOKUP(E2,��܎t�}�X�^!C:V,17,0)"
    ws.Range("E3:E12").Formula = "=IF($C$4="""",INDIRECT(""C""&ROW($A$171)+ROW()*3),VLOOKUP(����!$C$4,LOG!$A$1:$Y$1048507,ROW()*2-1,0))"
    ws.Range("F3:F12").Formula = "=IFERROR(IF($C$4="""",IFERROR(INDIRECT(""C""&ROW($A$172)+ROW()*3),""""),IF(E3<>"""",VLOOKUP(����!$C$4,LOG!$A$1:$Y$1048507,ROW()*2,0),"""")),J3)"
    ws.Range("I2").Formula = "=IF(A9<>"""",A9,E2)"
    'ws.Range("J2").Formula = "=DGET(��܎t�}�X�^[#���ׂ�],""�T�J������"",����!I1:I2)"
    ws.Range("I3:I12").Formula = "=INDIRECT(""C""&ROW($A$171)+ROW()*3)"
    ws.Range("J3:J12").Formula = "=INDIRECT(""C""&ROW($A$172)+ROW()*3)"
End Sub
Sub �����ǈٓ��͈���m�F()
    Dim rc As VbMsgBoxResult
    rc = MsgBox("�����ǈٓ��͂�������܂����H", vbYesNo + vbQuestion)

    If rc = vbYes Then
        Call �����Ǐ����ύX����PDF
    End If

End Sub
Sub FindMatchingStrings()
    Dim rng As Range
    Dim cell As Range
    Dim targetStr As String
    Dim ws As Worksheet
    Dim resultWs As Worksheet
    Dim resultRow As Integer

    Set ws = ThisWorkbook.Sheets("����") ' �V�[�g����K�؂ɐݒ肵�Ă�������
    
    ws.Range("L2:L30").Clear
    ' ������������������`���܂�
    targetStr = ws.Cells(16, 2)

    ' �����Ώۂ͈̔͂�ݒ肵�܂�
    Set rng = Sheets("�͏o�ꗗ�e�[�u��").Range("CU2:GX70")  ' B2:B100 �������Ώۂ͈̔͂ɒu�������Ă�������

    resultRow = 2 ' ���ʂ��o�͂���J�n�s�i�����ł�Sheet2��1�s�ڂ���J�n�j
    ' �e�Z�������[�v�Œ��ׂ܂�
    For Each cell In rng

    If cell.value = targetStr Then
        ' �}�b�`���镶���񂪌���������Ή�����A��̃Z���̒l���o�͂��܂�
        ws.Cells(resultRow, 12).value = Sheets("�͏o�ꗗ�e�[�u��").Range("B" & cell.Row).value
        resultRow = resultRow + 1 ' ���ʏo�͍s�����ɐi�߂܂�
    End If

    Next cell

End Sub

Sub ��܎t�ύX�}�N��()
    Dim strName As String, PharmacistName As String, apdateItem As String
    Dim updateRow As Integer, updateColumn As Integer
    Dim updWS As Worksheet, ws As Worksheet

    Set updWS = Worksheets("�͏o�ꗗ�e�[�u��")
    Set Kws = Worksheets("�����ύX")

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

    MsgBox "������܎t�̃f�[�^�x�[�X�X�V���������܂����I"
    Call �ύX���O�L�^
End Sub
Sub �V�X���F�o�^()
    Dim i As Integer, n As Integer, updColumn As Integer
    Dim updDate As Date
    Dim updContents As String
    Dim ws As Worksheet, updWS As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("����")
    Set updWS = ThisWorkbook.Worksheets("�͏o�ꗗ�e�[�u��")
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
