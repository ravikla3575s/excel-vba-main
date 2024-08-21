Attribute VB_Name = "Module4"
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
            saveName = "�y�����ǁz�ٓ���"
            strName = ThisWorkbook.Sheets("�����ύX").Cells(2, 1).value & Format(ThisWorkbook.Sheets("�����ύX").Cells(19, 2).value, "0000")
            updateDate = Format(ThisWorkbook.Sheets("�����ύX").Cells(3, 3).value, "yyyymmdd")
            updateContent = "_" & ThisWorkbook.Sheets("�����ύX").Cells(3, 2).value

            For i = 1 To 9
                If ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 2).value <> "" Then
                    updateContent = updateContent & ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 2).value
                    Select Case True
                        Case ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 3).value <> "" And ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 4).value <> ""
                            updateContent = updateContent & "(�}��)"
                        Case ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 3).value <> ""
                            updateContent = updateContent & "(+��)"
                        Case ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 4).value <> ""
                            updateContent = updateContent & "(-��)"
                    End Select
                Else
                    Exit For
                End If
            Next i

        Case "pharmacy"
            saveName = "�y�ی����z���̑���܎t�ύX"
            strName = ThisWorkbook.Sheets("����").Cells(2, 2).value & Format(ThisWorkbook.Sheets("����").Cells(19, 3).value, "0000")
            updateDate = Format(ThisWorkbook.Sheets("����").Cells(2, 1).value, "yyyymmdd")
            If ThisWorkbook.Sheets("����").Cells(11, 2).value <> "" Then
                updateContent = "_" & ThisWorkbook.Sheets("����").Cells(11, 2).value & "(+" & ThisWorkbook.Sheets("����").Cells(11, 3).value & "hr)"
            Else
                updateContent = "_(-hr)"
            End If
            If ThisWorkbook.Sheets("����").Cells(12, 2).value <> "" Then
                updateContent = updateContent & ThisWorkbook.Sheets("����").Cells(12, 2).value & "(+" & ThisWorkbook.Sheets("����").Cells(12, 3).value & "hr)"
            End If
            KupdateContent = "_" & ThisWorkbook.Sheets("�����ύX").Cells(3, 2).value

            For i = 1 To 9
                Select Case True
                    Case ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 3).value <> "" And ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 4).value <> ""
                        KupdateContent = KupdateContent & ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 2).value & "(�}��)"
                    Case ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 3).value <> ""
                        KupdateContent = KupdateContent & ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 2).value & "(+��)"
                    Case ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 4).value <> ""
                        KupdateContent = KupdateContent & ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 2).value & "(-��)"
                End Select
            Next i

        Case "admin"
            saveName = "�y�����ǁE�ی����E�U���ǁE�J���ǁz�Ǘ���܎t�ύX"
            strName = ThisWorkbook.Sheets("����").Cells(2, 2).value & Format(ThisWorkbook.Sheets("����").Cells(19, 3).value, "0000")
            updateDate = Format(ThisWorkbook.Sheets("����").Cells(2, 1).value, "yyyymmdd")
            updateContent = "_" & ThisWorkbook.Sheets("����").Cells(7, 1).value & "��" & ThisWorkbook.Sheets("����").Cells(9, 1).value
            KupdateContent = "_" & ThisWorkbook.Sheets("�����ύX").Cells(3, 2).value

            For i = 1 To 9
                Select Case True
                    Case ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 3).value <> "" And ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 4).value <> ""
                        KupdateContent = KupdateContent & "(�}��)"
                    Case ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 3).value <> ""
                        KupdateContent = KupdateContent & "(+��)"
                    Case ThisWorkbook.Sheets("�����ύX").Cells(i + 2, 4).value <> ""
                        KupdateContent = KupdateContent & "(-��)"
                End Select
            Next i
    End Select

    ' Create PDFs folder if it doesn't exist
    pdfFolder = ThisWorkbook.Path & Application.PathSeparator & "PDFs"
    If Dir(pdfFolder, vbDirectory) = "" Then
        MkDir pdfFolder
    End If

    ' Set full path for PDF file
    On Error GoTo ErrorHandler
    pName = pdfFolder & Application.PathSeparator & updateDate & strName & saveName & updateContent & d & ".pdf"

    Debug.Print pName

    ' Log information based on file type
    Set RWs = ThisWorkbook.Sheets("�쐬���ރ��l�[���p")
    Select Case fileName
        Case "government"
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " �y�����ǁz�ٓ���" & updateContent
        Case "pharmacy"
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " �y�����ǁz�ٓ���" & KupdateContent
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " �y�ی����z���̑���܎t�ύX��" & updateContent
        Case "admin"
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " �y�����ǁz�ٓ���" & updateContent & KupdateContent
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " �y�ی����z�Ǘ���܎t�ύX��" & updateContent
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " �y�ی����z���x�Ǘ��@��Ǘ��ҕύX��" & updateContent
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " �y�ی����z�����x��(�琶�X��)�Ǘ���܎t�ύX��" & updateContent
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " �y�U���ǁz�����x��(���_�ʉ@)�Ǘ���܎t�ύX��" & updateContent
            RWs.Cells(RWs.Cells(100, 1).End(xlUp).Row + 1, 1) = updateDate & strName & " �y�J���ǁz�Ǘ���܎t�ύX��" & updateContent
        End Select
        ' Call the function to create PDF
        Call makePdfs(sheet, pName)
        Exit Sub
ErrorHandler:
MsgBox "�G���[���������܂���: " & Err.Description, vbCritical
End Sub
Sub Shapes(fileName As String)
    Dim Kws As Worksheet
    Dim j As Integer
    Dim sheetProtected1 As Boolean
    Dim sheetProtected2 As Boolean
    
    On Error GoTo ErrorHandler

    Set Kws = ThisWorkbook.Worksheets("�����ύX")
    
    ' �V�[�g�̕ی������
    sheetProtected1 = ThisWorkbook.Worksheets("�V<��>�ٓ���").ProtectContents
    If sheetProtected1 Then ThisWorkbook.Worksheets("�V<��>�ٓ���").Unprotect
    
    ' "�V<��>�ٓ���"�V�[�g�̃V�F�C�v��ݒ�
    ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("�ǖ�").Visible = False
    ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("�`�F�b�N1").Visible = False
    ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("�`�F�b�N2").Visible = False

    If fileName = "admin" Then
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("�ǖ�").Visible = True
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("�`�F�b�N1").Visible = True
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("�`�F�b�N2").Visible = True
    End If
    
    If Kws.Cells(3, 5).value = "���" Then
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("���").Visible = True
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("����").Visible = False
    ElseIf Kws.Cells(3, 5).value = "����" Then
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("���").Visible = False
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("����").Visible = True
    Else
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("���").Visible = False
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("����").Visible = False
    End If

    If Kws.Cells(3, 1).value = "�]��" Then
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("�]��").Visible = True
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("�]�o").Visible = False
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("����").Visible = True
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("�o��").Visible = False
    ElseIf Kws.Cells(3, 1).value = "�]�o" Then
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("�]��").Visible = False
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("�]�o").Visible = True
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("����").Visible = False
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("�o��").Visible = True
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("���").Visible = False
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("����").Visible = False
    Else
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("�]��").Visible = False
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("�]�o").Visible = False
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("����").Visible = False
        ThisWorkbook.Worksheets("�V<��>�ٓ���").Shapes("�o��").Visible = False
    End If
    
    ' �K�v�ɉ����ĕʂ̃V�[�g�̕ی������
    sheetProtected2 = ThisWorkbook.Worksheets("�V<��>�ʎ�").ProtectContents
    If sheetProtected2 Then ThisWorkbook.Worksheets("�V<��>�ʎ�").Unprotect
    
    ' "�V<��>�ʎ�"�V�[�g�̃V�F�C�v��ݒ�
    For j = 1 To 4
        If Kws.Cells(j + 3, 1).value = "�]��" Then
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("�]��" & j).Visible = True
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("�]�o" & j).Visible = False
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("����" & j).Visible = True
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("�o��" & j).Visible = False
        ElseIf Kws.Cells(j + 3, 1).value = "�]�o" Then
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("�]��" & j).Visible = False
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("�]�o" & j).Visible = True
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("����" & j).Visible = False
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("�o��" & j).Visible = True
        Else
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("�]��" & j).Visible = False
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("�]�o" & j).Visible = False
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("����" & j).Visible = False
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("�o��" & j).Visible = False
        End If

        If Kws.Cells(j + 3, 5).value = "���" Then
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("���" & j).Visible = True
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("����" & j).Visible = False
        ElseIf Kws.Cells(j + 3, 5).value = "����" Then
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("���" & j).Visible = False
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("����" & j).Visible = True
        Else
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("���" & j).Visible = False
            ThisWorkbook.Worksheets("�V<��>�ʎ�").Shapes("����" & j).Visible = False
        End If
    Next j

    ' �V�[�g�̕ی���ēx�L����
    If sheetProtected1 Then ThisWorkbook.Worksheets("�V<��>�ٓ���").Protect
    If sheetProtected2 Then ThisWorkbook.Worksheets("�V<��>�ʎ�").Protect

Exit Sub

ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description, vbExclamation
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
    MsgBox "PDF�̍쐬���ɃG���[���������܂���"
End Sub

