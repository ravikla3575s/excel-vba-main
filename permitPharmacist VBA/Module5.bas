Attribute VB_Name = "Module5"
Sub UpdatePharmacistInfoWithClass()
    Dim ws As Worksheet, Kws As Worksheet
    Set ws = ThisWorkbook.Sheets("�͏o�ꗗ�e�[�u��")
    Set Kws = ThisWorkbook.Sheets("�����ύX")
    Dim strName As String
    Dim updateRow As Long
    Dim startColumn As Long
    Dim pharmacists() As Class1
    Dim i As Long, j As Long, count As Long
    
    '�X����strName�ɓ����
    strName = Kws.Cells(2, 1)
    
    ' findUpdateRow�֐��̌Ăяo��
    updateRow = findUpdateRow(ws, strName)
    
    ' �J�n�J���� (�J������140���)
    For startColumn = 1 To 215
    If ws.Cells(1, startColumn) = "���Ζ�܎t10" Then
        startColumn = startColumn + 1
        Exit For
    End If
    Next startColumn
    
    ' 10�l���̃f�[�^���N���X�Ɋi�[
    ReDim pharmacists(1 To 10)
    
    For i = 1 To 10
        Set pharmacists(i) = New Class1
        
        ' EmployeeNumber�̏���
        empNum = ws.Cells(updateRow, startColumn + (i - 1) * 3).value
        If IsNumeric(empNum) And Len(empNum) <= 7 Then
            pharmacists(i).EmployeeNumber = CLng(empNum)
        Else
            pharmacists(i).EmployeeNumber = 0 ' �����ȏꍇ��0�ɐݒ�
        End If
        
        ' PharmacistName�̏���
        pharmacists(i).PharmacistName = ws.Cells(updateRow, startColumn + (i - 1) * 3 + 1).value
        
        ' WorkHour�̏���
        workHr = ws.Cells(updateRow, startColumn + (i - 1) * 3 + 2).value
        If IsNumeric(workHr) Then
            pharmacists(i).WorkHour = CSng(workHr)
        Else
            pharmacists(i).WorkHour = 0 ' �����ȏꍇ��0�ɐݒ�
        End If
    Next i
    
    ' ��̃C���X�^���X���l�߂鏈��
    count = 0
    For i = 1 To 10
        If pharmacists(i).EmployeeNumber <> 0 Or Len(pharmacists(i).PharmacistName) > 0 Or pharmacists(i).WorkHour <> 0 Then
            count = count + 1
            If count <> i Then
                Set pharmacists(count) = pharmacists(i)
            End If
        End If
    Next i
    
    ' �������ꂽ�z����ēx�V�[�g�ɏ�������
    For i = 1 To count
        ws.Cells(updateRow, startColumn + (i - 1) * 3).value = pharmacists(i).EmployeeNumber
        ws.Cells(updateRow, startColumn + (i - 1) * 3 + 1).value = pharmacists(i).PharmacistName
        ws.Cells(updateRow, startColumn + (i - 1) * 3 + 2).value = pharmacists(i).WorkHour
    Next i
    
    ' �c��̃Z�����N���A����
    For i = count + 1 To 10
        ws.Cells(updateRow, startColumn + (i - 1) * 3).ClearContents
        ws.Cells(updateRow, startColumn + (i - 1) * 3 + 1).ClearContents
        ws.Cells(updateRow, startColumn + (i - 1) * 3 + 2).ClearContents
    Next i

    Dim fullTimePharmacists() As String
    Dim partTimePharmacists() As String
    Dim fullTimeCount As Long
    Dim partTimeCount As Long
    Dim col As Long

    ' ���/���΂̕��ޗp�z���������
    ReDim fullTimePharmacists(1 To 10)
    ReDim partTimePharmacists(1 To 5)
    fullTimeCount = 0
    partTimeCount = 0

    ' ���/���΂𕪗�
    For i = 1 To count
        If pharmacists(i).WorkHour > 32 Then
            fullTimeCount = fullTimeCount + 1
            fullTimePharmacists(fullTimeCount) = pharmacists(i).PharmacistName
        Else
            partTimeCount = partTimeCount + 1
            partTimePharmacists(partTimeCount) = pharmacists(i).PharmacistName
        End If
    Next i

    ' ��Ζ�܎t�̓o�^
    col = findColumn(ws, "��Ζ�܎t1")
    If col > 0 Then
        For i = 1 To fullTimeCount
            ws.Cells(updateRow, col + (i - 1)).value = fullTimePharmacists(i)
        Next i
        ' �c��̃Z�����N���A����
        For i = fullTimeCount + 1 To 10
            ws.Cells(updateRow, col + (i - 1)).ClearContents
        Next i
    End If

    ' ���Ζ�܎t�̓o�^
    col = findColumn(ws, "���Ζ�܎t1")
    If col > 0 Then
        For i = 1 To partTimeCount
            ws.Cells(updateRow, col + (i - 1)).value = partTimePharmacists(i)
        Next i
        ' �c��̃Z�����N���A����
        For i = partTimeCount + 1 To 5
            ws.Cells(updateRow, col + (i - 1)).ClearContents
        Next i
    End If

End Sub

' �J������������֐�
Function findColumn(ws As Worksheet, headerName As String) As Long
    Dim lastColumn As Long
    Dim i As Long
    
    lastColumn = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    For i = 1 To lastColumn
        If ws.Cells(1, i).value = headerName Then
            findColumn = i
            Exit Function
        End If
    Next i
    
    findColumn = 0 ' ������Ȃ������ꍇ
End Function
Function findUpdateRow(ws As Worksheet, strName As String) As Long
    Dim lastRow As Long
    Dim i As Long
    Dim storeName As String
    storeName = strName ' �����ɓX�ܖ���ݒ�
    
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    
    For i = 1 To lastRow
        If ws.Cells(i, 2).value = storeName Then
            findUpdateRow = i
            Exit Function
        End If
    Next i
    
    findUpdateRow = 0 ' �X�ܖ���������Ȃ������ꍇ
End Function
