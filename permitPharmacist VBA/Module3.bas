Attribute VB_Name = "Module3"
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
    Dim storeName As String
    Dim nameParts() As String
    
    ' Excel�t�@�C���̂���t�H���_�p�X���w��
    folderPath = "/Users/yoshipc/Desktop/�ߘa6�N3�������҃��X�g/" ' <-- �t�H���_�p�X��K�X�ύX���Ă�������
    
    ' �����ύX�V�[�g�̐ݒ�
    Set Kws = ThisWorkbook.Sheets("�����ύX")
    
    ' �t�H���_���̍ŏ��̃t�@�C�����擾
    exfileName = Dir(folderPath & "*.xlsx")
    
    ' �t�H���_���̑S�Ẵt�@�C�������[�v
    On Error GoTo ErrLabel
    Do While exfileName <> ""
        ' �e�t�@�C�����J��
        Set wb = Workbooks.Open(folderPath & exfileName)
        
        ' A1�Z���̓X�ܖ����擾���A"�@" �܂��� " " �ŋ�؂��č����̕������X�ܖ��Ƃ���
        storeName = wb.Sheets(1).Range("A1").value
        nameParts = Split(storeName, "�@") ' �S�p�X�y�[�X�ŕ���
        
        If UBound(nameParts) = 0 Then
            nameParts = Split(storeName, " ") ' ���p�X�y�[�X�ōĕ���
        End If
        
        storeName = nameParts(0) ' �����̕�������擾
        
        ' �X�ܖ��̖�����"�X"���t���Ă��Ȃ��ꍇ�A"�X"��t��
        If Right(storeName, 1) <> "�X" Then
            storeName = storeName & "�X"
        End If
        
        ' �X�ܖ���A2�Z���ɃZ�b�g
        Kws.Cells(2, 1).value = storeName
        Kws.Range("E2").value = "����"
        Kws.Range("B3:D11").ClearContents
        
        ' �V���������������𔻒�
        If IsNewFormat(wb.Sheets(1)) Then
            ' �V�����̏ꍇ�̏���
            supporterData = GetSupporterDataFromSheet(wb.Sheets(1))
            
            ' �f�[�^�������ύX�V�[�g�ɔ��f
            For i = LBound(supporterData, 1) + 2 To UBound(supporterData, 1)
                supporterName = supporterData(i, 1)
                startDate = supporterData(i, 2)
                endDate = supporterData(i, 3)
                
                If supporterName = "" Then
                    Exit For
                End If
                
                ' ���t�𕶎��񂩂���t�^�ɕϊ�
                If Not IsDate(startDate) Then
                    startDate = CDate(startDate)
                End If
                If Not IsDate(endDate) Then
                    endDate = CDate(endDate)
                End If
                
                ' �����ύX�V�[�g�Ƀf�[�^���X�V
                UpdateSupporterInSheet Kws, supporterName, startDate, endDate
            Next i
        Else
            ' �������̏ꍇ�̏���
            supporterName = wb.Sheets(1).Range("C4").value
            Call SortAndIndexDates(wb.Sheets(1))
            startDate = wb.Sheets(1).Range("B4").value
            endDate = wb.Sheets(1).Range("B4").End(xlDown).value
            
            ' �����ύX�V�[�g�Ƀf�[�^���X�V
            UpdateSupporterInSheet Kws, supporterName, startDate, endDate
        End If
        
        ' B12�Z���`B16�Z�����ォ�猟�����A�l�������Ă���ꍇ��B13���珇�Ɉڂ����
        Call CopyValuesToThisWorkbook(wb.Sheets(1), Kws)
        
        ' �������I�������t�@�C�������
        wb.Close False
        
        ' PDF���쐬�i�K�v�ɉ����āj
        ThisWorkbook.Activate
        Call �����Ǐ����ύX����PDF
        
        ' ���̃t�@�C�����擾
        exfileName = Dir
    Loop
    
    Call UpdateMultiplePharmacists
    Kws.Range("B3:D11").ClearContents
    Kws.Range("B13:B17").ClearContents
    Exit Sub
ErrLabel:
    msg = "�G���[���������܂���"
    wb.Close False
    exfileName = Dir
End Sub

Function IsNewFormat(ws As Worksheet) As Boolean
    ' D1�Z���̒l���`�F�b�N���āA"���X�ܖ�����͂��Ă�������"�ł��邩�𔻒肵�܂�
    If ws.Range("D1").value = "���X�ܖ�����͂��Ă�������" Then
        IsNewFormat = True  ' �V����
    Else
        IsNewFormat = False ' ������
    End If
End Function

Sub SortAndIndexDates(ws As Worksheet)
    Dim rng As Range
    Dim cell As Range
    Dim dateArray() As Date
    Dim i As Long, j As Long
    Dim tempDate As Date
    
    ' ���t�����͂���Ă���͈́iB��j���w��
    Set rng = ws.Range("B4:B" & ws.Cells(ws.Rows.count, "B").End(xlUp).Row)
    
    ' ���t��z��Ɋi�[
    ReDim dateArray(1 To rng.Rows.count)
    i = 1
    For Each cell In rng
        If IsDate(cell.value) Then
            dateArray(i) = CDate(cell.value)
            i = i + 1
        End If
    Next cell
    
    ' ���t�z����\�[�g
    For i = LBound(dateArray) To UBound(dateArray) - 1
        For j = i + 1 To UBound(dateArray)
            If dateArray(i) > dateArray(j) Then
                tempDate = dateArray(i)
                dateArray(i) = dateArray(j)
                dateArray(j) = tempDate
            End If
        Next j
    Next i
    
    ' �\�[�g���ꂽ���t���V�[�g�ɔ��f
    i = 1
    For Each cell In rng
        cell.value = dateArray(i)
        i = i + 1
    Next cell
End Sub

Sub UpdateSupporterInSheet(Kws As Worksheet, Name As String, startDate As Variant, endDate As Variant)
    Dim lastRow As Long
    Dim found As Range
    
    ' ���O�����łɃV�[�g�ɂ��邩���m�F
    Set found = Kws.Columns("B").Find(Name, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' ���t�f�[�^��Date�^�łȂ��ꍇ�A�ϊ�
    If Not IsDate(startDate) Then
        startDate = CDate(startDate)
    End If
    If Not IsDate(endDate) Then
        endDate = CDate(endDate)
    End If
    
    If Not found Is Nothing Then
        ' ���O�����������ꍇ�A�ŏI���t���X�V
        If found.Offset(0, 2).value < endDate Then
            found.Offset(0, 2).value = endDate
        End If
    Else
        ' �V�����s�ɒǉ�
        lastRow = Kws.Cells(11, "B").End(xlUp).Row + 1
        Kws.Cells(lastRow, 2).value = Name
        Kws.Cells(lastRow, 3).value = startDate
        Kws.Cells(lastRow, 4).value = endDate
    End If
End Sub

Sub CopyValuesToThisWorkbook(srcWs As Worksheet, destWs As Worksheet)
    Dim i As Long
    For i = 12 To 16
        If srcWs.Cells(i, 2).value <> "" Then
            destWs.Cells(i - 11 + 12, 2).value = srcWs.Cells(i, 2).value
        End If
    Next i
End Sub

Function GetSupporterDataFromSheet(ws As Worksheet) As Variant
    Dim lastRow As Long
    Dim dataRange As Range
    
    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    
    ' �f�[�^�͈͂�ݒ�
    Set dataRange = ws.Range("B2:D" & lastRow)
    
    ' �f�[�^��z��ɕϊ����ĕԂ�
    GetSupporterDataFromSheet = dataRange.value
End Function

Sub ExampleDir()
    Dim folderPath As String
    Dim fileName As String
    
    ' ��������t�H���_���w��
    folderPath = "/Users/yoshipc/Desktop/�ߘa6�N3�������҃��X�g/"
    
    ' �ŏ��̃t�@�C�����擾
    fileName = Dir(folderPath & "*.xlsx")
    
    ' ���[�v�Ńt�H���_���̂��ׂĂ� .xlsx �t�@�C�����擾
    Do While fileName <> ""
        ' �t�@�C�������o��
        Debug.Print "Found file: " & fileName
        
        ' ���̃t�@�C�����擾
        fileName = Dir
    Loop
End Sub

