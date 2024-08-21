Attribute VB_Name = "Module2"
Sub UpdateMultiplePharmacists()
    Dim i As Integer
    Dim PharmacistName As String, strName As String
    Dim updWS As Worksheet
    Dim Kws As Worksheet
    Dim updateRow As Integer
    Dim startColumn As Integer
    
    Set updWS = Worksheets("�͏o�ꗗ�e�[�u��")
    Set Kws = Worksheets("�����ύX")
    
    strName = Kws.Cells(2, 1).value
    
    ' ��J�������w��i�Ⴆ�΁A120��ڂ���Ƃ���j
    startColumn = 120
    
    ' B13:B17 �͈̔͂����ɏ���
    For i = 13 To 17
        PharmacistName = Kws.Cells(i, 2).value
        
        If PharmacistName = "" Or "0" Then
            Exit For
        End If
        
        ' �X�V����s��T��
        updateRow = findUpdateRow(updWS, strName)
        
        ' ��܎t�̓o�^�̗L�����m�F
        If IsPharmacistRegistered(updWS, PharmacistName, updateRow, startColumn) Then
            MsgBox PharmacistName & "�͊��ɓo�^����Ă��܂��B"
        Else
            ' �X�V������T���āA�l���X�V
            Call UpdatePharmacistAssignmentWithParams(updWS, PharmacistName, updateRow)
        End If
    Next i
End Sub

Function IsPharmacistRegistered(updWS As Worksheet, PharmacistName As String, updateRow As Integer, startColumn As Integer) As Boolean
    Dim i As Integer
    
    ' ��J��������E�Ɍ�������20���̃Z�����`�F�b�N
    On Error GoTo ErrLabel
    For i = 0 To 19
        If updWS.Cells(updateRow, startColumn + i).value = PharmacistName Then
            IsPharmacistRegistered = True
            Exit Function
        End If
    Next i
    
    ' ������Ȃ������ꍇ�AFalse��Ԃ�
    IsPharmacistRegistered = False
    Exit Function
      
ErrLabel:
    msg = "�G���[���������܂���"
        IsPharmacistRegistered = False
    Resume Next
End Function

Sub UpdatePharmacistAssignmentWithParams(updWS As Worksheet, PharmacistName As String, updateRow As Integer)
    Dim updateColumn As Integer
    
    ' �X�V������T��
    updateColumn = findUpdateColumn(updateRow)
    
    If updateColumn > 0 Then
        updWS.Cells(updateRow, updateColumn).value = PharmacistName
    End If
End Sub

Function findUpdateRow(updWS As Worksheet, strName As String) As Integer
    Dim updateRow As Integer
    
    ' �w�肳�ꂽ�X�ܖ�������s������
    For updateRow = 2 To 70
        If updWS.Cells(updateRow, 2).value = strName Then
            findUpdateRow = updateRow
            Exit Function
        End If
    Next updateRow
    
    ' ������Ȃ������ꍇ�A0��Ԃ�
    findUpdateRow = 0
End Function

Function findUpdateColumn(updateRow As Integer) As Integer
    Dim updateColumn As Integer
    Dim updWS As Worksheet
    Dim itemNames As Variant
    Dim i As Integer
    
    Set updWS = Worksheets("�͏o�ꗗ�e�[�u��")
    
    ' ���Ζ�܎t�̃X���b�g���`
    itemNames = Array("���Ζ�܎t6", "���Ζ�܎t7", "���Ζ�܎t8", "���Ζ�܎t9", "���Ζ�܎t10")
    
    ' itemNames �z������ɏ������A�󗓂̗��T��
    For i = LBound(itemNames) To UBound(itemNames)
        For updateColumn = 121 To 188
            If updWS.Cells(1, updateColumn).value = itemNames(i) Then
                ' �Y������Z�����󗓂����`�F�b�N
                If IsEmpty(updWS.Cells(updateRow, updateColumn)) Then
                    findUpdateColumn = updateColumn
                    Exit Function
                End If
                Exit For
            End If
        Next updateColumn
    Next i
    
    ' ������Ȃ������ꍇ�A0��Ԃ�
    findUpdateColumn = 0
End Function
