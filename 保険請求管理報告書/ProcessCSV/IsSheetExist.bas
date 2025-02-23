' シートが存在するか確認
Function IsSheetExist(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    IsSheetExist = False
    For Each ws In wb.Sheets
        If ws.Name = sheetName Then
            IsSheetExist = True
            Exit Function
        End If
    Next ws
End Function