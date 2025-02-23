' 既存ファイルを探す（なければ作成）
Function FindOrCreateReport(folderPath As String, targetYear As String, targetMonth As String, templatePath As String) As String
    Dim fso As Object
    Dim filePath As String
    Dim fileName As String
    Dim newWb As Workbook

    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = "保険請求管理報告書_" & targetYear & targetMonth & ".xlsx"
    filePath = folderPath & "\" & fileName

    ' ファイルが存在するか
    If Not fso.FileExists(filePath) Then
        ' 新規作成
        Set newWb = Workbooks.Open(templatePath)
        newWb.SaveAs filePath, FileFormat:=xlOpenXMLWorkbook
        newWb.Close
    End If

    FindOrCreateReport = filePath
End Function