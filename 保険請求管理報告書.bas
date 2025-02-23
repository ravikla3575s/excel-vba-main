
Sub ProcessCSV()
    Dim csvFile As String
    Dim fileName As String
    Dim targetYear As String
    Dim targetMonth As String
    Dim receiptYear As Integer
    Dim receiptMonth As Integer
    Dim reiwaYear As Integer
    Dim sendMonth As Integer
    Dim sendDate As String
    Dim savePath As String
    Dim templatePath As String
    Dim newBook As Workbook
    Dim wsTemplate As Worksheet
    Dim wsSheet As Worksheet
    Dim sheetName As String
    Dim fileType As String
    Dim fso As Object

    ' CSVファイル選択
    csvFile = Application.GetOpenFilename("CSVファイル (*.csv), *.csv", , "CSVファイルを選択してください")
    If csvFile = "False" Then Exit Sub

    ' ファイル名取得
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = fso.GetBaseName(csvFile)

    ' 診療年月をファイル名から取得
    If Len(fileName) < 22 Then
        MsgBox "ファイル名の形式が正しくありません。", vbExclamation
        Exit Sub
    End If

    ' ファイル種類判定（fmei → 振込額明細書, fixf → 請求確定状況）
    If InStr(fileName, "fmei") > 0 Then
        fileType = "振込額明細書"
    ElseIf InStr(fileName, "fixf") > 0 Then
        fileType = "請求確定状況"
    Else
        MsgBox "ファイル名から種類を判定できません。", vbExclamation
        Exit Sub
    End If

    ' 令和年と振込月を取得（19~23桁目）
    reiwaYear = CInt(Mid(fileName, 19, 2))
    receiptMonth = CInt(Mid(fileName, 21, 2))

    ' 西暦に変換
    receiptYear = 2018 + reiwaYear

    ' 診療月（振込月の前月）
    If receiptMonth = 1 Then
        receiptMonth = 12
        receiptYear = receiptYear - 1
    Else
        receiptMonth = receiptMonth - 1
    End If

    ' 診療年月フォーマット
    targetYear = CStr(receiptYear)
    targetMonth = Format(receiptMonth, "00")
    Dim diagnosisPeriod As String
    diagnosisPeriod = targetYear & "年" & targetMonth & "月診療分"

    ' 送信月
    sendMonth = receiptMonth + 1
    If sendMonth = 13 Then sendMonth = 1
    sendDate = sendMonth & "月10日送信分"

    ' B2, B3セルからテンプレートと保存フォルダのパスを取得
    templatePath = ThisWorkbook.Sheets(1).Range("B2").Value & "\" & "保険請求管理報告書テンプレート20250222.xltm"
    savePath = ThisWorkbook.Sheets(1).Range("B3").Value
    If templatePath = "" Or savePath = "" Then
        MsgBox "テンプレートまたは保存フォルダが設定されていません。", vbExclamation
        Exit Sub
    End If

    ' 対象ファイルを取得または作成
    Dim targetFile As String
    targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)

    ' ファイルを開く
    Set newBook = Workbooks.Open(targetFile)
    Set wsTemplate = newBook.Sheets(1)

    ' G2, H2, I2 に情報を転記
    wsTemplate.Range("G2").Value = diagnosisPeriod
    wsTemplate.Range("I2").Value = sendDate
    wsTemplate.Range("J2").Value = ThisWorkbook.Sheets(1).Range("B1").Value

    ' シート名（CSVファイル名）
    sheetName = Left(fileName, 30)

    ' 既に転記済みかチェック
    If Not IsSheetExist(newBook, sheetName) Then
        ' 新規シート作成
        Set wsSheet = newBook.Sheets.Add(After:=newBook.Worksheets(newBook.Worksheets.Count))
        wsSheet.Name = sheetName

        ' CSVデータ転記
        ImportCSVData csvFile, wsSheet, fileType
    Else
        MsgBox "このCSVデータは既に転記済みです。", vbInformation
    End If

    ' 保存して閉じる
    newBook.Save
    newBook.Close
    MsgBox "処理が完了しました！", vbInformation
End Sub

' CSVデータを転記
Sub ImportCSVData(csvFile As String, ws As Worksheet, fileType As String)
    Dim colMap As Object
    Dim fso As Object
    Dim ts As Object
    Dim lineText As String
    Dim dataArray As Variant
    Dim i As Long, j As Long, key

    ' 項目マッピングを取得
    Set colMap = GetColumnMapping(fileType)

    ' 1行目に項目名を転記
    j = 1
    For Each key In colMap.Keys
        ws.Cells(1, j).Value = colMap(key)
        j = j + 1
    Next key

    ' CSVデータを読み込んで転記
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(csvFile, 1, False, -2)

    ' データを転記
    i = 2
    Do While Not ts.AtEndOfStream
        lineText = ts.ReadLine
        dataArray = Split(lineText, ",")

        If i > 2 Then
            j = 1
            For Each key In colMap.Keys
                If key <= UBound(dataArray) Then
                    ws.Cells(i - 1, j).Value = Trim(dataArray(key - 1))
                End If
                j = j + 1
            Next key
        End If
        i = i + 1
    Loop
    ts.Close

    ' 列幅を自動調整
    ws.Cells.EntireColumn.AutoFit
End Sub

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

' CSVの種類ごとに項目をマッピング
Function GetColumnMapping(fileType As String) As Object
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")

    If fileType = "振込額明細書" Then
        colMap.Add 2, "診療（調剤）年月"
        colMap.Add 5, "受付番号"
        colMap.Add 14, "氏名"
        colMap.Add 22, "医療保険＿療養の給付＿請求点数"
        colMap.Add 23, "医療保険＿療養の給付＿決定点数"
        colMap.Add 24, "医療保険＿療養の給付＿一部負担金"
        colMap.Add 25, "医療保険＿療養の給付＿金額"
        
        ' 第一公費
        colMap.Add 34, "第一公費_請求点数"
        colMap.Add 35, "第一公費_決定点数"
        colMap.Add 36, "第一公費_患者負担金"
        colMap.Add 37, "第一公費_金額"
        
        ' 第二公費
        colMap.Add 44, "第二公費_請求点数"
        colMap.Add 45, "第二公費_決定点数"
        colMap.Add 46, "第二公費_患者負担金"
        colMap.Add 47, "第二公費_金額"

        ' 第三公費
        colMap.Add 54, "第三公費_請求点数"
        colMap.Add 55, "第三公費_決定点数"
        colMap.Add 56, "第三公費_患者負担金"
        colMap.Add 57, "第三公費_金額"

        ' 第四公費
        colMap.Add 64, "第四公費_請求点数"
        colMap.Add 65, "第四公費_決定点数"
        colMap.Add 66, "第四公費_患者負担金"
        colMap.Add 67, "第四公費_金額"

        ' 第五公費
        colMap.Add 74, "第五公費_請求点数"
        colMap.Add 75, "第五公費_決定点数"
        colMap.Add 76, "第五公費_患者負担金"
        colMap.Add 77, "第五公費_金額"

        colMap.Add 82, "算定額合計"

    ElseIf fileType = "請求確定状況" Then
        colMap.Add 4, "診療（調剤）年月"
        colMap.Add 5, "氏名"
        colMap.Add 9, "医療機関名称"

        ' 各種合計点数
        colMap.Add 13, "総合計点数"
        colMap.Add 17, "医療保険＿療養の給付＿請求点数"
        colMap.Add 20, "第一公費_請求点数"
        colMap.Add 23, "第二公費_請求点数"
        colMap.Add 26, "第三公費_請求点数"
        colMap.Add 29, "第四公費_請求点数"

        colMap.Add 30, "請求確定状況"
    End If

    Set GetColumnMapping = colMap
End Function