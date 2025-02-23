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
    Dim wsTemplate2 as worksheet
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

    ' ファイル種類判定（fmei → 振込額明細書, fixf → 請求確定状況）pjry hasp uasp kasp hjry zogn henr edbn skkg skks chng fggk tgft sast shst sakr tttr
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

    ' 診療月（請求月の前月）
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
    diagnosisPeriod = targetYear & "年" & targetMonth & "月調剤分"

    ' 請求月
    sendMonth = receiptMonth + 1
    If sendMonth = 13 Then sendMonth = 1
    sendDate = sendMonth & "月10日請求分"

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
    Set wsTemplate = newBook.Sheets("A")
    Set wsTemplate2 = newBook.Sheets("B")

    wsTemplate.name = "R" & receiptYear-2018 & "." & receiptMonth
    wsTemplate2.name = ConvertToCircledNumber(receiptMonth)

    ' G2, H2, I2 に情報を転記
    wsTemplate.Range("G2").Value = diagnosisPeriod
    wsTemplate.Range("I2").Value = sendDate
    wsTemplate.Range("J2").Value = ThisWorkbook.Sheets(1).Range("B1").Value
    wsTemplate2.Range("H1").Value = diagnosisPeriod
    wsTemplate2.Range("J1").Value = sendDate
    wsTemplate2.Range("L1").Value = ThisWorkbook.Sheets(1).Range("B1").Value

    ' シート名（CSVファイル名）
    sheetName = Left(fileName, 30)

    ' 既に転記済みかチェック
    If Not IsSheetExist(newBook, sheetName) Then
        ' 新規シート作成
        Set wsSheet = newBook.Sheets.Add(After:=newBook.Worksheets(newBook.Worksheets.Count))
        wsSheet.Name = sheetName

        ' CSVデータ転記
        ImportCSVData csvFile, wsSheet, fileType
        ' **追加処理: シート2へ請求詳細データを転記**
        TransferBillingDetails(newBook, sheetName)   ' 別マクロで処理
    Else
        MsgBox "このCSVデータは既に転記済みです。", vbInformation
    End If

    ' 保存して閉じる
    newBook.Save
    newBook.Close
    MsgBox "処理が完了しました！", vbInformation
End Sub