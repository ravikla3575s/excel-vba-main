Sub ProcessReceiptCSV()
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
    Dim wsSheet3 As Worksheet
    Dim fso As Object
    Dim ts As Object
    Dim lineText As String
    Dim dataArray As Variant
    Dim i As Long, j As Long
    
    ' CSVファイル選択ダイアログを表示
    csvFile = Application.GetOpenFilename("CSVファイル (*.csv), *.csv", , "振込額明細書のCSVファイルを選択してください")
    If csvFile = "False" Then Exit Sub

    ' ファイル名から診療年月を判定
    fileName = Dir(csvFile) ' フルパスからファイル名を取得
    If Len(fileName) < 22 Then
        MsgBox "ファイル名の形式が正しくありません。", vbExclamation
        Exit Sub
    End If

    ' 令和年と振込月を取得（12~16桁目）
    reiwaYear = CInt(Mid(fileName, 19, 2))  ' 例: "06" → 令和6年
    receiptMonth = CInt(Mid(fileName, 21, 2)) ' 例: "12" → 12月振込

    ' 西暦に変換（令和1年 = 2019年）
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

    ' 送信月（振込月と同じ月）
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

    ' テンプレートファイルを開く
    Set newBook = Workbooks.Open(templatePath)
    Set wsTemplate = newBook.Sheets(1)

    ' G2, H2, I2 に情報を転記
    wsTemplate.Range("G2").Value = diagnosisPeriod ' 診療年月
    wsTemplate.Range("I2").Value = sendDate ' 送信日
    wsTemplate.Range("J2").Value = ThisWorkbook.Sheets(1).Range("B1").Value ' 店舗名

    ' 新しいファイルを保存
    Dim newFilePath As String
    newFilePath = savePath & "\保険請求管理報告書_" & targetYear & targetMonth & ".xlsx"
    newBook.SaveAs newFilePath, FileFormat:=xlOpenXMLWorkbook

    ' === 作成したExcelファイルのシート3にCSVデータを転記 ===
    On Error Resume Next
    Set wsSheet3 = newBook.Sheets("振込額明細書")
    If wsSheet3 Is Nothing Then
        Set wsSheet3 = newBook.Sheets.Add(After:=newBook.Worksheets(2))
        wsSheet3.Name = "振込額明細書"
    End If
    On Error GoTo 0

    ' 列マッピング（列番号 → 項目名）
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    colMap.Add 2, "診療（調剤）年月"
    colMap.Add 3, "処理区分"
    colMap.Add 5, "受付番号"
    colMap.Add 7, "診療科＿診療科名"
    colMap.Add 14, "氏名"
    colMap.Add 22, "医療保険＿療養の給付＿請求点数"
    colMap.Add 23, "医療保険＿療養の給付＿決定点数"
    colMap.Add 24, "医療保険＿療養の給付＿一部負担金"
    colMap.Add 25, "医療保険＿療養の給付＿金額"
    colMap.Add 29, "医療保険＿算定額"
    colMap.Add 82, "算定額合計"

    ' 1行目に項目名を転記
    j = 1
    For Each key In colMap.Keys
        wsSheet3.Cells(1, j).Value = colMap(key)
        j = j + 1
    Next key

    ' CSVファイルを開いてデータを取得
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(csvFile, 1, False, -2)

    ' 3行目以降のデータを取得
    i = 2
    Do While Not ts.AtEndOfStream
        lineText = ts.ReadLine
        If i > 2 Then ' 3行目以降
            dataArray = Split(lineText, ",")
            j = 1
            For Each key In colMap.Keys
                If key <= UBound(dataArray) Then
                    wsSheet3.Cells(i - 1, j).Value = Trim(dataArray(key - 1))
                End If
                j = j + 1
            Next key
        End If
        i = i + 1
    Loop
    ts.Close

    ' 列幅を自動調整
    wsSheet3.Cells.EntireColumn.AutoFit

    ' ファイルを閉じて保存
    newBook.Close SaveChanges:=True

    ' 完了メッセージ
    MsgBox "処理が完了しました！" & vbCrLf & "新規ファイル: " & newFilePath, vbInformation

End Sub