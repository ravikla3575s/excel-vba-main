Sub ExcelToWordWith8Formats_Mac()
    Dim permitNumberAndDate As String
    Dim Address As String
    Dim PhoneNumberAndData As String
    Dim FormatChoice As String
    Dim TemplatePath As String
    Dim storeNames() As Variant ' 店舗名の配列
    Dim storeData() As Variant ' 店舗のデータを格納する配列
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim foundColumn As Long
    Dim currentStore As String
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim rng As Object

    ' シートの参照設定
    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = Sheets(1) ' 店舗名リストがあるシート
    Set ws2 = Sheets(2) ' 店舗データを検索するシート

    ' FormatChoiceはSheet(1)のB3セルに格納されていると仮定
    FormatChoice = ws1.Range("B3").Value

    ' フォーマットの選択に基づいてWordテンプレートを分岐
    Select Case FormatChoice
        Case "フォーマット1"
            TemplatePath = "Macintosh HD:Users:yourusername:path:to:your:template1.docx"
        Case "フォーマット2"
            TemplatePath = "Macintosh HD:Users:yourusername:path:to:your:template2.docx"
        Case "フォーマット3"
            TemplatePath = "Macintosh HD:Users:yourusername:path:to:your:template3.docx"
        Case "フォーマット4"
            TemplatePath = "Macintosh HD:Users:yourusername:path:to:your:template4.docx"
        Case "高度管理医療機器等販売業"
            TemplatePath = "Macintosh HD:Users:yourusername:path:to:高度管理医療機器等販売業許可更新申請書_フォーマット.docx"
            permitNumberAndDate = storeData(36, 1) & "　　" & storeData(38, 1)
        Case "フォーマット6"
            TemplatePath = "Macintosh HD:Users:yourusername:path:to:your:template6.docx"
        Case "フォーマット7"
            TemplatePath = "Macintosh HD:Users:yourusername:path:to:your:template7.docx"
        Case "フォーマット8"
            TemplatePath = "Macintosh HD:Users:yourusername:path:to:your:template8.docx"
        Case Else
            MsgBox "フォーマットが正しく選択されていません。"
            Exit Sub
    End Select

    ' テンプレートファイルが存在するか確認
    If Dir(TemplatePath) = "" Then
        MsgBox "テンプレートファイルが見つかりません。" & vbCrLf & TemplatePath
        Exit Sub
    End If

    ' 店舗名リストの最終行を取得（D列、空欄まで）
    lastRow = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row

    ' 店舗名をD2から最終行まで配列に格納
    storeNames = ws1.Range("D2:D" & lastRow).Value

    ' Wordアプリケーションを起動
    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application")
    On Error GoTo 0

    If WordApp Is Nothing Then
        Set WordApp = CreateObject("Word.Application")
    End If

    ' エラーチェック
    If WordApp Is Nothing Then
        MsgBox "Wordアプリケーションを起動できませんでした。"
        Exit Sub
    End If

    ' Wordのテンプレートを開く
    On Error GoTo ErrorHandler
    Set WordDoc = WordApp.Documents.Open(TemplatePath)

    ' Wordアプリケーションを可視化
    WordApp.Visible = True

    ' 店舗名を1つずつ処理
    For i = 1 To UBound(storeNames, 1)
        currentStore = storeNames(i, 1) ' 現在の店舗名を取得

        ' Sheet(2)の5行目を検索して店舗名を見つけた列番号を取得
        On Error Resume Next ' エラー処理（見つからない場合）
        foundColumn = 0
        foundColumn = ws2.Rows(5).Find(What:=currentStore, LookIn:=xlValues, LookAt:=xlWhole).Column
        On Error GoTo 0 ' エラー処理を解除

        If foundColumn > 0 Then
            ' 店舗名が見つかった場合、その列の6〜218行目を配列に格納
            storeData = ws2.Range(ws2.Cells(6, foundColumn), ws2.Cells(218, foundColumn)).Value

            ' storeData配列内のデータを確認（例としてイミディエイトウィンドウに出力）
            For j = LBound(storeData, 1) To UBound(storeData, 1)
                Debug.Print "店舗名: " & currentStore & " - 行 " & j + 5 & ": " & storeData(j, 1)
            Next j

            ' プレースホルダーの置換
            Set rng = WordDoc.Content
            rng.Find.Execute FindText:="<<permitNumberAndDate>>", ReplaceWith:=permitNumberAndDate, Replace:=2
        Else
            ' 店舗名が見つからない場合の処理
            Debug.Print "店舗名 '" & currentStore & "' がSheet(2)の5行目に見つかりません。"
        End If
    Next i

    ' ドキュメントを指定したパスに保存
    Dim savePath As String
    savePath = "Macintosh HD:Users:yourusername:path:to:output_document.docx"
    
    ' ファイル形式を指定して保存（.docx形式）
    WordDoc.SaveAs2 FileName:=savePath, FileFormat:=12 ' FileFormat:=12は.docx形式

    ' ドキュメントを閉じる
    WordDoc.Close
    WordApp.Quit

    ' オブジェクトの解放
    Set WordDoc = Nothing
    Set WordApp = Nothing

    MsgBox "処理が完了しました。ファイルは " & savePath & " に保存されました。"
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました。エラー内容: " & Err.Description
    If Not WordDoc Is Nothing Then WordDoc.Close False
    If Not WordApp Is Nothing Then WordApp.Quit
    Set WordDoc = Nothing
    Set WordApp = Nothing
End Sub
