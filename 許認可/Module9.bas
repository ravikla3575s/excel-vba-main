Sub ExcelToWordWith8Formats_Mac()
    Dim permitNumberAndDate As String
    Dim Address As String
    Dim PhoneNumberAndData As String
    Dim FormatChoice As String
    Dim TemplatePath As String
    Dim storeData As Variant
    Dim i As Long
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim rng As Object

    ' C19からC252の範囲を配列に格納
    storeData = Sheets(1).Range("C19:C252").value

    ' 配列のデータを確認（例として出力）
    For i = LBound(storeData) To UBound(storeData)
        Debug.Print i & ":" & storeData(i, 1) ' データを出力
    Next i

    ' フォーマット選択がC14にあると仮定
    FormatChoice = Sheets(1).Range("C14").value

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
            TemplatePath = "/Users/yoshipc/Documents/tsuruha/テンプレート/高度管理医療機器等販売業許可更新申請書_フォーマット.dotm"
            
            ' permitNumberAndDateを代入（C54とC56の値を組み合わせ）
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

    ' プレースホルダーを検索して置換する（Rangeオブジェクトを使用）
    ' permitNumberAndDate の置換
    Set rng = WordDoc.Content
    rng.Find.Execute FindText:="<<permitNumberAndDate>>", ReplaceWith:=permitNumberAndDate, Replace:=2

    ' 顧客名の置換
    rng.Find.Execute FindText:="<<CustomerName>>", ReplaceWith:="山田 太郎", Replace:=2

    ' 住所の置換
    rng.Find.Execute FindText:="<<Address>>", ReplaceWith:="東京都新宿区", Replace:=2

    ' 電話番号の置換
    rng.Find.Execute FindText:="<<PhoneNumberAndData>>", ReplaceWith:="090-1234-5678", Replace:=2

    ' ドキュメントを指定したパスに保存
    Dim savePath As String
    savePath = "/Users/yoshipc/Documents/tsuruha/permit/PDFs/output_document.dosx"
    WordDoc.SaveAs2 savePath

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
