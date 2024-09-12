Sub ExcelToWordWithFormats()
    Dim permitNumberAndDate As String
    Dim Address As String
    Dim PhoneNumber As String ' 電話番号変数の追加
    Dim jurisdictional As String
    Dim FormatChoice As String
    Dim TemplatePath As String
    Dim storeNames() As Variant ' 店舗名の配列
    Dim storeData() As Variant ' 店舗のデータを格納する配列
    Dim lastRow As Long
    Dim i As Long, j As Long, numberRow As Long, dateRow As Long
    Dim foundColumn As Long
    Dim currentStore As String
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim rng As Object

    ' 新たに追加した変数
    Dim submissionDate As String
    Dim dateOfChange As String
    Dim savePath As String
    Dim folderPath As String
    Dim selectedFormat As String

    ' シートの参照設定
    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = Sheets(1) ' 店舗名リストがあるシート
    Set ws2 = Sheets(2) ' 店舗データを検索するシート

    ' submissionDateとdateOfChangeをシートから取得
    submissionDate = ws1.Cells(2, 2).Value ' Sheet(1)のB2
    dateOfChange = ws1.Cells(1, 2).Value ' Sheet(1)のB1

    ' 電話番号をシートから取得
    PhoneNumber = ws1.Cells(12, 1).Value ' Sheet(1)の12行目の値を取得

    ' FormatChoiceはSheet(1)のB3セルに格納されていると仮定
    FormatChoice = ws1.Range("B3").Value

    ' 保存先のフォルダをユーザーに選ばせる
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "保存先のフォルダを選択してください"
        If .Show = -1 Then ' ユーザーがフォルダを選択した場合
            folderPath = .SelectedItems(1)
        Else
            MsgBox "保存先が選択されませんでした。処理を中止します。"
            Exit Sub
        End If
    End With

    ' 選択したFormatChoiceをフォルダ名として作成
    folderPath = folderPath & "\" & FormatChoice
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath ' フォルダが存在しない場合、作成
    End If

    ' 保存形式をユーザーに選ばせる（WordかPDF）
    selectedFormat = InputBox("保存形式を選択してください。1 = Word, 2 = PDF", "保存形式の選択")
    If selectedFormat <> "1" And selectedFormat <> "2" Then
        MsgBox "無効な形式が選択されました。処理を中止します。"
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

    ' 店舗名を1つずつ処理
    For i = 1 To UBound(storeNames, 1)
        currentStore = storeNames(i, 1) ' 現在の店舗名を取得

        ' Sheet(2)の5行目を検索して店舗名を見つけた列番号を取得
        On Error Resume Next ' エラー処理（見つからない場合）
        foundColumn = 0
        foundColumn = ws2.Rows(5).Find(What:=currentStore, LookIn:=xlValues, LookAt:=xlWhole).Column
        On Error GoTo 0 ' エラー処理を解除

        If foundColumn > 0 Then
            ' 店舗名が見つかった場合、その列の6?218行目を配列に格納
            storeData = ws2.Range(ws2.Cells(6, foundColumn), ws2.Cells(218, foundColumn)).Value

            ' jurisdictionalの値を確認
            jurisdictional = storeData(212, 1)

            ' 旭川市の場合、別のフォーマットを使用
            If jurisdictional = "旭川市" Then
                Select Case FormatChoice
                    Case "薬局開設許可"
                        TemplatePath = "C:\Users\ゲスト\OneDrive\ドキュメント\Office のカスタム テンプレート\薬局変更届_旭川市.dotm"
                        numberRow = 30
                        dateRow = 31
                    Case "高度管理医療機器等販売業"
                        TemplatePath = "C:\Users\admin\Desktop\旭川市用\高度管理用.dotm"
                        numberRow = 46
                        dateRow = 44
                    Case "毒物劇物等一般販売業"
                        TemplatePath = "C:\Users\admin\Desktop\旭川市用\毒劇物用.dotm"
                        numberRow = 42
                        dateRow = 40
                    Case "麻薬小売業"
                        TemplatePath = "C:\Users\admin\Desktop\旭川市用\麻薬小売業.dotm"
                        numberRow = 41
                        dateRow = 42
                    Case Else
                        MsgBox "フォーマットが正しく選択されていません。"
                        Exit Sub
                End Select
            Else
                ' 旭川市以外の場合のフォーマット
                Select Case FormatChoice
                    Case "薬局開設許可"
                        TemplatePath = "C:\Users\ゲスト\OneDrive\ドキュメント\Office のカスタム テンプレート\薬局変更届_旭川市以外.dotm"
                        numberRow = 30
                        dateRow = 31
                    Case "高度管理医療機器等販売業"
                        TemplatePath = "C:\Users\admin\Desktop\道保健所（旭川市以外）\高度管理変更届.dotm"
                        numberRow = 46
                        dateRow = 44
                    Case "毒物劇物等一般販売業"
                        TemplatePath = "C:\Users\admin\Desktop\道保健所（旭川市以外）\毒劇物09変更届.dotm"
                        numberRow = 42
                        dateRow = 40
                    Case "麻薬小売業"
                        TemplatePath = "C:\Users\admin\Desktop\道保健所（旭川市以外）\麻薬小売業.dotm"
                        numberRow = 41
                        dateRow = 42
                    Case Else
                        MsgBox "フォーマットが正しく選択されていません。"
                        Exit Sub
                End Select
            End If

            ' テンプレートから新しいドキュメントを作成
            Set WordDoc = WordApp.Documents.Add(TemplatePath)

            ' storeData配列内のデータを確認
            storeName = storeData(1, 1)
            permitNumber = storeData(numberRow, 1)
            permitDate = storeData(dateRow, 1)
            Address = storeData(6, 1)

            ' プレースホルダーの置換
            WordDoc.Content.Find.Execute FindText:="<<storeName>>", ReplaceWith:=storeName, Replace:=2
            WordDoc.Content.Find.Execute FindText:="<<permitNumber>>", ReplaceWith:=permitNumber, Replace:=2
            WordDoc.Content.Find.Execute FindText:="<<permitDate>>", ReplaceWith:=permitDate, Replace:=2
            WordDoc.Content.Find.Execute FindText:="<<address>>", ReplaceWith:=Address, Replace:=2
            WordDoc.Content.Find.Execute FindText:="<<jurisdictional>>", ReplaceWith:=jurisdictional, Replace:=2
            WordDoc.Content.Find.Execute FindText:="<<submisstionDate>>", ReplaceWith:=submissionDate, Replace:=2
            WordDoc.Content.Find.Execute FindText:="<<dateOfChange>>", ReplaceWith:=dateOfChange, Replace:=2
            WordDoc.Content.Find.Execute FindText:="<<phoneNumber>>", ReplaceWith:=PhoneNumber, Replace:=2 ' 電話番号の置換

            ' ファイル保存の処理
            If selectedFormat = "1" Then
                ' Wordファイルとして保存 (薬局開設許可用: "薬局変更届_店舗名" の形式)
                If FormatChoice = "薬局開設許可" Then
                    savePath = folderPath & "\薬局変更届_" & currentStore & ".docx"
                Else
                    savePath = folderPath & "\" & FormatChoice & "_" & currentStore & ".docx"
                End If
                WordDoc.SaveAs2 Filename:=savePath, FileFormat:=12 ' FileFormat:=12は.docx形式
            ElseIf selectedFormat = "2" Then
                ' PDFとして保存 (薬局開設許可用: "薬局変更届_店舗名" の形式)
                If FormatChoice = "薬局開設許可" Then
                    savePath = folderPath & "\薬局変更届_" & currentStore & ".pdf"
                Else
                    savePath = folderPath & "\" & FormatChoice & "_" & currentStore & ".pdf"
                End If
                WordDoc.SaveAs2 Filename:=savePath, FileFormat:=17 ' FileFormat:=17はPDF形式
            End If

            ' ドキュメントを閉じる (保存確認の回避)
            WordDoc.Close SaveChanges:=False
            Set WordDoc = Nothing

        Else
            ' 店舗名が見つからない場合の処理
            Debug.Print "店舗名 '" & currentStore & "' がSheet(2)の5行目に見つかりません。"
        End If
    Next i

    MsgBox "すべての店舗情報の処理が完了しました。"
    Exit Sub

ErrorHandlerLoop:
    MsgBox "エラーが発生しました。エラー内容: " & Err.Description
    If Not WordDoc Is Nothing Then WordDoc.Close SaveChanges:=False
    Resume Next
End Sub

