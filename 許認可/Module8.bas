Sub ExcelToWordWith8Formats()
    ' エクセルのデータを変数に取得
    Dim CustomerName As String
    Dim Address As String
    Dim PhoneNumber As String
    Dim FormatChoice As String
    Dim TemplatePath As String

    ' データの取得
    CustomerName = Sheets("Sheet1").Range("A1").value ' 顧客名がA1にあると仮定
    Address = Sheets("Sheet1").Range("B1").value ' 住所がB1にあると仮定
    PhoneNumber = Sheets("Sheet1").Range("C1").value ' 電話番号がC1にあると仮定
    FormatChoice = Sheets("Sheet1").Range("D1").value ' フォーマット選択がD1にあると仮定

    ' フォーマットの選択に基づいてWordテンプレートを分岐
    Select Case FormatChoice
        Case "フォーマット1"
            TemplatePath = "C:pathtoyourtemplate1.docx"ocx"
        Case "フォーマット2"
            TemplatePath = "C:pathtoyourtemplate2.docx"ocx"
        Case "フォーマット3"
            TemplatePath = "C:pathtoyourtemplate3.docx"ocx"
        Case "フォーマット4"
            TemplatePath = "C:pathtoyourtemplate4.docx"ocx"
        Case "フォーマット5"
            TemplatePath = "C:pathtoyourtemplate5.docx"ocx"
        Case "フォーマット6"
            TemplatePath = "C:pathtoyourtemplate6.docx"ocx"
        Case "フォーマット7"
            TemplatePath = "C:pathtoyourtemplate7.docx"ocx"
        Case "フォーマット8"
            TemplatePath = "C:pathtoyourtemplate8.docx"ocx"
        Case Else
            MsgBox "フォーマットが正しく選択されていません。"
            Exit Sub
    End Select

    ' Wordアプリケーションを起動
    Dim WordApp As Object
    Dim WordDoc As Object
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = True

    ' 選択されたテンプレートを開く
    Set WordDoc = WordApp.Documents.Open(TemplatePath)
    
    ' プレースホルダーを探して置換する
    With WordDoc.Content.Find
        .Text = "<<CustomerName>>"
        .Replacement.Text = CustomerName
        .Execute Replace:=wdReplaceAll
        
        .Text = "<<Address>>"
        .Replacement.Text = Address
        .Execute Replace:=wdReplaceAll
        
        .Text = "<<PhoneNumber>>"
        .Replacement.Text = PhoneNumber
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Wordドキュメントを保存して閉じる
    WordDoc.SaveAs "C:pathtooutputfilled_template_" & FormatChoice & ".docx"ocx"
    WordDoc.Close
    WordApp.Quit

    ' オブジェクト解放
    Set WordDoc = Nothing
    Set WordApp = Nothing
End Sub
