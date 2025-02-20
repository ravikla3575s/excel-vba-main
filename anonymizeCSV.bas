Sub AnonymizeCSV_WithDialog()
    Dim inputFile As Integer
    Dim outputFile As Integer
    Dim line As String
    Dim lines As Variant
    Dim fields As Variant
    Dim i As Integer, j As Integer
    Dim inputFilePath As String
    Dim outputFilePath As String
    Dim outputFolder As String
    Dim fileName As String
    Dim fileExt As String
    Dim encodedPrefix As String
    Dim data() As Variant
    Dim rowCount As Integer
    Dim columnMappoling As Object
    
    ' ユーザーに範囲指定を求める
    Set columnMapping = CreateObject("Scripting.Dictionary")
    columnMapping("個人名") = InputBox("個人名の列番号を入力してください（空欄可）:")
    columnMapping("住所") = InputBox("住所の列番号を入力してください（空欄可）:")
    columnMapping("年齢") = InputBox("年齢の列番号を入力してください（空欄可）:")
    columnMapping("性別") = InputBox("性別の列番号を入力してください（空欄可）:")
    columnMapping("店舗名") = InputBox("店舗名の列番号を入力してください（空欄可）:")
    columnMapping("社長名") = InputBox("社長名の列番号を入力してください（空欄可）:")
    columnMapping("店舗住所") = InputBox("店舗住所の列番号を入力してください（空欄可）:")
    columnMapping("医療機関コード") = InputBox("医療機関コードの列番号を入力してください（空欄可）:")
    columnMapping("処方元医療機関名") = InputBox("処方元医療機関名の列番号を入力してください（空欄可）:")
    
    ' CSVファイル選択ダイアログ
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "CSVファイルを選択してください"
        .Filters.Clear
        .Filters.Add "CSVファイル", "*.csv"
        .AllowMultiSelect = False
        If .Show = -1 Then
            inputFilePath = .SelectedItems(1)
        Else
            MsgBox "処理がキャンセルされました。", vbExclamation
            Exit Sub
        End If
    End With
    
    ' ファイル名と拡張子の分解
    fileName = Mid(inputFilePath, InStrRev(inputFilePath, "\") + 1) ' ファイル名取得
    fileExt = Mid(fileName, InStrRev(fileName, ".")) ' 拡張子取得
    fileName = Left(fileName, InStrRev(fileName, ".") - 1) ' 拡張子を除く
    
    ' ランダムな英文字3文字を生成
    encodedPrefix = RandomString(3)
    
    ' 出力先ファイル名（エンコード済みのプレフィックスを付加）
    outputFilePath = Left(inputFilePath, InStrRev(inputFilePath, "\")) & encodedPrefix & "_" & fileName & fileExt
    
    ' 入出力ファイルを開く
    inputFile = FreeFile()
    Open inputFilePath For Input As #inputFile
    
    ' データを2次元配列に格納
    rowCount = 0
    Do Until EOF(inputFile)
        Line Input #inputFile, line
        rowCount = rowCount + 1
        ReDim Preserve data(1 To rowCount)
        data(rowCount) = Split(line, ",")
    Loop
    Close #inputFile
    
    ' 個人情報の処理
    For i = 1 To rowCount
        For Each Key In columnMapping.Keys
            If columnMapping(Key) <> "" Then
                j = CInt(columnMapping(Key)) - 1
                If j >= LBound(data(i)) And j <= UBound(data(i)) Then
                    Select Case Key
                        Case "個人名"
                            data(i)(j) = GenerateRandomName()
                        Case "住所", "店舗住所"
                            data(i)(j) = GenerateRandomAddress()
                        Case "年齢"
                            data(i)(j) = Int(Rnd() * 100) + 1
                        Case "性別"
                            data(i)(j) = IIf(Rnd() > 0.5, "男", "女")
                        Case "店舗名"
                            data(i)(j) = "〇〇調剤薬局 △△店"
                        Case "社長名"
                            data(i)(j) = "〇〇 〇〇"
                        Case "医療機関コード"
                            data(i)(j) = Format(Int(Rnd() * 10000000), "0000000")
                        Case "処方元医療機関名"
                            data(i)(j) = GenerateMedicalInstitutionName()
                    End Select
                End If
            End If
        Next Key
    Next i
    
    ' 出力ファイルに書き込み
    outputFile = FreeFile()
    Open outputFilePath For Output As #outputFile
    For i = 1 To rowCount
        Print #outputFile, Join(data(i), ",")
    Next i
    Close #outputFile
    
    MsgBox "処理が完了しました！出力ファイル: " & outputFilePath, vbInformation
End Sub

' --- ランダムな日本人の名前を生成 ---
Function GenerateRandomName() As String
    Dim lastNames As Variant, firstNames As Variant
    lastNames = Array("佐藤", "鈴木", "高橋", "田中", "伊藤", "山本", "中村", "小林", "加藤", "吉田", "山田", "佐々木", "松本", "井上", "木村", "斎藤", "林", "清水", "山崎", "森")
    firstNames = Array("太郎", "次郎", "三郎", "花子", "美咲", "陽菜", "悠斗", "蓮", "結衣", "翔")
    GenerateRandomName = lastNames(Int(Rnd() * 20)) & " " & firstNames(Int(Rnd() * 10))
End Function

' --- ランダムな住所を生成 ---
Function GenerateRandomAddress() As String
    GenerateRandomAddress = "北海道" & Int(Rnd() * 50 + 1) & "市 " & Int(Rnd() * 20 + 1) & "条 " & Int(Rnd() * 20 + 1) & "丁目 " & Int(Rnd() * 100 + 1) & "-" & Int(Rnd() * 10 + 1)
End Function

' --- ランダムな医療機関名を生成 ---
Function GenerateMedicalInstitutionName() As String
    Dim names As Variant, types As Variant
    names = Array("〇〇", "△△", "市立◻︎◻︎")
    types = Array("病院", "医院", "歯科", "クリニック")
    GenerateMedicalInstitutionName = names(Int(Rnd() * 3)) & types(Int(Rnd() * 4))
End Function
