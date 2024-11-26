Sub ImportDataAndTransfer()

    ' ファイル選択ダイアログを表示
    Dim selectedFile As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim dataArray(1 To 13) As Variant
    Dim i As Long
    
    ' ファイル選択ダイアログを表示
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "勤務者情報ファイルを選択してください"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
        Else
            MsgBox "ファイルが選択されませんでした。", vbExclamation
            Exit Sub
        End If
    End With

    ' 選択されたファイルを開く
    On Error GoTo ErrorHandler
    Set wbSource = Workbooks.Open(selectedFile)
    Set wsSource = wbSource.Sheets(1) ' 勤務者情報ファイルの最初のシートを参照
    
    ' データを配列に格納
    dataArray(1) = wsSource.Range("B3").Value ' 社員番号
    dataArray(2) = NormalizeName(wsSource.Range("B4").Value) ' 氏名（全角スペース統一）
    dataArray(3) = NConvertToHalfWidth(wsSource.Range("B5").Value) ' シメイ
    dataArray(4) = wsSource.Range("B6").Value ' 保健薬剤師記号
    dataArray(5) = wsSource.Range("B7").Value ' 保健薬剤師登録番号
    dataArray(6) = wsSource.Range("B8").Value ' 薬剤師番号
    dataArray(7) = wsSource.Range("B9").Value ' 薬剤師番号登録日
    dataArray(8) = wsSource.Range("B10").Value ' 生年月日
    dataArray(9) = wsSource.Range("B11").Value ' 郵便番号
    dataArray(10) = wsSource.Range("B12").Value ' 都道府県
    dataArray(11) = NormalizeAddress(wsSource.Range("B13").Value) ' 住所（スペースと数字の形式統一）
    dataArray(12) = ExtractNumber(wsSource.Range("B14").Value) ' 週労働時間（数字のみ抽出）
    dataArray(13) = wsSource.Range("B15").Value ' 資格者区分
    
    ' 元ファイルを閉じる
    wbSource.Close SaveChanges:=False
    
    ' 転記先シートを設定
    Set wsTarget = ThisWorkbook.Sheets("届出一覧テーブル")
    
    ' データを転記
    For i = 1 To 13
        wsTarget.Cells(i + 1, 2).Value = dataArray(i) ' B列に転記（1行目はヘッダー想定）
    Next i

    MsgBox "データの転記が完了しました。", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    On Error GoTo 0
End Sub

' ----------------------------------------------------------
' 補助関数：氏名を全角スペースに統一
Function NormalizeName(name As String) As String
    NormalizeName = Replace(name, " ", "　") ' 半角スペースを全角スペースに置換
End Function

' 補助関数：文字列を半角に変換
Function ConvertToHalfWidth(value As String) As String
    Dim i As Long
    Dim char As String
    Dim result As String
    For i = 1 To Len(value)
        char = Mid(value, i, 1)
        ' 全角英数字・カタカナを半角に変換
        If AscW(char) >= &HFF01 And AscW(char) <= &HFF5E Then
            char = ChrW(AscW(char) - &HFEE0)
        ElseIf AscW(char) >= &H30A1 And AscW(char) <= &H30FC Then
            char = StrConv(char, vbNarrow) ' カタカナを半角に変換
        End If
        result = result & char
    Next i
    ConvertToHalfWidth = result
End Function

' 補助関数：住所のスペースを半角スペースに統一し、数字を全角に統一
Function NormalizeAddress(address As String) As String
    Dim i As Long
    Dim char As String
    Dim result As String
    result = Replace(address, "　", " ") ' 全角スペースを半角スペースに置換
    For i = 1 To Len(result)
        char = Mid(result, i, 1)
        If char Like "[0-9]" Then
            char = ChrW(AscW(char) + &HFF10 - &H30) ' 半角数字を全角数字に変換
        End If
        result = Application.Replace(result, i, 1, char)
    Next i
    NormalizeAddress = result
End Function

' 補助関数：文字列から数字のみ抽出
Function ExtractNumber(value As String) As Double
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "[^\d]" ' 数字以外を検出する正規表現
    regex.Global = True
    ExtractNumber = Val(regex.Replace(value, "")) ' 数字以外を削除して数値に変換
End Function