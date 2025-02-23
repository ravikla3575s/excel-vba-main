Function ConvertToCircledNumber(num As Integer) As String
    Dim circledNumbers As Object
    Set circledNumbers = CreateObject("Scripting.Dictionary")
    
    ' 丸付き数字のマッピング（Unicode）
    circledNumbers.Add 1, "①"
    circledNumbers.Add 2, "②"
    circledNumbers.Add 3, "③"
    circledNumbers.Add 4, "④"
    circledNumbers.Add 5, "⑤"
    circledNumbers.Add 6, "⑥"
    circledNumbers.Add 7, "⑦"
    circledNumbers.Add 8, "⑧"
    circledNumbers.Add 9, "⑨"
    circledNumbers.Add 10, "⑩"
    circledNumbers.Add 11, "⑪"
    circledNumbers.Add 12, "⑫"
    
    ' 数値が対応範囲内かチェック
    If circledNumbers.exists(num) Then
        ConvertToCircledNumber = circledNumbers(num)
    Else
        ConvertToCircledNumber = CStr(num) ' 範囲外ならそのまま返す
    End If
End Function