Function ConvertToWesternDate(dispensingMonth As String) As String
    Dim era As String, yearPart As Integer, westernYear As Integer, monthPart As String
    
    ' GYYMM 形式から元号と年月を取得
    era = Left(dispensingMonth, 1) ' 例: "5"（令和）
    yearPart = Mid(dispensingMonth, 2, 2) ' 例: "06"
    monthPart = Right(dispensingMonth, 2) ' 例: "06"

    ' 和暦を西暦に変換
    Select Case era
        Case "5": westernYear = 2018 + yearPart ' 令和（2019年開始）
        ' 他の元号（明治/大正/昭和/平成）は未対応
    End Select

    ' 変換結果（YY.MM）
    ConvertToWesternDate = Right(westernYear, 2) & "." & monthPart
End Function