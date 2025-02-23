' CSVの種類ごとに項目をマッピング
Function GetColumnMapping(fileType As String) As Object
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")

    If fileType = "振込額明細書" Then
        colMap.Add 2, "診療（調剤）年月"
        colMap.Add 5, "受付番号"
        colMap.Add 14, "氏名"
        colMap.Add 16, "生年月日"
        colMap.Add 22, "医療保険＿療養の給付＿請求点数"
        colMap.Add 23, "医療保険＿療養の給付＿決定点数"
        colMap.Add 24, "医療保険＿療養の給付＿一部負担金"
        colMap.Add 25, "医療保険＿療養の給付＿金額"
        
        ' 第一公費
        colMap.Add 34, "第一公費_請求点数"
        colMap.Add 35, "第一公費_決定点数"
        colMap.Add 36, "第一公費_患者負担金"
        colMap.Add 37, "第一公費_金額"
        
        ' 第二公費
        colMap.Add 44, "第二公費_請求点数"
        colMap.Add 45, "第二公費_決定点数"
        colMap.Add 46, "第二公費_患者負担金"
        colMap.Add 47, "第二公費_金額"

        ' 第三公費
        colMap.Add 54, "第三公費_請求点数"
        colMap.Add 55, "第三公費_決定点数"
        colMap.Add 56, "第三公費_患者負担金"
        colMap.Add 57, "第三公費_金額"

        ' 第四公費
        colMap.Add 64, "第四公費_請求点数"
        colMap.Add 65, "第四公費_決定点数"
        colMap.Add 66, "第四公費_患者負担金"
        colMap.Add 67, "第四公費_金額"

        ' 第五公費
        colMap.Add 74, "第五公費_請求点数"
        colMap.Add 75, "第五公費_決定点数"
        colMap.Add 76, "第五公費_患者負担金"
        colMap.Add 77, "第五公費_金額"

        colMap.Add 82, "算定額合計"

    ElseIf fileType = "請求確定状況" Then
        colMap.Add 4, "診療（調剤）年月"
        colMap.Add 5, "氏名"
        colMap.Add 7, "生年月日"
        colMap.Add 9, "医療機関名称"

        ' 各種合計点数
        colMap.Add 13, "総合計点数"
        colMap.Add 17, "医療保険＿療養の給付＿請求点数"
        colMap.Add 20, "第一公費_請求点数"
        colMap.Add 23, "第二公費_請求点数"
        colMap.Add 26, "第三公費_請求点数"
        colMap.Add 29, "第四公費_請求点数"

        colMap.Add 30, "請求確定状況"
        colMap.Add 31, "エラー区分"
    End If

    Set GetColumnMapping = colMap
End Function