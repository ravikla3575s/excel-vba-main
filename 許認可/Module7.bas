Sub AddPharmacistDataFromSearchSheet()
    Dim wsMaster As Worksheet
    Dim wsSearch As Worksheet
    Dim lastRow As Long
    Dim newRow As Long
    
    ' 薬剤師マスタシートと検索シートをセット
    Set wsMaster = ThisWorkbook.Sheets("薬剤師マスタ")
    Set wsSearch = ThisWorkbook.Sheets("検索") ' 実際のシート名に変更してください
    
    ' 薬剤師マスタの最終行を確認
    lastRow = wsMaster.Cells(wsMaster.Rows.count, "A").End(xlUp).Row
    newRow = lastRow + 1 ' 新しいデータの行

    ' 検索シートからデータを取得し、薬剤師マスタシートに追加
    wsMaster.Cells(newRow, 1).value = wsSearch.Cells(15, 11).value ' 社員番号 (A列)
    wsMaster.Cells(newRow, 4).value = wsSearch.Cells(16, 11).value ' 氏名 (D列)
    wsMaster.Cells(newRow, 7).value = wsSearch.Cells(17, 11).value ' ｼﾒｲ (G列)
    wsMaster.Cells(newRow, 8).value = wsSearch.Cells(27, 11).value ' 資格区分 (H列)
    wsMaster.Cells(newRow, 9).value = wsSearch.Cells(18, 11).value ' 保険薬剤師記号 (I列)
    wsMaster.Cells(newRow, 10).value = wsSearch.Cells(19, 11).value ' 保険薬剤師登録番号 (J列)
    wsMaster.Cells(newRow, 11).value = wsSearch.Cells(20, 11).value ' 薬剤師番号 (K列)
    wsMaster.Cells(newRow, 12).value = wsSearch.Cells(21, 11).value ' 薬剤師番号登録日 (L列)
    wsMaster.Cells(newRow, 13).value = wsSearch.Cells(22, 11).value ' 生年月日 (M列)
    wsMaster.Cells(newRow, 14).value = wsSearch.Cells(23, 11).value ' 郵便番号 (N列)
    wsMaster.Cells(newRow, 15).value = wsSearch.Cells(24, 11).value ' 都道府県 (O列)
    wsMaster.Cells(newRow, 17).value = wsSearch.Cells(25, 11).value ' 住所 (Q列)
    wsMaster.Cells(newRow, 19).value = wsSearch.Cells(26, 11).value ' 週労働時間 (S列)

    ' データが正しく追加されたことを知らせる
    MsgBox "検索シートから新しいデータが追加されました。"
    
End Sub

