Sub IncreaseAllStartRows(startRowDict As Object)
    Dim key As Variant
    For Each key In startRowDict.Keys
        startRowDict(key) = startRowDict(key) + 1
    Next key
End Sub