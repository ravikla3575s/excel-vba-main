Function IsStartRowOverlap(startRowDict As Object, newRow As Long) As Boolean
    Dim key As Variant
    For Each key In startRowDict.Keys
        If startRowDict(key) = newRow Then
            IsStartRowOverlap = True
            Exit Function
        End If
    Next key
    IsStartRowOverlap = False
End Function