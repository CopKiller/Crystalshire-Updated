Attribute VB_Name = "Conv_Database"
Public Sub ClearConv(ByVal Index As Long)
    Conv(Index) = EmptyConv
    Conv(Index).Name = vbNullString
    ReDim Conv(Index).Conv(1)
End Sub

Public Sub ClearConvs()
    Dim i As Long

    For i = 1 To MAX_CONVS
        Call ClearConv(i)
    Next

End Sub

