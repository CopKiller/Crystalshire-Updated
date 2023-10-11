Attribute VB_Name = "Resource_Database"
Public Sub ClearResource(ByVal Index As Long)
    Resource(Index) = EmptyResource
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).sound = "None."
End Sub

Public Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next

End Sub
