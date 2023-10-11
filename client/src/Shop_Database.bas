Attribute VB_Name = "Shop_Database"
Public Sub ClearShop(ByVal Index As Long)
    Shop(Index) = EmptyShop
    Shop(Index).Name = vbNullString
End Sub

Public Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub
