Attribute VB_Name = "Shop_Database"
' ***********
' ** Shops **
' ***********
Public Sub SaveShop(ByVal shopNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\shops\shop" & shopNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Shop(shopNum)
    Close #f
End Sub

Public Sub SaveShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next

End Sub

Public Sub CheckShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Not FileExist(App.Path & "\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If

    Next

End Sub

Public Sub LoadShops()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckShops

    For i = 1 To MAX_SHOPS
        filename = App.Path & "\data\shops\shop" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Shop(i)
        Close #f
    Next

End Sub

Public Sub ClearShop(ByVal index As Long)
    Shop(index) = EmptyShop
    Shop(index).Name = vbNullString
End Sub

Public Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub
