Attribute VB_Name = "Item_Database"
' ***********
' ** Items **
' ***********

Public Sub SaveItem(ByVal ItemNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\items\item" & ItemNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Item(ItemNum)
    Close #f
End Sub

Public Sub SaveItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next

End Sub

Public Sub CheckItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Not FileExist(App.Path & "\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If

    Next

End Sub

Public Sub LoadItems()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckItems

    For i = 1 To MAX_ITEMS
        filename = App.Path & "\data\Items\Item" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Item(i)
        Close #f
    Next

End Sub

Public Sub ClearItem(ByVal index As Long)
    Item(index) = EmptyItem
    Item(index).Name = vbNullString
    Item(index).Desc = vbNullString
    Item(index).Sound = "None."
End Sub

Public Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub
