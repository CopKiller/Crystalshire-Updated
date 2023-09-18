Attribute VB_Name = "Item_TCP"
Public Sub SendUpdateItemTo(ByVal index As Long, ByVal itemNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(itemNum))
    
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemNum)), ItemSize
    
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong itemNum
    Buffer.WriteBytes ItemData
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendItems(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(index, i)
        End If

    Next

End Sub

Public Sub SendUpdateItemToAll(ByVal itemNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(itemNum))
    
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemNum)), ItemSize
    
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong itemNum
    Buffer.WriteBytes ItemData
    
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
