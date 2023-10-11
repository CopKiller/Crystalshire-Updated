Attribute VB_Name = "NPC_Database"
Public Sub ClearNPC(ByVal Index As Long)
    Npc(Index) = EmptyNpc
    Npc(Index).Name = vbNullString
    Npc(Index).sound = "None."
End Sub

Public Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next

End Sub
