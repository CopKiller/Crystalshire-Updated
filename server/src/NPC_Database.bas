Attribute VB_Name = "NPC_Database"
' **********
' ** NPCs **
' **********
Public Sub SaveNpc(ByVal npcNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\npcs\npc" & npcNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Npc(npcNum)
    Close #f
End Sub

Public Sub SaveNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next

End Sub

Public Sub CheckNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS

        If Not FileExist(App.Path & "\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If

    Next

End Sub

Public Sub LoadNpcs()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckNpcs

    For i = 1 To MAX_NPCS
        filename = App.Path & "\data\npcs\npc" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Npc(i)
        Close #f
    Next

End Sub

Public Sub ClearNpc(ByVal index As Long)
    Npc(index) = EmptyNpc
    Npc(index).Name = vbNullString
    Npc(index).AttackSay = vbNullString
    Npc(index).Sound = "None."
End Sub

Public Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

End Sub
