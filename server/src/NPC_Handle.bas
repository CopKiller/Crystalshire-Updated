Attribute VB_Name = "NPC_Handle"
' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Public Sub HandleRequestEditNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcEditor
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub HandleRequestNPCS(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendNpcs index
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Public Sub HandleSaveNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim npcNum As Long
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    npcNum = Buffer.ReadLong

    ' Prevent hacking
    If npcNum < 0 Or npcNum > MAX_NPCS Then
        Exit Sub
    End If

    NPCSize = LenB(Npc(npcNum))
    ReDim NPCData(NPCSize - 1)
    NPCData = Buffer.ReadBytes(NPCSize)
    CopyMemory ByVal VarPtr(Npc(npcNum)), ByVal VarPtr(NPCData(0)), NPCSize
    ' Save it
    Call SendUpdateNpcToAll(npcNum)
    Call SaveNpc(npcNum)
    Call AddLog(GetPlayerName(index) & " saved Npc #" & npcNum & ".", ADMIN_LOG)
End Sub
