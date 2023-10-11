Attribute VB_Name = "NPC_TCP"
Public Sub SendRequestEditNpc()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditNpc
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendSaveNpc(ByVal NpcNum As Long)
    Dim buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    Set buffer = New clsBuffer
    NpcSize = LenB(Npc(NpcNum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(Npc(NpcNum)), NpcSize
    buffer.WriteLong CSaveNpc
    buffer.WriteLong NpcNum
    buffer.WriteBytes NpcData
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendRequestNPCS()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestNPCS
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub
