Attribute VB_Name = "Quest_TCP"
Public Sub SendUpdateQuestTo(ByVal index As Long, ByVal QuestNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    
    Set Buffer = New clsBuffer
    
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendQuests(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_QUESTS

        If LenB(Trim$(Quest(i).name)) > 0 Then
            Call SendUpdateQuestTo(index, i)
        End If

    Next

End Sub

Public Sub SendUpdateQuestToAll(ByVal QuestNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    
    Set Buffer = New clsBuffer
    
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData

    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

