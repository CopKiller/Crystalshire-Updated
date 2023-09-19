Attribute VB_Name = "Quest_Handle"
' :::::::::::::::::::::::::::::
' :: Request edit Quest packet ::
' :::::::::::::::::::::::::::::
Public Sub HandleRequestEditQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SQuestEditor
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub HandleRequestQuests(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendQuests(index)
End Sub

' :::::::::::::::::::::
' :: Save Quest packet ::
' :::::::::::::::::::::
Public Sub HandleSaveQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim QuestNum As Long
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    QuestNum = Buffer.ReadLong

    ' Prevent hacking
    If QuestNum < 0 Or QuestNum > MAX_QUESTS Then
        Exit Sub
    End If

    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    QuestData = Buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(Quest(QuestNum)), ByVal VarPtr(QuestData(0)), QuestSize
    ' Save it
    Call SendUpdateQuestToAll(QuestNum)
    Call SaveQuest(QuestNum)
    Call AddLog(GetPlayerName(index) & " saved Quest #" & QuestNum & ".", ADMIN_LOG)
End Sub

