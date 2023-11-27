Attribute VB_Name = "Quest_TCP"
Public Sub SendRequestEditMission()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditMission
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendSaveMission(ByVal N As Long)
    Dim buffer As clsBuffer
    Dim MissionSize As Long
    Dim MissionData() As Byte
    Set buffer = New clsBuffer
    MissionSize = LenB(Mission(N))
    ReDim MissionData(MissionSize - 1)
    CopyMemory MissionData(0), ByVal VarPtr(Mission(N)), MissionSize
    buffer.WriteLong CSaveMission
    buffer.WriteLong N
    buffer.WriteBytes MissionData
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendRequestMissions()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestMissions
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendAcceptMissionRequest(ByVal Index_Offer)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptMissionRequest
    buffer.WriteLong inOffer(Index_Offer)
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendDeclineMissionRequest()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineMissionRequest
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

