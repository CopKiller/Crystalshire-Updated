Attribute VB_Name = "Quest_TCP"
Public Sub SendUpdateMissionTo(ByVal index As Long, ByVal N As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim MissionSize As Long
    Dim MissionData() As Byte
    
    Set Buffer = New clsBuffer
    MissionSize = LenB(Mission(N))
    
    ReDim MissionData(MissionSize - 1)
    CopyMemory MissionData(0), ByVal VarPtr(Mission(N)), MissionSize
    
    Buffer.WriteLong SUpdateMission
    Buffer.WriteLong N
    Buffer.WriteBytes MissionData
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendMissions(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_MISSIONS

        If LenB(Trim$(Mission(i).Name)) > 0 Then
            Call SendUpdateMissionTo(index, i)
        End If

    Next

End Sub

Public Sub SendUpdateMissionToAll(ByVal N As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim MissionSize As Long
    Dim MissionData() As Byte
    
    Set Buffer = New clsBuffer
    MissionSize = LenB(Mission(N))
    
    ReDim MissionData(MissionSize - 1)
    CopyMemory MissionData(0), ByVal VarPtr(Mission(N)), MissionSize
    
    Buffer.WriteLong SUpdateMission
    Buffer.WriteLong N
    Buffer.WriteBytes MissionData
    
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

