Attribute VB_Name = "Quest_Handle"
' :::::::::::::::::::::::::::::
' :: Request edit Mission packet ::
' :::::::::::::::::::::::::::::
Public Sub HandleRequestEditMission(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SMissionEditor
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub HandleRequestMissions(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendMissions(index)
End Sub

' :::::::::::::::::::::
' :: Save Mission packet ::
' :::::::::::::::::::::
Public Sub HandleSaveMission(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim Buffer As clsBuffer
    Dim MissionSize As Long
    Dim MissionData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    N = Buffer.ReadLong 'CLng(Parse(1))

    If N < 0 Or N > MAX_MISSIONS Then
        Exit Sub
    End If

    ' Update the Mission
    MissionSize = LenB(Mission(N))
    ReDim MissionData(MissionSize - 1)
    MissionData = Buffer.ReadBytes(MissionSize)
    CopyMemory ByVal VarPtr(Mission(N)), ByVal VarPtr(MissionData(0)), MissionSize
    Buffer.Flush: Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateMissionToAll(N)
    Call SaveMission(N)
    Call AddLog(GetPlayerName(index) & " saved Mission #" & N & ".", ADMIN_LOG)
End Sub
