Attribute VB_Name = "Quest_Handle"
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' MISSION EDITORES
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Public Sub HandleMissionEditor()
    Dim i As Long

    With frmEditor_Quest
        Editor = EDITOR_MISSION
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_MISSIONS
            .lstIndex.AddItem i & ": " & Trim$(Mission(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        MissionEditorInit
    End With

End Sub

Public Sub HandleUpdateMission(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Dim MissionSize As Long
    Dim MissionData() As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    N = buffer.ReadLong
    MissionSize = LenB(Mission(N))
    
    ReDim MissionData(MissionSize - 1)
    MissionData = buffer.ReadBytes(MissionSize)
    
    ClearMission N
    CopyMemory ByVal VarPtr(Mission(N)), ByVal VarPtr(MissionData(0)), MissionSize
    
    buffer.Flush: Set buffer = Nothing
End Sub
