Attribute VB_Name = "Quest_Database"
Public Sub ClearMission(ByVal Index As Long)
    Mission(Index) = EmptyMission
    Mission(Index).Name = vbNullString
    Mission(Index).Incomplete = vbNullString
    Mission(Index).Completed = vbNullString
End Sub

Public Sub ClearMissions()
    Dim i As Long

    For i = 1 To MAX_MISSIONS
        Call ClearMission(i)
    Next

End Sub

