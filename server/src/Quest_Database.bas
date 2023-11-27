Attribute VB_Name = "Quest_Database"
' **********
' ** Missions **
' **********
Public Sub SaveMission(ByVal n As Long)
    Dim filename As String
    Dim i As Long, x As Long, f As Long

    filename = App.Path & "\data\missions\Mission" & n & ".dat"
    f = FreeFile

    Open filename For Binary As #f
    With Mission(n)
        Put #f, , .Name
        Put #f, , .Type
        Put #f, , .Repeatable
        Put #f, , .Description
        
        Put #f, , .KillNPC
        Put #f, , .KillNPCAmount
        
        Put #f, , .CollectItem
        Put #f, , .CollectItemAmount
        
        Put #f, , .TalkNPC
        
        Put #f, , .PreviousMissionComplete
        
        Put #f, , .Incomplete
        Put #f, , .Completed
        
        For i = 1 To 5
            Put #f, , .RewardItem(i).ItemNum
            Put #f, , .RewardItem(i).ItemAmount
        Next
        
        Put #f, , .RewardExperience
        
    End With
    Close #f
End Sub

Public Sub SaveMissions()
    Dim i As Long

    For i = 1 To MAX_MISSIONS
        Call SaveMission(i)
    Next

End Sub

Public Sub CheckMissions()
    Dim i As Long

    For i = 1 To MAX_MISSIONS
        If Not FileExist(App.Path & "\data\missions\Mission" & i & ".dat") Then
            Call SaveMission(i)
        End If
    Next

End Sub

Public Sub LoadMissions()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Dim sLen As Long

    Call CheckMissions

    For i = 1 To MAX_MISSIONS
        filename = App.Path & "\data\missions\Mission" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Mission(i)
        Close #f
    Next

End Sub

Public Sub ClearMission(ByVal index As Long)
    Mission(index) = EmptyMission
    Mission(index).Name = vbNullString
    Mission(index).Description = vbNullString
    Mission(index).Incomplete = vbNullString
    Mission(index).Completed = vbNullString
End Sub

Public Sub ClearMissions()
    Dim i As Long

    For i = 1 To MAX_MISSIONS
        Call ClearMission(i)
    Next
End Sub
