Attribute VB_Name = "Client_UpdateWindow"
Public Sub Window_QuestUpdate()
    Dim X As Long
    Dim IsActive As Boolean
    
    With Windows(GetWindowIndex("winPlayerQuests"))
            For X = 1 To MAX_PLAYER_MISSIONS
                If Player(MyIndex).Mission(X).ID <> 0 Then
                    IsActive = True
                    .Controls(GetControlIndex("winPlayerQuests", "btnMission" & X)).visible = True
                    .Controls(GetControlIndex("winPlayerQuests", "btnMission" & X)).text = Trim$(Mission(Player(MyIndex).Mission(X).ID).Name)
                End If
            Next
            If IsActive Then
                btnMissionActive = 1
            Else
                btnMissionActive = 0
            End If
            If btnMissionActive <> 0 Then
                .Controls(GetControlIndex("winPlayerQuests", "lblDescription")).text = Trim$(Mission(Player(MyIndex).Mission(btnMissionActive).ID).Description)
                Select Case Mission(Player(MyIndex).Mission(btnMissionActive).ID).Type
                    Case MissionType.TypeCollect
                        .Controls(GetControlIndex("winPlayerQuests", "lblGoal")).text = "You must collect " & Trim$(Item(Mission(Player(MyIndex).Mission(btnMissionActive).ID).CollectItem).Name) & " (" & Player(MyIndex).Mission(btnMissionActive).Count & "/" & Mission(Player(MyIndex).Mission(btnMissionActive).ID).CollectItemAmount & ")"
                    Case MissionType.TypeKill
                        .Controls(GetControlIndex("winPlayerQuests", "lblGoal")).text = "You must kill " & Trim$(Npc(Mission(Player(MyIndex).Mission(btnMissionActive).ID).KillNPC).Name) & " (" & Player(MyIndex).Mission(btnMissionActive).Count & "/" & Mission(Player(MyIndex).Mission(btnMissionActive).ID).KillNPCAmount & ")"
                    Case MissionType.TypeTalk
                        .Controls(GetControlIndex("winPlayerQuests", "lblGoal")).text = "You should talk to " & Trim$(Npc(Mission(Player(MyIndex).Mission(btnMissionActive).ID).KillNPC).Name)
                End Select
            Else
                For X = 1 To MAX_PLAYER_MISSIONS
                    .Controls(GetControlIndex("winPlayerQuests", "btnMission" & X)).visible = False
                Next
                .Controls(GetControlIndex("winPlayerQuests", "lblDescription")).text = ""
                .Controls(GetControlIndex("winPlayerQuests", "lblGoal")).text = ""
            End If
        End With
End Sub
