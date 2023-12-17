Attribute VB_Name = "Client_UpdateWindow"
Public Sub UpdateWindowOffer(ByVal Index_Offer As Long)
    Dim i As Long
    ' gui stuff
    With Windows(GetWindowIndex("winOffer"))
        ' set main text
        If Index_Offer <> 0 Then
            .Controls(GetControlIndex("winOffer", "picBGOffer" & Index_Offer)).visible = True
            .Controls(GetControlIndex("winOffer", "picOfferBG" & Index_Offer)).visible = True
            .Controls(GetControlIndex("winOffer", "lblTitleOffer" & Index_Offer)).visible = True
            .Controls(GetControlIndex("winOffer", "btnAccept" & Index_Offer)).visible = True
            .Controls(GetControlIndex("winOffer", "btnRecuse" & Index_Offer)).visible = True
            Select Case inOfferType(Index_Offer)
                Case Offers.Offer_Type_Mission
                    .Controls(GetControlIndex("winOffer", "lblTitleOffer" & Index_Offer)).text = "Quest: " & Trim$(Mission(inOffer(Index_Offer)).Name) & "?"
                Case Offers.Offer_Type_Party
                    .Controls(GetControlIndex("winOffer", "lblTitleOffer" & Index_Offer)).text = inOfferInvite(Index_Offer) & " has invited you to a party."
                Case Offers.Offer_Type_Trade
                    .Controls(GetControlIndex("winOffer", "lblTitleOffer" & Index_Offer)).text = inOfferInvite(Index_Offer) & "  has invited you to trade."
            End Select
            ShowWindow GetWindowIndex("winOffer")
        Else
            For i = 1 To MAX_OFFER
                .Controls(GetControlIndex("winOffer", "picBGOffer" & i)).visible = False
                .Controls(GetControlIndex("winOffer", "picOfferBG" & i)).visible = False
                .Controls(GetControlIndex("winOffer", "lblTitleOffer" & i)).visible = False
                .Controls(GetControlIndex("winOffer", "btnAccept" & i)).visible = False
                .Controls(GetControlIndex("winOffer", "btnRecuse" & i)).visible = False
            Next
            HideWindow GetWindowIndex("winOffer")
        End If
    End With
End Sub

Public Sub Window_QuestButtonUpdate()
    Dim x As Long
    Dim isActive As Boolean
    
    With Windows(GetWindowIndex("winPlayerQuests"))
        For x = 1 To MAX_PLAYER_MISSIONS
            If Player(MyIndex).Mission(x).ID <> 0 Then
                isActive = True
                .Controls(GetControlIndex("winPlayerQuests", "btnMission" & x)).visible = True
                .Controls(GetControlIndex("winPlayerQuests", "btnMission" & x)).text = Trim$(Mission(Player(MyIndex).Mission(x).ID).Name)
            End If
        Next
        If isActive Then
            Button_MissionActive = 1
            Window_QuestLabelUpdate
        Else
            Button_MissionActive = 0
            For x = 1 To MAX_PLAYER_MISSIONS
                If Player(MyIndex).Mission(x).ID = 0 Then
                    .Controls(GetControlIndex("winPlayerQuests", "btnMission" & x)).visible = False
                End If
            Next
        End If
    End With
End Sub

Public Sub Window_QuestLabelUpdate()
    With Windows(GetWindowIndex("winPlayerQuests"))
        If Button_MissionActive <> 0 Then
            If Player(MyIndex).Mission(Button_MissionActive).ID <= 0 Then Exit Sub
            .Controls(GetControlIndex("winPlayerQuests", "lblDescription")).text = Trim$(Mission(Player(MyIndex).Mission(Button_MissionActive).ID).Description)
            Select Case Mission(Player(MyIndex).Mission(Button_MissionActive).ID).Type
                Case MissionType.Mission_TypeCollect
                If Mission(Player(MyIndex).Mission(Button_MissionActive).ID).CollectItemAmount = 0 Then Exit Sub
                    .Controls(GetControlIndex("winPlayerQuests", "lblGoal")).text = "You must collect " & Trim$(Item(Mission(Player(MyIndex).Mission(Button_MissionActive).ID).CollectItem).Name) & " (" & Player(MyIndex).Mission(Button_MissionActive).Count & "/" & Mission(Player(MyIndex).Mission(Button_MissionActive).ID).CollectItemAmount & ")"
                Case MissionType.Mission_TypeKill
                    If Mission(Player(MyIndex).Mission(Button_MissionActive).ID).KillNPC = 0 Then Exit Sub
                    .Controls(GetControlIndex("winPlayerQuests", "lblGoal")).text = "You must kill " & Trim$(Npc(Mission(Player(MyIndex).Mission(Button_MissionActive).ID).KillNPC).Name) & " (" & Player(MyIndex).Mission(Button_MissionActive).Count & "/" & Mission(Player(MyIndex).Mission(Button_MissionActive).ID).KillNPCAmount & ")"
                Case MissionType.Mission_TypeTalk
                    If Mission(Player(MyIndex).Mission(Button_MissionActive).ID).TalkNPC = 0 Then Exit Sub
                    .Controls(GetControlIndex("winPlayerQuests", "lblGoal")).text = "You should talk to " & Trim$(Npc(Mission(Player(MyIndex).Mission(Button_MissionActive).ID).TalkNPC).Name)
            End Select
            
            .Controls(GetControlIndex("winPlayerQuests", "lblEXP")).text = Str(Mission(Player(MyIndex).Mission(Button_MissionActive).ID).RewardExperience) & " EXP"
        Else
            For x = 1 To MAX_PLAYER_MISSIONS
                .Controls(GetControlIndex("winPlayerQuests", "btnMission" & x)).visible = False
            Next
            .Controls(GetControlIndex("winPlayerQuests", "lblDescription")).text = ""
            .Controls(GetControlIndex("winPlayerQuests", "lblGoal")).text = ""
        End If
    End With
End Sub
