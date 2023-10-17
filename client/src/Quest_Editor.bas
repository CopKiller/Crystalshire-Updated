Attribute VB_Name = "Quest_Editor"
Option Explicit

Public Mission_Changed(1 To MAX_MISSIONS) As Boolean

' ////////////////////
' // Mission Editor //
' ////////////////////
Public Sub MissionEditorInit()
    Dim i As Long, N As Long

    If frmEditor_Quest.visible = False Then Exit Sub
    EditorIndex = frmEditor_Quest.lstIndex.ListIndex + 1

    With frmEditor_Quest
        ' Set Default
        .scrlItemNum.max = MAX_ITEMS
        .scrlItemAmount.max = 32000
        
        ' Set Attributes
        .txtName.text = Trim$(Mission(EditorIndex).Name)
        
        ' Set Mission Type
        If Mission(EditorIndex).Type = MissionType.TypeTalk Then
            .frmTalkQuest.visible = True
            .frmKillQuest.visible = False
            .frmCollectQuest.visible = False
            .cmbTalkNPC.ListIndex = Mission(EditorIndex).TalkNPC
            
        ElseIf Mission(EditorIndex).Type = MissionType.TypeKill Then
            .frmKillQuest.visible = True
            .frmCollectQuest.visible = False
            .frmTalkQuest.visible = False
            ' Chain Mission
            .cmbKillNPC.Clear
            .cmbKillNPC.AddItem "None"
            For i = 1 To MAX_NPCS
                .cmbKillNPC.AddItem i & ": " & Trim$(Npc(EditorIndex).Name)
            Next
            .cmbKillNPC.ListIndex = Mission(EditorIndex).KillNPC
            .scrlKillAmount.value = Mission(EditorIndex).KillNPCAmount
            
        ElseIf Mission(EditorIndex).Type = MissionType.TypeCollect Then
            .frmKillQuest.visible = False
            .frmCollectQuest.visible = True
            .frmTalkQuest.visible = False
            .cmbCollectItem.ListIndex = Mission(EditorIndex).CollectItem
            .scrlCollectAmount.value = Mission(EditorIndex).CollectItemAmount
        End If
        
        ' Set Mission Repeatable
        If Mission(EditorIndex).Repeatable = 1 Then
            .optRepeatableYes.value = True
            .optRepeatableNo.value = False
        Else
            .optRepeatableYes.value = False
            .optRepeatableNo.value = True
        End If
        
        .txtDialogue.text = Mission(EditorIndex).Dialogue
        
        ' Chain Mission
        .cmbPreviousQuest.Clear
        .cmbPreviousQuest.AddItem "None"
        For i = 1 To MAX_MISSIONS
            .cmbPreviousQuest.AddItem i & ": " & Trim$(Mission(EditorIndex).Name)
        Next
        .cmbPreviousQuest.ListIndex = Mission(EditorIndex).PreviousMissionComplete

        ' Message
        .txtIncomplete = Mission(EditorIndex).Incomplete
        .txtCompleted.text = Mission(EditorIndex).Completed
        
        ' Reward
        For i = 1 To 5
            .scrlItemNum.value = Mission(EditorIndex).RewardItem(i).ItemNum
            .scrlItemAmount.value = Mission(EditorIndex).RewardItem(i).ItemAmount
        Next
        .scrlRewardNum.value = 1
        .scrlRewardExperience.value = Mission(EditorIndex).RewardExperience
        .lblRewardExperience.caption = "Reward Experience: " & Mission(EditorIndex).RewardExperience
    End With

    Mission_Changed(EditorIndex) = True
End Sub

Public Sub MissionEditorOk()
    Dim i As Long

    For i = 1 To MAX_MISSIONS

        If Mission_Changed(i) Then
            Call SendSaveMission(i)
        End If

    Next

    Unload frmEditor_Quest
    Editor = 0
    ClearChanged_Mission
End Sub

Public Sub MissionEditorCancel()
    Editor = 0
    Unload frmEditor_Quest
    ClearChanged_Mission
    ClearMissions
    SendRequestMissions
End Sub

Public Sub ClearChanged_Mission()
    ZeroMemory Mission_Changed(1), MAX_MISSIONS * 2 ' 2 = boolean length
End Sub
