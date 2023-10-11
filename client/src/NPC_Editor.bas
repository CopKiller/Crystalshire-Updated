Attribute VB_Name = "NPC_Editor"
Option Explicit

Public NPC_Changed(1 To MAX_NPCS) As Boolean
' Temp event storage
Public tmpNPC As NpcRec

' ////////////////
' // Npc Editor //
' ////////////////
Public Sub NpcEditorInit()
    Dim i As Long
    Dim SoundSet As Boolean

    If frmEditor_NPC.visible = False Then Exit Sub
    EditorIndex = frmEditor_NPC.lstIndex.ListIndex + 1

    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If

    ' add the array to the combo
    frmEditor_NPC.cmbSound.Clear
    frmEditor_NPC.cmbSound.AddItem "None."

    For i = 1 To UBound(soundCache)
        frmEditor_NPC.cmbSound.AddItem soundCache(i)
    Next

    ' finished populating
    With frmEditor_NPC
        .scrlDrop.max = MAX_NPC_DROPS
        .scrlSpell.max = MAX_NPC_SPELLS
        .txtName.text = Trim$(Npc(EditorIndex).Name)
        .txtAttackSay.text = Trim$(Npc(EditorIndex).AttackSay)

        If Npc(EditorIndex).sprite < 0 Or Npc(EditorIndex).sprite > .scrlSprite.max Then Npc(EditorIndex).sprite = 0
        .scrlSprite.value = Npc(EditorIndex).sprite
        .txtSpawnSecs.text = CStr(Npc(EditorIndex).SpawnSecs)
        .cmbBehaviour.ListIndex = Npc(EditorIndex).Behaviour
        ' Build cmbClassReq
            .cmbMission.Clear
            .cmbMission.AddItem "None"
    
            For i = 1 To MAX_MISSIONS
                .cmbMission.AddItem Trim$(Mission(i).Name)
            Next
    
            .cmbMission.ListIndex = Npc(EditorIndex).Mission
        
        .scrlRange.value = Npc(EditorIndex).Range
        .txtHP.text = Npc(EditorIndex).HP
        .txtEXP.text = Npc(EditorIndex).EXP
        .txtLevel.text = Npc(EditorIndex).Level
        .txtDamage.text = Npc(EditorIndex).Damage
        .scrlConv.value = Npc(EditorIndex).Conv
        .scrlAnimation.value = Npc(EditorIndex).Animation

        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then

            For i = 0 To .cmbSound.ListCount

                If .cmbSound.list(i) = Trim$(Npc(EditorIndex).sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If

            Next

            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If

        For i = 1 To Stats.Stat_Count - 1
            .scrlStat(i).value = Npc(EditorIndex).Stat(i)
        Next

        ' show 1 data
        .scrlDrop.value = 1
        .scrlSpell.value = 1
    End With

    NPC_Changed(EditorIndex) = True
End Sub

Public Sub NpcEditorOk()
    Dim i As Long

    For i = 1 To MAX_NPCS

        If NPC_Changed(i) Then
            Call SendSaveNpc(i)
        End If

    Next

    Unload frmEditor_NPC
    Editor = 0
    ClearChanged_NPC
End Sub

Sub NpcEditorCopy()
    CopyMemory ByVal VarPtr(tmpNPC), ByVal VarPtr(Npc(EditorIndex)), LenB(Npc(EditorIndex))
End Sub

Sub NpcEditorPaste()
    CopyMemory ByVal VarPtr(Npc(EditorIndex)), ByVal VarPtr(tmpNPC), LenB(tmpNPC)
    NpcEditorInit
    frmEditor_NPC.txtName_Validate False
End Sub

Public Sub NpcEditorCancel()
    Editor = 0
    Unload frmEditor_NPC
    ClearChanged_NPC
    ClearNpcs
    SendRequestNPCS
End Sub

Public Sub ClearChanged_NPC()
    ZeroMemory NPC_Changed(1), MAX_NPCS * 2 ' 2 = boolean length
End Sub
