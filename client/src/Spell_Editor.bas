Attribute VB_Name = "Spell_Editor"
Option Explicit

Public Spell_Changed(1 To MAX_SPELLS) As Boolean
' //////////////////
' // Spell Editor //
' //////////////////
Public Sub SpellEditorCopy()
    CopyMemory ByVal VarPtr(tmpSpell), ByVal VarPtr(Spell(EditorIndex)), LenB(Spell(EditorIndex))
End Sub

Public Sub SpellEditorPaste()
    CopyMemory ByVal VarPtr(Spell(EditorIndex)), ByVal VarPtr(tmpSpell), LenB(tmpSpell)
    SpellEditorInit
    frmEditor_Spell.txtName_Validate False
End Sub

Public Sub IsTrap()
    If frmEditor_Spell.chkTrap.Value = 1 Then
        frmEditor_Spell.scrlProjectileSpeed.enabled = False
        frmEditor_Spell.scrlProjectileRange.enabled = False
        frmEditor_Spell.scrlProjectileRotation.enabled = False
        frmEditor_Spell.scrlDurationProjectile.enabled = True
        Spell(EditorIndex).Projectile.Speed = 5000
        Spell(EditorIndex).Range = 1
        Spell(EditorIndex).Projectile.Rotation = 0
    Else
        frmEditor_Spell.scrlProjectileSpeed.enabled = True
        frmEditor_Spell.scrlProjectileRange.enabled = True
        frmEditor_Spell.scrlProjectileRotation.enabled = True
        frmEditor_Spell.scrlDurationProjectile.enabled = False
        Spell(EditorIndex).Projectile.Despawn = 0
    End If
End Sub

Public Sub SpellEditorInit()
    Dim i As Long
    Dim SoundSet As Boolean

    If frmEditor_Spell.visible = False Then Exit Sub
    EditorIndex = frmEditor_Spell.lstIndex.ListIndex + 1

    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If

    ' add the array to the combo
    frmEditor_Spell.cmbSound.Clear
    frmEditor_Spell.cmbSound.AddItem "None."

    For i = 1 To UBound(soundCache)
        frmEditor_Spell.cmbSound.AddItem soundCache(i)
    Next

    ' finished populating
    With frmEditor_Spell
        ' set max values
        .scrlAnimCast.max = MAX_ANIMATIONS
        .scrlAnim.max = MAX_ANIMATIONS
        .scrlRange.max = MAX_BYTE
        .scrlMap.max = MAX_MAPS
        .scrlNext.max = MAX_SPELLS
        ' set values
        .txtName.text = Trim$(Spell(EditorIndex).Name)
        .cmbType.ListIndex = Spell(EditorIndex).Type
        .scrlMP.Value = Spell(EditorIndex).MPCost
        .scrlLevel.Value = Spell(EditorIndex).LevelReq
        .scrlAccess.Value = Spell(EditorIndex).AccessReq
        ' build class combo
            .cmbClass.Clear
            .cmbClass.AddItem "None"
    
            For i = 1 To Max_Classes
                .cmbClass.AddItem Trim$(Class(i).Name)
            Next
    
            .cmbClass.ListIndex = 0
            .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
        ' End build class combo
        .scrlCast.Value = Spell(EditorIndex).CastTime
        .scrlCool.Value = Spell(EditorIndex).CDTime
        .scrlStun.Value = Spell(EditorIndex).StunDuration
        .scrlIcon.Value = Spell(EditorIndex).icon
        
        If .cmbType.ListIndex = SPELL_TYPE_PROJECTILE Then
            ' Definições
            .scrlProjectilePic.max = CountProjectile
            .fraProjectile.visible = True
            .fraSpellData.visible = False
            ' Sets
            .scrlProjectileSpeed.Value = Spell(EditorIndex).Projectile.Speed
            
            If .scrlProjectileSpeed.Value = 5000 Then
                .chkTrap.Value = 1
            Else
                .chkTrap.Value = 0
            End If
            
            Call IsTrap
            .scrlDamageProjectile.Value = Spell(EditorIndex).Vital
            
            If Spell(EditorIndex).Projectile.RecuringDamage Then
                .chkRecuringDamage.Value = 1
            Else
                .chkRecuringDamage.Value = 0
            End If
            
            If Spell(EditorIndex).IsAoE Then
                .chkProjectileAoE.Value = 1
            Else
                .chkProjectileAoE.Value = 0
            End If
            
            .scrlProjectilePic.Value = Spell(EditorIndex).Projectile.Graphic
            
            
            If Spell(EditorIndex).IsDirectional Then
                .chkDirectionalProjectile.Value = 1
            Else
                .chkDirectionalProjectile.Value = 0
            End If
            
            .scrlProjectileRange.Value = Spell(EditorIndex).Range
            .scrlProjectileRotation.Value = Spell(EditorIndex).Projectile.Rotation
            .scrlProjectileAmmo.Value = Spell(EditorIndex).Projectile.Ammo
            .scrlDurationProjectile.Value = Int(Spell(EditorIndex).Projectile.Despawn / 100)
            .scrlProjectileAnimOnHit.Value = Spell(EditorIndex).Projectile.AnimOnHit
            
            .lblX0.caption = "X Offset: 0"
            .lblY0.caption = "Y Offset: 0"
            .cmbDirection.ListIndex = 0
            .scrlOffsetProjectileX = 0
            .scrlOffsetProjectileY = 0
            .scrlOffsetProjectileX.enabled = False
            .scrlOffsetProjectileY.enabled = False
            
            .lblAoEProjectile.caption = "No AoE damage settings"
            .scrlProjectileRadiusX = 0
            .scrlProjectileRadiusY = 0
            .scrlProjectileRadiusX.enabled = False
            .scrlProjectileRadiusY.enabled = False
        Else
            .fraSpellData.visible = True
            .fraProjectile.visible = False
            
            .scrlMap.Value = Spell(EditorIndex).Map
            .scrlX.Value = Spell(EditorIndex).X
            .scrlY.Value = Spell(EditorIndex).Y
            .scrlDir.Value = Spell(EditorIndex).Dir
            .scrlVital.Value = Spell(EditorIndex).Vital
            .scrlDuration.Value = Spell(EditorIndex).Duration
            .scrlInterval.Value = Spell(EditorIndex).Interval
            .scrlRange.Value = Spell(EditorIndex).Range
            
            If Spell(EditorIndex).IsAoE Then
                .chkAOE.Value = 1
            Else
                .chkAOE.Value = 0
            End If
    
            If Spell(EditorIndex).IsDirectional Then
                .chkDirectional.Value = 1
            Else
                .chkDirectional.Value = 0
            End If
            
            .cmbAoEDirection.ListIndex = 0
            .scrlRadiusX.Value = Spell(EditorIndex).RadiusX
            .scrlRadiusY.Value = Spell(EditorIndex).RadiusY
            .scrlAnimCast.Value = Spell(EditorIndex).CastAnim
            .scrlAnim.Value = Spell(EditorIndex).SpellAnim
        End If

        .txtDesc.text = Trim$(Spell(EditorIndex).Desc)
        .scrlIndex.Value = Spell(EditorIndex).UniqueIndex
        .scrlNext.Value = Spell(EditorIndex).NextRank
        .scrlUses.Value = Spell(EditorIndex).NextUses

        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then

            For i = 0 To .cmbSound.ListCount

                If .cmbSound.list(i) = Trim$(Spell(EditorIndex).sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If

            Next

            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If

    End With

    Spell_Changed(EditorIndex) = True
End Sub

Public Sub SpellEditorOk()
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Spell_Changed(i) Then
            Call SendSaveSpell(i)
        End If

    Next

    Unload frmEditor_Spell
    Editor = 0
    ClearChanged_Spell
End Sub

Public Sub SpellEditorCancel()
    Editor = 0
    Unload frmEditor_Spell
    ClearChanged_Spell
    ClearSpells
    SendRequestSpells
End Sub

Public Sub ClearChanged_Spell()
    ZeroMemory Spell_Changed(1), MAX_SPELLS * 2 ' 2 = boolean length
End Sub
