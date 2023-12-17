Attribute VB_Name = "NPC_Combat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetNpcSpellDamage(ByVal npcNum As Long, ByVal spellNum As Long) As Long
Dim Damage As Long

    ' Check for subscript out of range
    If npcNum <= 0 Or npcNum > MAX_NPCS Then Exit Function
    
    ' return damage
    Damage = Spell(spellNum).Vital
    ' 10% modifier
    If Damage <= 0 Then Damage = 1
    GetNpcSpellDamage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
End Function

Function GetNpcMaxVital(ByVal npcNum As Long, ByVal Vital As Vitals) As Long
    Dim X As Long

    ' Prevent subscript out of range
    If npcNum <= 0 Or npcNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            GetNpcMaxVital = Npc(npcNum).HP
        Case MP
            GetNpcMaxVital = 30 + (Npc(npcNum).Stat(Intelligence) * 10) + 2
    End Select

End Function

Function GetNpcVitalRegen(ByVal npcNum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    'Prevent subscript out of range
    If npcNum <= 0 Or npcNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (Npc(npcNum).Stat(Stats.Willpower) * 0.8) + 6
        Case MP
            i = (Npc(npcNum).Stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetNpcVitalRegen = i

End Function

Function GetNpcDamage(ByVal npcNum As Long) As Long
    ' return the calculation
    GetNpcDamage = Npc(npcNum).Damage + (((Npc(npcNum).Damage / 100) * 5) * Npc(npcNum).Stat(Stats.Strength))
End Function

Function GetNpcDefence(ByVal npcNum As Long) As Long
Dim Defence As Long
    
    ' base defence
    Defence = 2
    
    ' add in a player's agility
    GetNpcDefence = Defence + (((Defence / 100) * 2.5) * (Npc(npcNum).Stat(Stats.Agility) / 2))
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanNpcBlock(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcBlock = False

    rate = 0
    ' TODO : make it based on shield lol
End Function

Public Function CanNpcCrit(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcCrit = False

    rate = Npc(npcNum).Stat(Stats.Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcCrit = True
    End If
End Function

Public Function CanNpcDodge(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodge = False

    rate = Npc(npcNum).Stat(Stats.Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcParry(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParry = False

    rate = Npc(npcNum).Stat(Stats.Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal mapNpcNum As Long, ByVal Index As Long)
Dim mapnum As Long, npcNum As Long, blockAmount As Long, Damage As Long, Defence As Long

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(mapNpcNum, Index) Then
        mapnum = GetPlayerMap(Index)
        npcNum = MapNpc(mapnum).Npc(mapNpcNum).Num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(Index) Then
            SendActionMsg mapnum, "Dodge!", Pink, 1, (Player(Index).X * 32), (Player(Index).Y * 32)
            Exit Sub
        End If
        If CanPlayerParry(Index) Then
            SendActionMsg mapnum, "Parry!", Pink, 1, (Player(Index).X * 32), (Player(Index).Y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(npcNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(Index)
        Damage = Damage - blockAmount
        
        ' take away armour
        Defence = GetPlayerDefence(Index)
        If Defence > 0 Then
            Damage = Damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If Damage <= 0 Then Damage = 1
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if crit hit
        If CanNpcCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (MapNpc(mapnum).Npc(mapNpcNum).X * 32), (MapNpc(mapnum).Npc(mapNpcNum).Y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(mapNpcNum, Index, Damage)
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal mapNpcNum As Long, ByVal Index As Long, Optional ByVal isSpell As Boolean = False) As Boolean
    Dim mapnum As Long
    Dim npcNum As Long
    Dim PlayerX As Long, PlayerY As Long

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index)).Npc(mapNpcNum).Num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(Index)
    npcNum = MapNpc(mapnum).Npc(mapNpcNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).Npc(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If
    
    ' exit out early if it's a spell
    If isSpell Then
        If IsPlaying(Index) Then
            If npcNum > 0 Then
                CanNpcAttackPlayer = True
                Exit Function
            End If
        End If
    End If
    
    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(mapnum).Npc(mapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If
    MapNpc(mapnum).Npc(mapNpcNum).AttackTimer = GetTickCount

    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If npcNum > 0 Then

                        ' Check if at same coordinates
            Select Case MapNpc(mapnum).Npc(mapNpcNum).Dir
                Case DIR_UP
                    PlayerX = GetPlayerX(Index)
                    PlayerY = GetPlayerY(Index) + 1
                    
                    If PlayerX >= MapNpc(mapnum).Npc(mapNpcNum).X - 1 And PlayerX <= MapNpc(mapnum).Npc(mapNpcNum).X + 1 Then
                        If PlayerY = MapNpc(mapnum).Npc(mapNpcNum).Y Then
                            If MapNpc(mapnum).Npc(mapNpcNum).Dir <> DIR_UP Then
                                Call NpcDir(mapnum, mapNpcNum, DIR_UP)
                            End If
                            CanNpcAttackPlayer = True
                        End If
                    End If
                Case DIR_DOWN
                    PlayerX = GetPlayerX(Index)
                    PlayerY = GetPlayerY(Index) - 1
                    
                    If PlayerX >= MapNpc(mapnum).Npc(mapNpcNum).X - 1 And PlayerX <= MapNpc(mapnum).Npc(mapNpcNum).X + 1 Then
                        If PlayerY = MapNpc(mapnum).Npc(mapNpcNum).Y Then
                            If MapNpc(mapnum).Npc(mapNpcNum).Dir <> DIR_DOWN Then
                                Call NpcDir(mapnum, mapNpcNum, DIR_DOWN)
                            End If
                            CanNpcAttackPlayer = True
                        End If
                    End If
                
                Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                    PlayerX = GetPlayerX(Index) + 1
                    PlayerY = GetPlayerY(Index)
                    
                    If PlayerX = MapNpc(mapnum).Npc(mapNpcNum).X Then
                        If PlayerY >= MapNpc(mapnum).Npc(mapNpcNum).Y - 1 And PlayerY <= MapNpc(mapnum).Npc(mapNpcNum).Y + 1 Then
                            If MapNpc(mapnum).Npc(mapNpcNum).Dir <> DIR_LEFT Then
                                Call NpcDir(mapnum, mapNpcNum, DIR_LEFT)
                            End If
                            CanNpcAttackPlayer = True
                        End If
                    End If
                
                Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                    PlayerX = GetPlayerX(Index) - 1
                    PlayerY = GetPlayerY(Index)
                    
                    If PlayerX = MapNpc(mapnum).Npc(mapNpcNum).X Then
                        If PlayerY >= MapNpc(mapnum).Npc(mapNpcNum).Y - 1 And PlayerY <= MapNpc(mapnum).Npc(mapNpcNum).Y + 1 Then
                            If MapNpc(mapnum).Npc(mapNpcNum).Dir <> DIR_RIGHT Then
                                Call NpcDir(mapnum, mapNpcNum, DIR_RIGHT)
                            End If
                            CanNpcAttackPlayer = True
                        End If
                    End If
            
            End Select
        End If
    End If
End Function

Sub NpcAttackPlayer(ByVal mapNpcNum As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal spellNum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim mapnum As Long
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or IsPlaying(victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(victim)).Npc(mapNpcNum).Num <= 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(victim)
    Name = Trim$(Npc(MapNpc(mapnum).Npc(mapNpcNum).Num).Name)
    
    ' Send this packet so they can see the npc attacking
    Set buffer = New clsBuffer
    buffer.WriteLong SNpcAttack
    buffer.WriteLong mapNpcNum
    
    SendDataToMap mapnum, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(mapnum).Npc(mapNpcNum).stopRegen = True
    MapNpc(mapnum).Npc(mapNpcNum).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        If spellNum > 0 Then
            SendMapSound victim, MapNpc(mapnum).Npc(mapNpcNum).X, MapNpc(mapnum).Npc(mapNpcNum).Y, SoundEntity.seSpell, spellNum
        Else
            SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(mapnum).Npc(mapNpcNum).Num
        End If
        
        ' send animation
        If Not overTime Then
            If spellNum = 0 Then Call SendAnimation(mapnum, Npc(MapNpc(mapnum).Npc(mapNpcNum).Num).Animation, GetPlayerX(victim), GetPlayerY(victim))
        End If
        
        ' kill player
        KillPlayer victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & Name, BrightRed)

        ' Set NPC target to 0
        MapNpc(mapnum).Npc(mapNpcNum).Target = 0
        MapNpc(mapnum).Npc(mapNpcNum).targetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        
        ' send the sound
        If spellNum > 0 Then
            SendMapSound victim, MapNpc(mapnum).Npc(mapNpcNum).X, MapNpc(mapnum).Npc(mapNpcNum).Y, SoundEntity.seSpell, spellNum
        Else
            SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(mapnum).Npc(mapNpcNum).Num
        End If
        
        ' send animation
        If Not overTime Then
            If spellNum = 0 Then Call SendAnimation(mapnum, Npc(MapNpc(GetPlayerMap(victim)).Npc(mapNpcNum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, victim)
        End If
        
        ' if stunning spell, stun the npc
        If spellNum > 0 Then
            If Spell(spellNum).StunDuration > 0 Then StunPlayer victim, spellNum
            ' DoT
            If Spell(spellNum).Duration > 0 Then
                ' TODO: Add Npc vs Player DOTs
            End If
        End If
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(mapnum).Npc(mapNpcNum).Num
        
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
    End If

End Sub

' ############
' ## Spells ##
' ############

Public Sub NpcBufferSpell(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal npcSpellSlot As Long)
Dim spellNum As Long, mpCost As Long, Range As Long, HasBuffered As Boolean, targetType As Byte, Target As Long, SpellCastType As Long, i As Long

    ' prevent rte9
    If npcSpellSlot <= 0 Or npcSpellSlot > MAX_NPC_SPELLS Then Exit Sub
    
    With MapNpc(mapnum).Npc(mapNpcNum)
        ' set the spell number
        spellNum = Npc(.Num).Spell(npcSpellSlot)
        
        ' prevent rte9
        If spellNum <= 0 Or spellNum > MAX_SPELLS Then Exit Sub
        
        ' make sure we're not already buffering
        If .spellBuffer.Spell > 0 Then Exit Sub
        
        ' see if cooldown as finished
        If .SpellCD(npcSpellSlot) > GetTickCount Then Exit Sub
        
        ' Set the MP Cost
        mpCost = Spell(spellNum).mpCost
        
        ' have they got enough mp?
        If .Vital(Vitals.MP) < mpCost Then Exit Sub
        
        ' find out what kind of spell it is! self cast, target or AOE
        If Spell(spellNum).Range > 0 Then
            ' ranged attack, single target or aoe?
            If Not Spell(spellNum).IsAoE Then
                SpellCastType = 2 ' targetted
            Else
                SpellCastType = 3 ' targetted aoe
            End If
        Else
            If Not Spell(spellNum).IsAoE Then
                SpellCastType = 0 ' self-cast
            Else
                SpellCastType = 1 ' self-cast AoE
            End If
        End If
        
        targetType = .targetType
        Target = .Target
        Range = Spell(spellNum).Range
        HasBuffered = False
        
        ' make sure on the map
        If GetPlayerMap(Target) <> mapnum Then Exit Sub
        
        Select Case SpellCastType
            Case 0, 1 ' self-cast & self-cast AOE
                HasBuffered = True
            Case 2, 3 ' targeted & targeted AOE
                ' if it's a healing spell then heal a friend
                If Spell(spellNum).Type = SPELL_TYPE_HEALHP Then
                    ' find a friend who needs healing
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).Npc(i).Num > 0 Then
                            If MapNpc(mapnum).Npc(i).Vital(Vitals.HP) < Npc(MapNpc(mapnum).Npc(i).Num).HP Then
                                targetType = TARGET_TYPE_NPC
                                Target = i
                                HasBuffered = True
                            End If
                        End If
                    Next
                Else
                    ' check if have target
                    If Not Target > 0 Then Exit Sub
                    ' make sure it's a player
                    If targetType = TARGET_TYPE_PLAYER Then
                        ' if have target, check in range
                        If Not isInRange(Range, .X, .Y, GetPlayerX(Target), GetPlayerY(Target)) Then
                            Exit Sub
                        Else
                            If CanNpcAttackPlayer(mapNpcNum, Target, True) Then
                                HasBuffered = True
                            End If
                        End If
                    End If
                End If
        End Select
        
        If HasBuffered Then
            SendAnimation mapnum, Spell(spellNum).CastAnim, 0, 0, TARGET_TYPE_NPC, mapNpcNum
            .spellBuffer.Spell = npcSpellSlot
            .spellBuffer.Timer = GetTickCount
            .spellBuffer.Target = Target
            .spellBuffer.tType = targetType
        End If
    End With
End Sub

Public Sub NpcCastSpell(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal spellSlot As Long, ByVal Target As Long, ByVal targetType As Long)
Dim spellNum As Long, mpCost As Long, Vital As Long, DidCast As Boolean, i As Long, AoE As Long, Range As Long, vitalType As Byte, increment As Boolean, X As Long, Y As Long, SpellCastType As Long

    DidCast = False
    
    ' rte9
    If spellSlot <= 0 Or spellSlot > MAX_NPC_SPELLS Then Exit Sub
    
    With MapNpc(mapnum).Npc(mapNpcNum)
        ' cache spell num
        spellNum = Npc(.Num).Spell(spellSlot)
        
        ' cache mp cost
        mpCost = Spell(spellNum).mpCost
        
        ' make sure still got enough mp
        If .Vital(Vitals.MP) < mpCost Then Exit Sub
        
        ' find out what kind of spell it is! self cast, target or AOE
        If Spell(spellNum).Range > 0 Then
            ' ranged attack, single target or aoe?
            If Not Spell(spellNum).IsAoE Then
                SpellCastType = 2 ' targetted
            Else
                SpellCastType = 3 ' targetted aoe
            End If
        Else
            If Not Spell(spellNum).IsAoE Then
                SpellCastType = 0 ' self-cast
            Else
                SpellCastType = 1 ' self-cast AoE
            End If
        End If
        
        ' get damage
        Vital = GetNpcSpellDamage(.Num, spellNum) 'GetPlayerSpellDamage(index, spellNum)
        
        ' store data
        AoE = Spell(spellNum).RadiusX
        Range = Spell(spellNum).Range
        
        Select Case SpellCastType
            Case 0 ' self-cast target
                Select Case Spell(spellNum).Type
                    Case SPELL_TYPE_HEALHP
                        SpellNpc_Effect Vitals.HP, True, mapNpcNum, Vital, spellNum, mapnum
                        DidCast = True
                    Case SPELL_TYPE_HEALMP
                        SpellNpc_Effect Vitals.MP, True, mapNpcNum, Vital, spellNum, mapnum
                        DidCast = True
                End Select
            Case 1, 3 ' self-cast AOE & targetted AOE
                If SpellCastType = 1 Then
                    X = .X
                    Y = .Y
                ElseIf SpellCastType = 3 Then
                    If targetType = 0 Then Exit Sub
                    If Target = 0 Then Exit Sub
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        X = GetPlayerX(Target)
                        Y = GetPlayerY(Target)
                    Else
                        X = MapNpc(mapnum).Npc(Target).X
                        Y = MapNpc(mapnum).Npc(Target).Y
                    End If
                    
                    If Not isInRange(Range, .X, .Y, X, Y) Then Exit Sub
                End If
                Select Case Spell(spellNum).Type
                    Case SPELL_TYPE_DAMAGEHP
                        For i = 1 To Player_HighIndex
                            If IsPlaying(i) Then
                                If GetPlayerMap(i) = mapnum Then
                                    If isInRange(AoE, .X, .Y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanNpcAttackPlayer(mapNpcNum, i, True) Then
                                            SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            NpcAttackPlayer mapNpcNum, i, Vital, spellNum
                                            DidCast = True
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP
                        If Spell(spellNum).Type = SPELL_TYPE_HEALHP Then
                            vitalType = Vitals.HP
                            increment = True
                        ElseIf Spell(spellNum).Type = SPELL_TYPE_HEALMP Then
                            vitalType = Vitals.MP
                            increment = True
                        End If
                        
                        If Spell(spellNum).Type = SPELL_TYPE_HEALHP Or Spell(spellNum).Type = SPELL_TYPE_HEALMP Then
                            For i = 1 To MAX_MAP_NPCS
                                If MapNpc(mapnum).Npc(i).Num > 0 Then
                                    If MapNpc(mapnum).Npc(i).Vital(HP) > 0 Then
                                        If isInRange(AoE, X, Y, MapNpc(mapnum).Npc(i).X, MapNpc(mapnum).Npc(i).Y) Then
                                            SpellNpc_Effect vitalType, increment, i, Vital, spellNum, mapnum
                                            DidCast = True
                                        End If
                                    End If
                                End If
                            Next
                        End If
                End Select
            Case 2 ' targetted
                If targetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
                
                If targetType = TARGET_TYPE_PLAYER Then
                    X = GetPlayerX(Target)
                    Y = GetPlayerY(Target)
                Else
                    X = MapNpc(mapnum).Npc(Target).X
                    Y = MapNpc(mapnum).Npc(Target).Y
                End If
                    
                If Not isInRange(Range, .X, .Y, X, Y) Then Exit Sub
                
                Select Case Spell(spellNum).Type
                    Case SPELL_TYPE_DAMAGEHP
                        If targetType = TARGET_TYPE_PLAYER Then
                            If CanNpcAttackPlayer(mapNpcNum, Target, True) Then
                                If Vital > 0 Then
                                    SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                    NpcAttackPlayer mapNpcNum, Target, Vital, spellNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Case SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                        If Spell(spellNum).Type = SPELL_TYPE_HEALMP Then
                            vitalType = Vitals.MP
                            increment = True
                        ElseIf Spell(spellNum).Type = SPELL_TYPE_HEALHP Then
                            vitalType = Vitals.HP
                            increment = True
                        End If
                        
                        If targetType = TARGET_TYPE_NPC Then
                            SpellNpc_Effect vitalType, increment, Target, Vital, spellNum, mapnum
                            DidCast = True
                        End If
                End Select
        End Select
        
        If DidCast Then
            .Vital(Vitals.MP) = .Vital(Vitals.MP) - mpCost
            .SpellCD(spellSlot) = GetTickCount + (Spell(spellNum).CDTime * 1000)
        End If
    End With
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal spellNum As Long, ByVal mapnum As Long)
Dim sSymbol As String * 1
Dim colour As Long
Dim npcNum As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then colour = BrightGreen
            If Vital = Vitals.MP Then colour = BrightBlue
        Else
            sSymbol = "-"
            colour = Blue
        End If
    
        SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Index
        SendActionMsg mapnum, sSymbol & Damage, colour, ACTIONMSG_SCROLL, MapNpc(mapnum).Npc(Index).X * 32, MapNpc(mapnum).Npc(Index).Y * 32
        
        ' send the sound
        SendMapSound Index, MapNpc(mapnum).Npc(Index).X, MapNpc(mapnum).Npc(Index).Y, SoundEntity.seSpell, spellNum
        
        npcNum = MapNpc(mapnum).Npc(Index).Num
        If increment Then
            MapNpc(mapnum).Npc(Index).Vital(Vital) = MapNpc(mapnum).Npc(Index).Vital(Vital) + Damage
            ' make sure doesn't go over max
            With MapNpc(mapnum).Npc(Index)
                If .Vital(Vital) > GetNpcMaxVital(npcNum, Vital) Then
                    .Vital(Vital) = GetNpcMaxVital(npcNum, Vital)
                End If
            End With
            If Spell(spellNum).Duration > 0 Then
                AddHoT_Npc mapnum, Index, spellNum
            End If
        ElseIf Not increment Then
            MapNpc(mapnum).Npc(Index).Vital(Vital) = MapNpc(mapnum).Npc(Index).Vital(Vital) - Damage
        End If
    End If
End Sub

Public Sub AddDoT_Npc(ByVal mapnum As Long, ByVal Index As Long, ByVal spellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(mapnum).Npc(Index).DoT(i)
            If .Spell = spellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal mapnum As Long, ByVal Index As Long, ByVal spellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(mapnum).Npc(Index).HoT(i)
            If .Spell = spellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Npc(ByVal mapnum As Long, ByVal Index As Long, ByVal dotNum As Long)
    With MapNpc(mapnum).Npc(Index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, Index, True) Then
                    PlayerAttackNpc .Caster, Index, GetPlayerSpellDamage(.Caster, .Spell), , True
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Npc(ByVal mapnum As Long, ByVal Index As Long, ByVal hotNum As Long)
Dim npcNum As Long

    With MapNpc(mapnum).Npc(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg mapnum, "+" & GetPlayerSpellDamage(.Caster, .Spell), BrightGreen, ACTIONMSG_SCROLL, MapNpc(mapnum).Npc(Index).X * 32, MapNpc(mapnum).Npc(Index).Y * 32
                MapNpc(mapnum).Npc(Index).Vital(Vitals.HP) = MapNpc(mapnum).Npc(Index).Vital(Vitals.HP) + GetPlayerSpellDamage(.Caster, .Spell)
                ' make sure not over max
                npcNum = MapNpc(mapnum).Npc(Index).Num
                If MapNpc(mapnum).Npc(Index).Vital(Vitals.HP) > GetNpcMaxVital(npcNum, Vitals.HP) Then
                    MapNpc(mapnum).Npc(Index).Vital(Vitals.HP) = GetNpcMaxVital(npcNum, Vitals.HP)
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub StunNPC(ByVal Index As Long, ByVal mapnum As Long, ByVal spellNum As Long)
    ' check if it's a stunning spell
    If Spell(spellNum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(mapnum).Npc(Index).StunDuration = Spell(spellNum).StunDuration
        MapNpc(mapnum).Npc(Index).StunTimer = GetTickCount
    End If
End Sub
