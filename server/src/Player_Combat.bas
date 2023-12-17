Attribute VB_Name = "Player_Combat"
' ################################
' ##      Basic Calculations    ##
' ################################

Public Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    If Index > MAX_PLAYERS Then Exit Function
    Select Case Vital
        Case HP
            Select Case GetPlayerClass(Index)
                Case 1 ' Warrior
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 15 + 150
                Case 2 ' Wizard
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 5 + 65
                Case 3 ' Whisperer
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 5 + 65
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 15 + 150
            End Select
        Case MP
            Select Case GetPlayerClass(Index)
                Case 1 ' Warrior
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 5 + 25
                Case 2 ' Wizard
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 30 + 85
                Case 3 ' Whisperer
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 30 + 85
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 5 + 25
            End Select
    End Select
End Function

Public Function GetPlayerVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = 10 '(GetPlayerStat(index, Stats.Willpower) * 0.8) + 6
        Case MP
            i = 10 '(GetPlayerStat(index, Stats.Willpower) / 4) + 12.5
    End Select

    If i < 2 Then i = 2
    GetPlayerVitalRegen = i
End Function

Public Function GetPlayerDamage(ByVal Index As Long) As Long
    Dim weaponNum As Long
    
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipment(Index, Weapon)
        GetPlayerDamage = Item(weaponNum).Data2 + (((Item(weaponNum).Data2 / 100) * 5) * GetPlayerStat(Index, Strength))
    Else
        GetPlayerDamage = 1 + (((0.01) * 5) * GetPlayerStat(Index, Strength))
    End If

End Function

Public Function GetPlayerDefence(ByVal Index As Long) As Long
    Dim Defence As Long, i As Long, ItemNum As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    ' base defence
    For i = 1 To Equipment.Equipment_Count - 1
        If i <> Equipment.Weapon Then
            ItemNum = GetPlayerEquipment(Index, i)
            If ItemNum > 0 Then
                If Item(ItemNum).Data2 > 0 Then
                    Defence = Defence + Item(ItemNum).Data2
                End If
            End If
        End If
    Next
    
    ' divide by 3
    Defence = Defence / 3
    
    ' floor it at 1
    If Defence < 1 Then Defence = 1
    
    ' add in a player's agility
    GetPlayerDefence = Defence + (((Defence / 100) * 2.5) * (GetPlayerStat(Index, Agility) / 2))
End Function

Public Function GetPlayerSpellDamage(ByVal Index As Long, ByVal spellNum As Long) As Long
    Dim Damage As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    ' return damage
    Damage = Spell(spellNum).Vital
    ' 10% modifier
    If Damage <= 0 Then Damage = 1
    GetPlayerSpellDamage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerBlock(ByVal Index As Long) As Boolean
    Dim rate As Long
    Dim rndNum As Long

    CanPlayerBlock = False

    rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerCrit(ByVal Index As Long) As Boolean
    Dim rate As Long
    Dim rndNum As Long

    CanPlayerCrit = False

    rate = GetPlayerStat(Index, Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodge(ByVal Index As Long) As Boolean
    Dim rate As Long
    Dim rndNum As Long

    CanPlayerDodge = False

    rate = GetPlayerStat(Index, Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerParry(ByVal Index As Long) As Boolean
    Dim rate As Long
    Dim rndNum As Long

    CanPlayerParry = False

    rate = GetPlayerStat(Index, Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################
Public Sub TryPlayerAttackNpc(ByVal Index As Long, ByVal mapNpcNum As Long)
Dim blockAmount As Long
Dim npcNum As Long
Dim mapnum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(Index, mapNpcNum) Then
    
        mapnum = GetPlayerMap(Index)
        npcNum = MapNpc(mapnum).Npc(mapNpcNum).Num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(npcNum) Then
            SendActionMsg mapnum, "Dodge!", Pink, 1, (MapNpc(mapnum).Npc(mapNpcNum).X * 32), (MapNpc(mapnum).Npc(mapNpcNum).Y * 32)
            Exit Sub
        End If
        If CanNpcParry(npcNum) Then
            SendActionMsg mapnum, "Parry!", Pink, 1, (MapNpc(mapnum).Npc(mapNpcNum).X * 32), (MapNpc(mapnum).Npc(mapNpcNum).Y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(mapNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        'damage = damage - RAND(1, (Npc(NpcNum).Stat(Stats.Agility) * 2))
        Damage = Damage - RAND((GetNpcDefence(npcNum) / 100) * 10, (GetNpcDefence(npcNum) / 100) * 10)
        ' randomise from 1 to max hit
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(Index, mapNpcNum, Damage)
        Else
            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal Attacker As Long, ByVal mapNpcNum As Long, Optional ByVal isSpell As Boolean = False) As Boolean
    Dim mapnum As Long
    Dim npcNum As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).Npc(mapNpcNum).Num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(Attacker)
    npcNum = MapNpc(mapnum).Npc(mapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).Npc(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
            Exit Function
        End If
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
    
        ' exit out early
        If isSpell Then
             If npcNum > 0 Then
                If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    CanPlayerAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        ' attack speed from weapon
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(Attacker, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If npcNum > 0 And GetTickCount > TempPlayer(Attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    If MapNpc(mapnum).Npc(mapNpcNum).X >= Player(Attacker).X - 1 And MapNpc(mapnum).Npc(mapNpcNum).X <= Player(Attacker).X + 1 Then
                        If MapNpc(mapnum).Npc(mapNpcNum).Y + 1 = Player(Attacker).Y Then
                            If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                                CanPlayerAttackNpc = True
                            ElseIf Npc(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Then
                                ' init conversation if it's friendly
                                If Npc(npcNum).Conv > 0 Then
                                    InitChat Attacker, mapnum, mapNpcNum
                                ElseIf Npc(npcNum).Mission > 0 Then
                                    Call CheckHas_Mission(Attacker, mapNpcNum)
                                End If
                            End If
                        End If
                    End If
                Case DIR_DOWN
                    If MapNpc(mapnum).Npc(mapNpcNum).X >= Player(Attacker).X - 1 And MapNpc(mapnum).Npc(mapNpcNum).X <= Player(Attacker).X + 1 Then
                        If MapNpc(mapnum).Npc(mapNpcNum).Y - 1 = Player(Attacker).Y Then
                            If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                                CanPlayerAttackNpc = True
                            ElseIf Npc(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Then
                                ' init conversation if it's friendly
                                If Npc(npcNum).Conv > 0 Then
                                    InitChat Attacker, mapnum, mapNpcNum
                                ElseIf Npc(npcNum).Mission > 0 Then
                                    Call CheckHas_Mission(Attacker, mapNpcNum)
                                End If
                            End If
                        End If
                    End If
                Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                    If MapNpc(mapnum).Npc(mapNpcNum).X + 1 = Player(Attacker).X Then
                        If MapNpc(mapnum).Npc(mapNpcNum).Y >= Player(Attacker).Y - 1 And MapNpc(mapnum).Npc(mapNpcNum).Y <= Player(Attacker).Y + 1 Then
                            If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                                CanPlayerAttackNpc = True
                            ElseIf Npc(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Then
                                ' init conversation if it's friendly
                                If Npc(npcNum).Conv > 0 Then
                                    InitChat Attacker, mapnum, mapNpcNum
                                ElseIf Npc(npcNum).Mission > 0 Then
                                    Call CheckHas_Mission(Attacker, mapNpcNum)
                                End If
                            End If
                        End If
                    End If
                Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                    If MapNpc(mapnum).Npc(mapNpcNum).X - 1 = Player(Attacker).X Then
                        If MapNpc(mapnum).Npc(mapNpcNum).Y >= Player(Attacker).Y - 1 And MapNpc(mapnum).Npc(mapNpcNum).Y <= Player(Attacker).Y + 1 Then
                            If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                                CanPlayerAttackNpc = True
                            ElseIf Npc(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Then
                                ' init conversation if it's friendly
                                If Npc(npcNum).Conv > 0 Then
                                    InitChat Attacker, mapnum, mapNpcNum
                                ElseIf Npc(npcNum).Mission > 0 Then
                                    Call CheckHas_Mission(Attacker, mapNpcNum)
                                End If
                            End If
                        End If
                    End If
            End Select
        End If
    End If

End Function

Public Sub PlayerAttackNpc(ByVal Attacker As Long, ByVal mapNpcNum As Long, ByVal Damage As Long, Optional ByVal spellNum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim STR As Long
    Dim DEF As Long
    Dim mapnum As Long
    Dim npcNum As Long
    Dim Mission_ID As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(Attacker)
    npcNum = MapNpc(mapnum).Npc(mapNpcNum).Num
    Name = Trim$(Npc(npcNum).Name)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = GetTickCount

    If Damage >= MapNpc(mapnum).Npc(mapNpcNum).Vital(Vitals.HP) Then
    
        SendActionMsg GetPlayerMap(Attacker), "-" & MapNpc(mapnum).Npc(mapNpcNum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(mapnum).Npc(mapNpcNum).X * 32), (MapNpc(mapnum).Npc(mapNpcNum).Y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).Npc(mapNpcNum).X, MapNpc(mapnum).Npc(mapNpcNum).Y
        
        ' send the sound
        If spellNum > 0 Then SendMapSound Attacker, MapNpc(mapnum).Npc(mapNpcNum).X, MapNpc(mapnum).Npc(mapNpcNum).Y, SoundEntity.seSpell, spellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If spellNum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, MapNpc(mapnum).Npc(mapNpcNum).X, MapNpc(mapnum).Npc(mapNpcNum).Y)
            End If
        End If

        ' Calculate exp to give attacker
        exp = Npc(npcNum).exp

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 1
        End If

        ' in party?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker, Npc(npcNum).Level
        Else
            ' no party - keep exp for self
            GivePlayerEXP Attacker, exp, Npc(npcNum).Level
        End If
        
        'Drop the goods if they get it
        For n = 1 To MAX_NPC_DROPS
            If Npc(npcNum).DropItem(n) = 0 Then Exit For
            If Rnd <= Npc(npcNum).DropChance(n) Then
                Call SpawnItem(Npc(npcNum).DropItem(n), Npc(npcNum).DropItemValue(n), mapnum, MapNpc(mapnum).Npc(mapNpcNum).X, MapNpc(mapnum).Npc(mapNpcNum).Y, GetPlayerName(Attacker))
            End If
        Next
        
        ' destroy map npcs
        If Map(mapnum).MapData.Moral = MAP_MORAL_BOSS Then
            If mapNpcNum = Map(mapnum).MapData.BossNpc Then
                ' kill all the other npcs
                For i = 1 To MAX_MAP_NPCS
                    If Map(mapnum).MapData.Npc(i) > 0 Then
                        ' only kill dangerous npcs
                        If Npc(Map(mapnum).MapData.Npc(i)).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(Map(mapnum).MapData.Npc(i)).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                            ' kill!
                            MapNpc(mapnum).Npc(i).Num = 0
                            MapNpc(mapnum).Npc(i).SpawnWait = GetTickCount
                            MapNpc(mapnum).Npc(i).Vital(Vitals.HP) = 0
                            ' send kill command
                            SendNpcDeath mapnum, i
                        End If
                    End If
                Next
            End If
        End If

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapnum).Npc(mapNpcNum).Num = 0
        MapNpc(mapnum).Npc(mapNpcNum).SpawnWait = GetTickCount
        MapNpc(mapnum).Npc(mapNpcNum).Vital(Vitals.HP) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(mapnum).Npc(mapNpcNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapnum).Npc(mapNpcNum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' Check Mission Kill NPC
        Call Check_Mission(Attacker, npcNum)
        
        ' send death to the map
        SendNpcDeath mapnum, mapNpcNum
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = mapnum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).Target = mapNpcNum Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(mapnum).Npc(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).Npc(mapNpcNum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        SendActionMsg mapnum, "-" & Damage, BrightRed, 1, (MapNpc(mapnum).Npc(mapNpcNum).X * 32), (MapNpc(mapnum).Npc(mapNpcNum).Y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).Npc(mapNpcNum).X, MapNpc(mapnum).Npc(mapNpcNum).Y
        
        ' send the sound
        If spellNum > 0 Then SendMapSound Attacker, MapNpc(mapnum).Npc(mapNpcNum).X, MapNpc(mapnum).Npc(mapNpcNum).Y, SoundEntity.seSpell, spellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If spellNum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(mapnum).Npc(mapNpcNum).targetType = 1 ' player
        MapNpc(mapnum).Npc(mapNpcNum).Target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(mapnum).Npc(mapNpcNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapnum).Npc(i).Num = MapNpc(mapnum).Npc(mapNpcNum).Num Then
                    MapNpc(mapnum).Npc(i).Target = Attacker
                    MapNpc(mapnum).Npc(i).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(mapnum).Npc(mapNpcNum).stopRegen = True
        MapNpc(mapnum).Npc(mapNpcNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If spellNum > 0 Then
            If Spell(spellNum).StunDuration > 0 Then StunNPC mapNpcNum, mapnum, spellNum
            ' DoT
            If Spell(spellNum).Duration > 0 Then
                AddDoT_Npc mapnum, mapNpcNum, spellNum, Attacker
            End If
        End If
        
        SendMapNpcVitals mapnum, mapNpcNum
        
        ' set the player's target if they don't have one
        If TempPlayer(Attacker).Target = 0 Then
            TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
            TempPlayer(Attacker).Target = mapNpcNum
            SendTarget Attacker
        End If
    End If

    If spellNum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).AttackTimer = GetTickCount
    End If
End Sub
' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal Attacker As Long, ByVal victim As Long)
Dim blockAmount As Long, npcNum As Long, mapnum As Long, Damage As Long, Defence As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(Attacker, victim) Then
    
        mapnum = GetPlayerMap(Attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(victim) Then
            SendActionMsg mapnum, "Dodge!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If
        If CanPlayerParry(victim) Then
            SendActionMsg mapnum, "Parry!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Defence = GetPlayerDefence(victim)
        If Defence > 0 Then
            Damage = Damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If Damage <= 0 Then Damage = 1
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if can crit
        If CanPlayerCrit(Attacker) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(Attacker, victim, Damage)
        Else
            Call PlayerMsg(Attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal Attacker As Long, ByVal victim As Long, Optional ByVal isSpell As Boolean = False) As Boolean
Dim partynum As Long, i As Long

    If Not isSpell Then
        ' Check attack timer
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If GetTickCount < TempPlayer(Attacker).AttackTimer + Item(GetPlayerEquipment(Attacker, Weapon)).Speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function
    
    ' make sure it's not you
    If victim = Attacker Then
        PlayerMsg Attacker, "Cannot attack yourself.", BrightRed
        Exit Function
    End If
    
    ' check co-ordinates if not spell
    If Not isSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_LEFT
                If Not ((GetPlayerY(victim) = GetPlayerY(Attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_RIGHT
                If Not ((GetPlayerY(victim) = GetPlayerY(Attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_UP_LEFT
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_UP_RIGHT
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN_LEFT
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN_RIGHT
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).MapData.Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < 5 Then
        Call PlayerMsg(Attacker, "You are below level 5, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 5 Then
        Call PlayerMsg(Attacker, GetPlayerName(victim) & " is below level 5, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If
    
    ' make sure not in your party
    partynum = TempPlayer(Attacker).inParty
    If partynum > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partynum).Member(i) > 0 Then
                If victim = Party(partynum).Member(i) Then
                    PlayerMsg Attacker, "Cannot attack party members.", BrightRed
                    Exit Function
                End If
            End If
        Next
    End If

    CanPlayerAttackPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal Attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal spellNum As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        If spellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, spellNum
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        ' Calculate exp to give attacker
        exp = (GetPlayerExp(victim) \ 10)

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 0
        End If

        If exp = 0 Then
            Call PlayerMsg(victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(Attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(victim, GetPlayerExp(victim) - exp)
            SendEXP victim
            Call PlayerMsg(victim, "You lost " & exp & " exp.", BrightRed)
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker, GetPlayerLevel(victim)
            Else
                ' not in party, get exp for self
                GivePlayerEXP Attacker, exp, GetPlayerLevel(victim)
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).Target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).Target = victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If

        Call OnDeath(victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        If spellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, spellNum
        
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If spellNum > 0 Then
            If Spell(spellNum).StunDuration > 0 Then StunPlayer victim, spellNum
            ' DoT
            If Spell(spellNum).Duration > 0 Then
                AddDoT_Player victim, spellNum, Attacker
            End If
        End If
        
        ' change target if need be
        If TempPlayer(Attacker).Target = 0 Then
            TempPlayer(Attacker).targetType = TARGET_TYPE_PLAYER
            TempPlayer(Attacker).Target = victim
            SendTarget Attacker
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = GetTickCount
End Sub

' ############
' ## Spells ##
' ############
Public Sub BufferSpell(ByVal Index As Long, ByVal spellSlot As Long)
    Dim spellNum As Long, mpCost As Long, LevelReq As Long, mapnum As Long, ClassReq As Long
    Dim AccessReq As Long, Range As Long, HasBuffered As Boolean, targetType As Byte, Target As Long
    Dim PlayerProjectileSlot As Long, ProjectileSlot As Long
    
    ' Prevent subscript out of range
    If spellSlot <= 0 Or spellSlot > MAX_PLAYER_SPELLS Then Exit Sub
    
    spellNum = Player(Index).Spell(spellSlot).Spell
    mapnum = GetPlayerMap(Index)
    
    If spellNum <= 0 Or spellNum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(Index, spellNum) Then Exit Sub
    
    ' make sure we're not buffering already
    If TempPlayer(Index).spellBuffer.Spell = spellSlot Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(Index).SpellCD(spellSlot) > GetTickCount Then
        PlayerMsg Index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    mpCost = Spell(spellNum).mpCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < mpCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(spellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(spellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(spellNum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellNum).Type = SPELL_TYPE_PROJECTILE Then
        If Spell(spellNum).Range = 0 Then Exit Sub
        
        If Spell(spellNum).Projectile.Ammo > 0 Then
            If HasItem(Index, Spell(spellNum).Projectile.Ammo) Then
                TakeInvItem Index, Spell(spellNum).Projectile.Ammo, 1
            Else
                PlayerMsg Index, "Suas ferramentas acabaram.", BrightRed
                Exit Sub
            End If
        End If
        TempPlayer(Index).SpellCastType = 0
    Else
        If Spell(spellNum).Range > 0 Then
            ' ranged attack, single target or aoe?
            If Not Spell(spellNum).IsAoE Then
                TempPlayer(Idex).SpellCastType = 2 ' targetted
            Else
                TempPlayer(Idex).SpellCastType = 3 ' targetted aoe
            End If
        Else
            If Not Spell(spellNum).IsAoE Then
                TempPlayer(Idex).SpellCastType = 0 ' self-cast
            Else
                TempPlayer(Idex).SpellCastType = 1 ' self-cast AoE
            End If
        End If
    End If
    
    targetType = TempPlayer(Index).targetType
    Target = TempPlayer(Index).Target
    Range = Spell(spellNum).Range
    HasBuffered = False
    
    Select Case TempPlayer(Index).SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not Target > 0 Then
                PlayerMsg Index, "You do not have a target.", BrightRed
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(Target), GetPlayerY(Target)) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                Else
                    ' go through spell types
                    If Spell(spellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(Index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if beneficial magic then self-cast it instead
                If Spell(spellNum).Type = SPELL_TYPE_HEALHP Or Spell(spellNum).Type = SPELL_TYPE_HEALMP Then
                    Target = Index
                    targetType = TARGET_TYPE_PLAYER
                    HasBuffered = True
                Else
                    ' if have target, check in range
                    If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), MapNpc(mapnum).Npc(Target).X, MapNpc(mapnum).Npc(Target).Y) Then
                        PlayerMsg Index, "Target not in range.", BrightRed
                        HasBuffered = False
                    Else
                        ' go through spell types
                        If Spell(spellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                            HasBuffered = True
                        Else
                            If CanPlayerAttackNpc(Index, Target, True) Then
                                HasBuffered = True
                            End If
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation mapnum, Spell(spellNum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, Index, 1
        TempPlayer(Index).spellBuffer.Spell = spellSlot
        TempPlayer(Index).spellBuffer.Timer = GetTickCount
        TempPlayer(Index).spellBuffer.Target = Target
        TempPlayer(Index).spellBuffer.tType = targetType
        Exit Sub
    Else
        SendClearSpellBuffer Index
    End If
End Sub

Public Sub CastSpell(ByVal Index As Long, ByVal spellSlot As Long, ByVal Target As Long, ByVal targetType As Byte)
    Dim spellNum As Long, mpCost As Long, LevelReq As Long
    Dim mapnum As Long, Vital As Long, DidCast As Boolean, ClassReq As Long
    Dim AccessReq As Long, i As Long, AoE As Long, Range As Long
    Dim vitalType As Byte, increment As Boolean, X As Long, Y As Long
    Dim buffer As clsBuffer, SpellCastType As Long
    Dim ProjectileSlot As Long, PlayerProjectileSlot As Long
    
    DidCast = False

    ' Prevent subscript out of range
    If spellSlot <= 0 Or spellSlot > MAX_PLAYER_SPELLS Then Exit Sub

    spellNum = Player(Index).Spell(spellSlot).Spell
    mapnum = GetPlayerMap(Index)

    ' Make sure player has the spell
    If Not HasSpell(Index, spellNum) Then Exit Sub

    mpCost = Spell(spellNum).mpCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < mpCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(spellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(spellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(spellNum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
'    ' find out what kind of spell it is! self cast, target or AOE
'    If Spell(spellNum).Range > 0 Then
'        ' ranged attack, single target or aoe?
'        If Not Spell(spellNum).IsAoE Then
'            SpellCastType = 2 ' targetted
'        Else
'            SpellCastType = 3 ' targetted aoe
'        End If
'    Else
'        If Not Spell(spellNum).IsAoE Then
'            SpellCastType = 0 ' self-cast
'        Else
'            SpellCastType = 1 ' self-cast AoE
'        End If
'    End If
    
    ' get damage
    Vital = GetPlayerSpellDamage(Index, spellNum)
    
    ' store data
    AoE = Spell(spellNum).RadiusX
    Range = Spell(spellNum).Range
    
    Select Case TempPlayer(Index).SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(spellNum).Type
                Case SPELL_TYPE_HEALHP
                    SpellPlayer_Effect Vitals.HP, True, Index, Vital, spellNum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, Index, Vital, spellNum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    PlayerWarp Index, Spell(spellNum).Map, Spell(spellNum).X, Spell(spellNum).Y
                    SendAnimation GetPlayerMap(Index), Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    DidCast = True
                Case SPELL_TYPE_PROJECTILE
                    SpellPlayer_Projectile Index, spellNum, GetPlayerMap(Index)
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                X = GetPlayerX(Index)
                Y = GetPlayerY(Index)
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
                
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), X, Y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    SendClearSpellBuffer Index
                End If
            End If
            Select Case Spell(spellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> Index Then
                                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                    If isInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(Index, i, True) Then
                                            SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            PlayerAttackPlayer Index, i, Vital, spellNum
                                            DidCast = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).Npc(i).Num > 0 Then
                            If MapNpc(mapnum).Npc(i).Vital(HP) > 0 Then
                                If isInRange(AoE, X, Y, MapNpc(mapnum).Npc(i).X, MapNpc(mapnum).Npc(i).Y) Then
                                    If CanPlayerAttackNpc(Index, i, True) Then
                                        SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc Index, i, Vital, spellNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(spellNum).Type = SPELL_TYPE_HEALHP Then
                        vitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_HEALMP Then
                        vitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        vitalType = Vitals.MP
                        increment = False
                    End If
                    
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                If isInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect vitalType, increment, i, Vital, spellNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
                    
                    If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
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
                
            If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), X, Y) Then
                PlayerMsg Index, "Target not in range.", BrightRed
                SendClearSpellBuffer Index
                Exit Sub
            End If
            
            Select Case Spell(spellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(Index, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                PlayerAttackPlayer Index, Target, Vital, spellNum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(Index, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                PlayerAttackNpc Index, Target, Vital, spellNum
                                DidCast = True
                            End If
                        End If
                    End If
                    
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        vitalType = Vitals.MP
                        increment = False
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_HEALMP Then
                        vitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_HEALHP Then
                        vitalType = Vitals.HP
                        increment = True
                    End If
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(Index, Target, True) Then
                                SpellPlayer_Effect vitalType, increment, Target, Vital, spellNum
                                DidCast = True
                            End If
                        Else
                            SpellPlayer_Effect vitalType, increment, Target, Vital, spellNum
                            DidCast = True
                        End If
                    Else
                        If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(Index, Target, True) Then
                                SpellNpc_Effect vitalType, increment, Target, Vital, spellNum, mapnum
                                DidCast = True
                            End If
                        Else
                            SpellNpc_Effect vitalType, increment, Target, Vital, spellNum, mapnum
                            DidCast = True
                        End If
                    End If
            End Select
    End Select
    
    If DidCast Then
        Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - mpCost)
        Call SendVital(Index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
        
        TempPlayer(Index).SpellCD(spellSlot) = GetTickCount + (Spell(spellNum).CDTime * 1000)
        Call SendCooldown(Index, spellSlot)
        
        ' if has a next rank then increment usage
        SetPlayerSpellUsage Index, spellSlot
    End If
End Sub

Public Sub SetPlayerSpellUsage(ByVal Index As Long, ByVal spellSlot As Long)
    Dim spellNum As Long, i As Long
    spellNum = Player(Index).Spell(spellSlot).Spell
    ' if has a next rank then increment usage
    If Spell(spellNum).NextRank > 0 Then
        If Player(Index).Spell(spellSlot).Uses < Spell(spellNum).NextUses - 1 Then
            Player(Index).Spell(spellSlot).Uses = Player(Index).Spell(spellSlot).Uses + 1
        Else
            If GetPlayerLevel(Index) >= Spell(Spell(spellNum).NextRank).LevelReq Then
                Player(Index).Spell(spellSlot).Spell = Spell(spellNum).NextRank
                Player(Index).Spell(spellSlot).Uses = 0
                PlayerMsg Index, "Your spell has ranked up!", Blue
                ' update hotbar
                For i = 1 To MAX_HOTBAR
                    If Player(Index).Hotbar(i).Slot > 0 Then
                        If Player(Index).Hotbar(i).sType = 2 Then ' spell
                            If Spell(Player(Index).Hotbar(i).Slot).UniqueIndex = Spell(Spell(spellNum).NextRank).UniqueIndex Then
                                Player(Index).Hotbar(i).Slot = Spell(spellNum).NextRank
                                SendHotbar Index
                            End If
                        End If
                    End If
                Next
            Else
                Player(Index).Spell(spellSlot).Uses = Spell(spellNum).NextUses
            End If
        End If
        SendPlayerSpells Index
    End If
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal spellNum As Long)
    Dim sSymbol As String * 1
    Dim colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then colour = BrightGreen
            If Vital = Vitals.MP Then colour = BrightBlue
        Else
            sSymbol = "-"
            colour = Blue
        End If
    
        SendAnimation GetPlayerMap(Index), Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        SendActionMsg GetPlayerMap(Index), sSymbol & Damage, colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        
        ' send the sound
        SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, spellNum
        
        If increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) + Damage
            If Spell(spellNum).Duration > 0 Then
                AddHoT_Player Index, spellNum
            End If
        ElseIf Not increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) - Damage
        End If
        
        ' send update
        SendVital Index, Vital
    End If
End Sub

Public Sub SpellPlayer_Projectile(ByVal Index As Long, spellNum As Long, mapnum As Long)
    Dim targetType As Byte, Target As Long, Range As Byte, i As Long
    Dim xT As Long, yT As Long, Rotate As Long
    
    Dim ProjectileIndex As Integer
    'Get the next open projectile slot
    Do
        ProjectileIndex = ProjectileIndex + 1
        
        'Update LastProjectile if we go over the size of the current array
        If ProjectileIndex > MapProjectile_HighIndex Then
            MapProjectile_HighIndex = ProjectileIndex
            ReDim Preserve MapProjectile(1 To MapProjectile_HighIndex)
            Exit Do
        End If
        
    Loop While MapProjectile(ProjectileIndex).Graphic > 0
    
    targetType = TempPlayer(Index).targetType
    Target = TempPlayer(Index).Target
    Range = Spell(spellNum).Range
    
    With MapProjectile(ProjectileIndex)
        
        ' SE É UMA PROJECTILE
        If Spell(spellNum).Projectile.Speed < 5000 Then
            ' DEFINE OS VALORES INICIAIS
            .X = GetPlayerX(Index)
            .Y = GetPlayerY(Index)
            ' SE TEMOS UMA PROJECTILE DE DANO EM AREA
            If Spell(spellNum).IsAoE Then
                Select Case GetPlayerDir(Index)
                    Case DIR_UP
                        .tX = GetPlayerX(Index)
                        If GetPlayerY(Index) - Spell(spellNum).Range >= 0 Then
                            .tY = GetPlayerY(Index) - Spell(spellNum).Range
                        Else
                            .tY = 0
                        End If
                    Case DIR_DOWN
                        .tX = GetPlayerX(Index)
                        If GetPlayerY(Index) + Spell(spellNum).Range <= Map(mapnum).MapData.MaxY Then
                            .tY = GetPlayerY(Index) + Spell(spellNum).Range
                        Else
                            .tY = Map(mapnum).MapData.MaxY
                        End If
                    Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                        If GetPlayerX(Index) - Spell(spellNum).Range >= 0 Then
                            .tX = GetPlayerX(Index) - Spell(spellNum).Range
                        Else
                            .tX = 0
                        End If
                        .tY = GetPlayerY(Index)
                    Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                        If GetPlayerX(Index) + Spell(spellNum).Range <= Map(mapnum).MapData.MaxX Then
                            .tX = GetPlayerX(Index) + Spell(spellNum).Range
                        Else
                            .tX = Map(mapnum).MapData.MaxX
                        End If
                        .tY = GetPlayerY(Index)
                End Select
            ' DEFINIR A POSIÇÃO DO ALVO
            Else
                ' SE TEMOS UM ALVO
                If Target > 0 Then
                    ' SE É UM ALVO DO TIPO PLAYER
                    If targetType = TARGET_TYPE_PLAYER Then
                        ' SE ESTÁ FORA DE ALCANCE
                        If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(Target), GetPlayerY(Target)) Then
                            
                        ' SE ESTÁ DENTRO DO ALCANCE
                        Else
                            .tX = GetPlayerX(TempPlayer(Index).Target)
                            .tY = GetPlayerY(TempPlayer(Index).Target)
                        End If
                    ' SE É UM ALVO DO TIPO NPC
                    ElseIf targetType = TARGET_TYPE_NPC Then
                        ' SE ESTÁ FORA DA ALCANCE
                        If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), MapNpc(mapnum).Npc(Target).X, MapNpc(mapnum).Npc(Target).Y) Then
                            yT = MapNpc(mapnum).Npc(Target).Y
                            xT = MapNpc(mapnum).Npc(Target).X
                            Do
                                ' Up left
                                If GetPlayerY(Index) < yT And GetPlayerX(Index) < xT Then
                                    yT = yT - 1
                                    xT = xT - 1
                                End If
                                    
                                ' Up right
                                If GetPlayerY(Index) < yT And GetPlayerX(Index) > xT Then
                                    yT = yT - 1
                                    xT = xT + 1
                                End If
                                    
                                ' Down left
                                If GetPlayerY(Index) > yT And GetPlayerX(Index) < xT Then
                                    yT = yT + 1
                                    xT = xT - 1
                                End If
                                    
                                ' Down right
                                If GetPlayerY(Index) > yT And GetPlayerX(Index) > xT Then
                                    yT = yT + 1
                                    xT = xT + 1
                                End If
                                    
                                ' Up
                                If GetPlayerY(Index) < yT Then
                                    yT = yT - 1
                                End If
                                    
                                ' Down
                                If GetPlayerY(Index) > yT Then
                                    yT = yT + 1
                                End If
                                    
                                ' left
                                If GetPlayerX(Index) < xT Then
                                    xT = xT - 1
                                End If
                                    
                                ' right
                                If GetPlayerX(Index) > xT Then
                                    xT = xT + 1
                                End If
                                
                            Loop Until isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), xT, yT)
                            .tX = xT
                            .tY = yT
                        ' SE ESTÁ DENTRO DO ALCANCE
                        Else
                            .tX = MapNpc(mapnum).Npc(TempPlayer(Index).Target).X
                            .tY = MapNpc(mapnum).Npc(TempPlayer(Index).Target).Y
                        End If
                    End If
                ' SE NÃO TEMOS UM ALVO DEFINIR O ALVO NO ALCANCE MÁXIMO
                Else
                    Select Case GetPlayerDir(Index)
                        Case DIR_UP
                            .tX = GetPlayerX(Index)
                            If GetPlayerY(Index) - Spell(spellNum).Range >= 0 Then
                                .tY = GetPlayerY(Index) - Spell(spellNum).Range
                            Else
                                .tY = 0
                            End If
                        Case DIR_DOWN
                            .tX = GetPlayerX(Index)
                            If GetPlayerY(Index) + Spell(spellNum).Range <= Map(mapnum).MapData.MaxY Then
                                .tY = GetPlayerY(Index) + Spell(spellNum).Range
                            Else
                                .tY = Map(mapnum).MapData.MaxY
                            End If
                        Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                            If GetPlayerX(Index) - Spell(spellNum).Range >= 0 Then
                                .tX = GetPlayerX(Index) - Spell(spellNum).Range
                            Else
                                .tX = 0
                            End If
                            .tY = GetPlayerY(Index)
                        Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                            If GetPlayerX(Index) + Spell(spellNum).Range <= Map(mapnum).MapData.MaxX Then
                                .tX = GetPlayerX(Index) + Spell(spellNum).Range
                            Else
                                .tX = Map(mapnum).MapData.MaxX
                            End If
                            .tY = GetPlayerY(Index)
                    End Select
                End If
            End If 'If Spell(spellNum).IsAoE Then
            
            If Spell(spellNum).IsAoE Then
                Select Case GetPlayerDir(Index)
                    Case DIR_UP
                        .xTargetAoE = (GetPlayerX(Index) - Int(Spell(spellNum).DirectionAoE(DIR_UP + 1).X / 2)) * PIC_X
                        .yTargetAoE = (GetPlayerY(Index) - 1) * PIC_Y
                        If Spell(spellNum).Projectile.RecuringDamage Then
                            .Duration = Spell(spellNum).DirectionAoE(DIR_UP + 1).Y
                        Else
                            .Duration = 1
                        End If
                    Case DIR_DOWN
                        .xTargetAoE = (GetPlayerX(Index) - Int(Spell(spellNum).DirectionAoE(DIR_DOWN + 1).X / 2)) * PIC_X
                        .yTargetAoE = (GetPlayerY(Index) + 1) * PIC_Y
                        If Spell(spellNum).Projectile.RecuringDamage Then
                            .Duration = Spell(spellNum).DirectionAoE(DIR_UP + 1).Y
                        Else
                            .Duration = 1
                        End If
                    Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                        .xTargetAoE = (GetPlayerX(Index) - 1) * PIC_X
                        .yTargetAoE = (GetPlayerY(Index) - Int(Spell(spellNum).DirectionAoE(DIR_LEFT + 1).Y) / 2) * PIC_Y
                        If Spell(spellNum).Projectile.RecuringDamage Then
                            .Duration = Spell(spellNum).DirectionAoE(DIR_UP + 1).X
                        Else
                            .Duration = 1
                        End If
                    Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                        .xTargetAoE = (GetPlayerX(Index) + 1) * PIC_X
                        .yTargetAoE = (GetPlayerY(Index) - Int(Spell(spellNum).DirectionAoE(DIR_RIGHT + 1).Y / 2)) * PIC_Y
                        If Spell(spellNum).Projectile.RecuringDamage Then
                            .Duration = Spell(spellNum).DirectionAoE(DIR_UP + 1).X
                        Else
                            .Duration = 1
                        End If
                End Select
            End If
            
            ' DEFINE O ANGULO INICIAL DE ROTAÇÃO
            .Rotate = Engine_GetAngle(.X, .Y, .tX, .tY)
            ' DEFINE A VELOCIDADE DE ROTAÇÃO
            .RotateSpeed = Spell(spellNum).Projectile.Rotation
            
            ' DEFINE O LADO QUE O PLAYER DEVE VIRAR ANTES DE SOLTAR A SKILL
            If .Rotate >= 315 And .Rotate <= 360 Then
                Call SetPlayerDir(Index, DIR_UP)
            ElseIf .Rotate >= 0 And .Rotate <= 45 Then
                Call SetPlayerDir(Index, DIR_UP)
            ElseIf .Rotate >= 225 And .Rotate <= 315 Then
                Call SetPlayerDir(Index, DIR_LEFT)
            ElseIf .Rotate >= 135 And .Rotate <= 225 Then
                Call SetPlayerDir(Index, DIR_DOWN)
            ElseIf .Rotate >= 45 And .Rotate <= 135 Then
                Call SetPlayerDir(Index, DIR_RIGHT)
            End If
            
            Dim buffer As clsBuffer
            
            Set buffer = New clsBuffer
            buffer.WriteLong SPlayerDir
            buffer.WriteLong Index
            buffer.WriteLong GetPlayerDir(Index)
            Call SendDataToMap(mapnum, buffer.ToArray())
            Set buffer = Nothing
        ' SE É UMA TRAP
        Else
            If Spell(spellNum).IsAoE Then
                Select Case GetPlayerDir(Index)
                    Case DIR_UP
                        .xTargetAoE = (GetPlayerX(Index) - Int(Spell(spellNum).DirectionAoE(DIR_UP + 1).X / 2)) * PIC_X
                        .yTargetAoE = (GetPlayerY(Index) - 1) * PIC_Y
                    Case DIR_DOWN
                        .xTargetAoE = (GetPlayerX(Index) - Int(Spell(spellNum).DirectionAoE(DIR_DOWN + 1).X / 2)) * PIC_X
                        .yTargetAoE = (GetPlayerY(Index) + 1) * PIC_Y
                    Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                        .xTargetAoE = (GetPlayerX(Index) - 1) * PIC_X
                        .yTargetAoE = (GetPlayerY(Index) - Int(Spell(spellNum).DirectionAoE(DIR_LEFT + 1).Y) / 2) * PIC_Y
                    Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                        .xTargetAoE = (GetPlayerX(Index) + 1) * PIC_X
                        .yTargetAoE = (GetPlayerY(Index) - Int(Spell(spellNum).DirectionAoE(DIR_RIGHT + 1).Y / 2)) * PIC_Y
                End Select
            End If
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    .X = GetPlayerX(Index)
                    If GetPlayerY(Index) - 1 < 0 Then
                        Exit Sub
                    Else
                        .Y = GetPlayerY(Index) - 1
                    End If
                Case DIR_DOWN
                    .X = GetPlayerX(Index)
                    If GetPlayerY(Index) + 1 > Map(mapnum).MapData.MaxY Then
                        Exit Sub
                    Else
                        .Y = GetPlayerY(Index) + 1
                    End If
                Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                    If GetPlayerX(Index) - 1 < 0 Then
                        Exit Sub
                    Else
                        .X = GetPlayerX(Index) - 1
                    End If
                    .Y = GetPlayerY(Index)
                Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                    If GetPlayerX(Index) + 1 > Map(mapnum).MapData.MaxX Then
                        Exit Sub
                    Else
                        .X = GetPlayerX(Index) + 1
                    End If
                    .Y = GetPlayerY(Index)
            End Select
            
            .Duration = Spell(spellNum).Projectile.Despawn
        End If
        
        ' DEFINE OS DADOS DO DONO DA PROJECTILE
        .Owner = Index
        .OwnerType = TARGET_TYPE_PLAYER
        ' DEFINE A DIREÇÃO DA PROJECTILE
        .direction = GetPlayerDir(Index)
        ' DEFINE O GRÁFICO DO PROJECTILE
        .Graphic = Spell(spellNum).Projectile.Graphic
        ' DEFINE A VELOCIDADE DA PROJECTILE
        .Speed = Spell(spellNum).Projectile.Speed
        ' ALTERA AS POSIÇÕES DA X,Y E TARGET X,Y
        .X = .X * PIC_X
        .Y = .Y * PIC_Y
        .tX = .tX * PIC_X
        .tY = .tY * PIC_Y
        .spellNum = spellNum
        ' DEFINE OS OFFSET DE X E Y PARA EXIBIR NA POSIÇÃO CERTA NO MAPA
        For i = 1 To 4
            .ProjectileOffset(i).X = Spell(spellNum).Projectile.ProjectileOffset(i).X
            .ProjectileOffset(i).Y = Spell(spellNum).Projectile.ProjectileOffset(i).Y
        Next
        
        Call SendProjectile(mapnum, ProjectileIndex, Spell(spellNum).IsDirectional)
        If .Speed >= 5000 Then
            .Duration = .Duration + tick
        End If
    End With
End Sub

Public Sub AddDoT_Player(ByVal Index As Long, ByVal spellNum As Long, ByVal Caster As Long)
    Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(i)
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

Public Sub AddHoT_Player(ByVal Index As Long, ByVal spellNum As Long)
    Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).HoT(i)
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

Public Sub HandleDoT_Player(ByVal Index As Long, ByVal dotNum As Long)
    With TempPlayer(Index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, Index, True) Then
                    PlayerAttackPlayer .Caster, Index, GetPlayerSpellDamage(.Caster, .Spell)
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
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

Public Sub HandleHoT_Player(ByVal Index As Long, ByVal hotNum As Long)
    With TempPlayer(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg Player(Index).Map, "+" & GetPlayerSpellDamage(.Caster, .Spell), BrightGreen, ACTIONMSG_SCROLL, Player(Index).X * 32, Player(Index).Y * 32
                Player(Index).Vital(Vitals.HP) = Player(Index).Vital(Vitals.HP) + GetPlayerSpellDamage(.Caster, .Spell)
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
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

Public Sub StunPlayer(ByVal Index As Long, ByVal spellNum As Long)
    ' check if it's a stunning spell
    If Spell(spellNum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(Index).StunDuration = Spell(spellNum).StunDuration
        TempPlayer(Index).StunTimer = GetTickCount
        ' send it to the index
        SendStunned Index
        ' tell him he's stunned
        PlayerMsg Index, "You have been stunned.", BrightRed
    End If
End Sub
