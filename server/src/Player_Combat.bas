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
    Dim I As Long

    ' Prevent subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            I = 10 '(GetPlayerStat(index, Stats.Willpower) * 0.8) + 6
        Case MP
            I = 10 '(GetPlayerStat(index, Stats.Willpower) / 4) + 12.5
    End Select

    If I < 2 Then I = 2
    GetPlayerVitalRegen = I
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
    Dim Defence As Long, I As Long, ItemNum As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    ' base defence
    For I = 1 To Equipment.Equipment_Count - 1
        If I <> Equipment.Weapon Then
            ItemNum = GetPlayerEquipment(Index, I)
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
    Dim damage As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    ' return damage
    damage = Spell(spellNum).Vital
    ' 10% modifier
    If damage <= 0 Then damage = 1
    GetPlayerSpellDamage = RAND(damage - ((damage / 100) * 10), damage + ((damage / 100) * 10))
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
Dim MapNum As Long
Dim damage As Long

    damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(Index, mapNpcNum) Then
    
        MapNum = GetPlayerMap(Index)
        npcNum = MapNpc(MapNum).Npc(mapNpcNum).Num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(npcNum) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (MapNpc(MapNum).Npc(mapNpcNum).x * 32), (MapNpc(MapNum).Npc(mapNpcNum).y * 32)
            Exit Sub
        End If
        If CanNpcParry(npcNum) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (MapNpc(MapNum).Npc(mapNpcNum).x * 32), (MapNpc(MapNum).Npc(mapNpcNum).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        damage = GetPlayerDamage(Index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(mapNpcNum)
        damage = damage - blockAmount
        
        ' take away armour
        'damage = damage - RAND(1, (Npc(NpcNum).Stat(Stats.Agility) * 2))
        damage = damage - RAND((GetNpcDefence(npcNum) / 100) * 10, (GetNpcDefence(npcNum) / 100) * 10)
        ' randomise from 1 to max hit
        damage = RAND(damage - ((damage / 100) * 10), damage + ((damage / 100) * 10))
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(Index) Then
            damage = damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
            
        If damage > 0 Then
            Call PlayerAttackNpc(Index, mapNpcNum, damage)
        Else
            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal attacker As Long, ByVal mapNpcNum As Long, Optional ByVal isSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim npcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(attacker)).Npc(mapNpcNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(attacker)
    npcNum = MapNpc(MapNum).Npc(mapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If IsPlaying(attacker) Then
    
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
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(attacker, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If npcNum > 0 And getTime > TempPlayer(attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(attacker)
                Case DIR_UP
                    ' Define cordenadas à frente
                    NpcX = MapNpc(MapNum).Npc(mapNpcNum).x
                    NpcY = MapNpc(MapNum).Npc(mapNpcNum).y + 1
                    
                    If NpcX >= GetPlayerX(attacker) - 1 And NpcX <= GetPlayerX(attacker) + 1 Then
                        If NpcY = GetPlayerY(attacker) Then
                            If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                                CanPlayerAttackNpc = True
                            ElseIf Npc(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Then
                                ' init conversation if it's friendly
                                InitChat attacker, MapNum, mapNpcNum
                            End If
                        End If
                    End If
                Case DIR_DOWN
                    NpcX = MapNpc(MapNum).Npc(mapNpcNum).x
                    NpcY = MapNpc(MapNum).Npc(mapNpcNum).y - 1
                    
                    If NpcX >= GetPlayerX(attacker) - 1 And NpcX <= GetPlayerX(attacker) + 1 Then
                        If NpcY = GetPlayerY(attacker) Then
                            If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                                CanPlayerAttackNpc = True
                            ElseIf Npc(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Then
                                ' init conversation if it's friendly
                                InitChat attacker, MapNum, mapNpcNum
                            End If
                        End If
                    End If
                Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                    NpcX = MapNpc(MapNum).Npc(mapNpcNum).x + 1
                    NpcY = MapNpc(MapNum).Npc(mapNpcNum).y
                    
                    If NpcX = GetPlayerX(attacker) Then
                        If NpcY >= GetPlayerY(attacker) - 1 And NpcY <= GetPlayerY(attacker) + 1 Then
                            If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                                CanPlayerAttackNpc = True
                            ElseIf Npc(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Then
                                ' init conversation if it's friendly
                                InitChat attacker, MapNum, mapNpcNum
                            End If
                        End If
                    End If
                Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                    NpcX = MapNpc(MapNum).Npc(mapNpcNum).x - 1
                    NpcY = MapNpc(MapNum).Npc(mapNpcNum).y
                    
                    If NpcX = GetPlayerX(attacker) Then
                        If NpcY >= GetPlayerY(attacker) - 1 And NpcY <= GetPlayerY(attacker) + 1 Then
                            If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                                CanPlayerAttackNpc = True
                            ElseIf Npc(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Then
                                ' init conversation if it's friendly
                                InitChat attacker, MapNum, mapNpcNum
                            End If
                        End If
                    End If
            End Select
        End If
    End If

End Function

Public Sub PlayerAttackNpc(ByVal attacker As Long, ByVal mapNpcNum As Long, ByVal damage As Long, Optional ByVal spellNum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim N As Long
    Dim I As Long
    Dim STR As Long
    Dim DEF As Long
    Dim MapNum As Long
    Dim npcNum As Long
    Dim Mission_ID As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or damage < 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(attacker)
    npcNum = MapNpc(MapNum).Npc(mapNpcNum).Num
    Name = Trim$(Npc(npcNum).Name)
    
    ' Check for weapon
    N = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        N = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount

    If damage >= MapNpc(MapNum).Npc(mapNpcNum).Vital(Vitals.HP) Then
    
        SendActionMsg GetPlayerMap(attacker), "-" & MapNpc(MapNum).Npc(mapNpcNum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(MapNum).Npc(mapNpcNum).x * 32), (MapNpc(MapNum).Npc(mapNpcNum).y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y
        
        ' send the sound
        If spellNum > 0 Then SendMapSound attacker, MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y, SoundEntity.seSpell, spellNum
        
        ' send animation
        If N > 0 Then
            If Not overTime Then
                If spellNum = 0 Then Call SendAnimation(MapNum, Item(GetPlayerEquipment(attacker, Weapon)).Animation, MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y)
            End If
        End If

        ' Calculate exp to give attacker
        exp = Npc(npcNum).exp

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 1
        End If

        ' in party?
        If TempPlayer(attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(attacker).inParty, exp, attacker, Npc(npcNum).Level
        Else
            ' no party - keep exp for self
            GivePlayerEXP attacker, exp, Npc(npcNum).Level
        End If
        
        'Drop the goods if they get it
        For N = 1 To MAX_NPC_DROPS
            If Npc(npcNum).DropItem(N) = 0 Then Exit For
            If Rnd <= Npc(npcNum).DropChance(N) Then
                Call SpawnItem(Npc(npcNum).DropItem(N), Npc(npcNum).DropItemValue(N), MapNum, MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y, GetPlayerName(attacker))
            End If
        Next
        
        ' destroy map npcs
        If Map(MapNum).MapData.Moral = MAP_MORAL_BOSS Then
            If mapNpcNum = Map(MapNum).MapData.BossNpc Then
                ' kill all the other npcs
                For I = 1 To MAX_MAP_NPCS
                    If Map(MapNum).MapData.Npc(I) > 0 Then
                        ' only kill dangerous npcs
                        If Npc(Map(MapNum).MapData.Npc(I)).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(Map(MapNum).MapData.Npc(I)).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                            ' kill!
                            MapNpc(MapNum).Npc(I).Num = 0
                            MapNpc(MapNum).Npc(I).SpawnWait = GetTickCount
                            MapNpc(MapNum).Npc(I).Vital(Vitals.HP) = 0
                            ' send kill command
                            SendNpcDeath MapNum, I
                        End If
                    End If
                Next
            End If
        End If

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum).Npc(mapNpcNum).Num = 0
        MapNpc(MapNum).Npc(mapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum).Npc(mapNpcNum).Vital(Vitals.HP) = 0
        
        ' clear DoTs and HoTs
        For I = 1 To MAX_DOTS
            With MapNpc(MapNum).Npc(mapNpcNum).DoT(I)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(MapNum).Npc(mapNpcNum).HoT(I)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' Check Mission Kill NPC
        Call Check_Mission(attacker, npcNum)
        
        ' send death to the map
        SendNpcDeath MapNum, mapNpcNum
        
        'Loop through entire map and purge NPC from targets
        For I = 1 To Player_HighIndex
            If IsPlaying(I) And IsConnected(I) Then
                If Player(I).Map = MapNum Then
                    If TempPlayer(I).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(I).target = mapNpcNum Then
                            TempPlayer(I).target = 0
                            TempPlayer(I).targetType = TARGET_TYPE_NONE
                            SendTarget I
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum).Npc(mapNpcNum).Vital(Vitals.HP) = MapNpc(MapNum).Npc(mapNpcNum).Vital(Vitals.HP) - damage

        ' Check for a weapon and say damage
        SendActionMsg MapNum, "-" & damage, BrightRed, 1, (MapNpc(MapNum).Npc(mapNpcNum).x * 32), (MapNpc(MapNum).Npc(mapNpcNum).y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y
        
        ' send the sound
        If spellNum > 0 Then SendMapSound attacker, MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y, SoundEntity.seSpell, spellNum
        
        ' send animation
        If N > 0 Then
            If Not overTime Then
                If spellNum = 0 Then Call SendAnimation(MapNum, Item(GetPlayerEquipment(attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(MapNum).Npc(mapNpcNum).targetType = 1 ' player
        MapNpc(MapNum).Npc(mapNpcNum).target = attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum).Npc(mapNpcNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).Npc(I).Num = MapNpc(MapNum).Npc(mapNpcNum).Num Then
                    MapNpc(MapNum).Npc(I).target = attacker
                    MapNpc(MapNum).Npc(I).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(MapNum).Npc(mapNpcNum).stopRegen = True
        MapNpc(MapNum).Npc(mapNpcNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If spellNum > 0 Then
            If Spell(spellNum).StunDuration > 0 Then StunNPC mapNpcNum, MapNum, spellNum
            ' DoT
            If Spell(spellNum).Duration > 0 Then
                AddDoT_Npc MapNum, mapNpcNum, spellNum, attacker
            End If
        End If
        
        SendMapNpcVitals MapNum, mapNpcNum
        
        ' set the player's target if they don't have one
        If TempPlayer(attacker).target = 0 Then
            TempPlayer(attacker).targetType = TARGET_TYPE_NPC
            TempPlayer(attacker).target = mapNpcNum
            SendTarget attacker
        End If
    End If

    If spellNum = 0 Then
        ' Reset attack timer
        TempPlayer(attacker).AttackTimer = GetTickCount
    End If
End Sub
' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long)
Dim blockAmount As Long, npcNum As Long, MapNum As Long, damage As Long, Defence As Long

    damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(attacker, victim) Then
    
        MapNum = GetPlayerMap(attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If
        If CanPlayerParry(victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        damage = GetPlayerDamage(attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(victim)
        damage = damage - blockAmount
        
        ' take away armour
        Defence = GetPlayerDefence(victim)
        If Defence > 0 Then
            damage = damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If damage <= 0 Then damage = 1
        damage = RAND(damage - ((damage / 100) * 10), damage + ((damage / 100) * 10))
        
        ' * 1.5 if can crit
        If CanPlayerCrit(attacker) Then
            damage = damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(attacker) * 32), (GetPlayerY(attacker) * 32)
        End If

        If damage > 0 Then
            Call PlayerAttackPlayer(attacker, victim, damage)
        Else
            Call PlayerMsg(attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, Optional ByVal isSpell As Boolean = False) As Boolean
Dim partynum As Long, I As Long

    If Not isSpell Then
        ' Check attack timer
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            If GetTickCount < TempPlayer(attacker).AttackTimer + Item(GetPlayerEquipment(attacker, Weapon)).Speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function
    
    ' make sure it's not you
    If victim = attacker Then
        PlayerMsg attacker, "Cannot attack yourself.", BrightRed
        Exit Function
    End If
    
    ' check co-ordinates if not spell
    If Not isSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
            Case DIR_UP
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_DOWN
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_LEFT
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(attacker))) Then Exit Function
            Case DIR_RIGHT
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(attacker))) Then Exit Function
            Case DIR_UP_LEFT
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(attacker))) Then Exit Function
            Case DIR_UP_RIGHT
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(attacker))) Then Exit Function
            Case DIR_DOWN_LEFT
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(attacker))) Then Exit Function
            Case DIR_DOWN_RIGHT
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).MapData.Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(victim) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(attacker) < 5 Then
        Call PlayerMsg(attacker, "You are below level 5, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 5 Then
        Call PlayerMsg(attacker, GetPlayerName(victim) & " is below level 5, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If
    
    ' make sure not in your party
    partynum = TempPlayer(attacker).inParty
    If partynum > 0 Then
        For I = 1 To MAX_PARTY_MEMBERS
            If Party(partynum).Member(I) > 0 Then
                If victim = Party(partynum).Member(I) Then
                    PlayerMsg attacker, "Cannot attack party members.", BrightRed
                    Exit Function
                End If
            End If
        Next
    End If

    CanPlayerAttackPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, ByVal damage As Long, Optional ByVal spellNum As Long = 0)
    Dim exp As Long
    Dim N As Long
    Dim I As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(victim) = False Or damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    N = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        N = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount

    If damage >= GetPlayerVital(victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        If spellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, spellNum
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(attacker), BrightRed)
        ' Calculate exp to give attacker
        exp = (GetPlayerExp(victim) \ 10)

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 0
        End If

        If exp = 0 Then
            Call PlayerMsg(victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(victim, GetPlayerExp(victim) - exp)
            SendEXP victim
            Call PlayerMsg(victim, "You lost " & exp & " exp.", BrightRed)
            
            ' check if we're in a party
            If TempPlayer(attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(attacker).inParty, exp, attacker, GetPlayerLevel(victim)
            Else
                ' not in party, get exp for self
                GivePlayerEXP attacker, exp, GetPlayerLevel(victim)
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For I = 1 To Player_HighIndex
            If IsPlaying(I) And IsConnected(I) Then
                If Player(I).Map = GetPlayerMap(attacker) Then
                    If TempPlayer(I).target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(I).target = victim Then
                            TempPlayer(I).target = 0
                            TempPlayer(I).targetType = TARGET_TYPE_NONE
                            SendTarget I
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(victim) = NO Then
            If GetPlayerPK(attacker) = NO Then
                Call SetPlayerPK(attacker, YES)
                Call SendPlayerData(attacker)
                Call GlobalMsg(GetPlayerName(attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If

        Call OnDeath(victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - damage)
        Call SendVital(victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        If spellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, spellNum
        
        SendActionMsg GetPlayerMap(victim), "-" & damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If spellNum > 0 Then
            If Spell(spellNum).StunDuration > 0 Then StunPlayer victim, spellNum
            ' DoT
            If Spell(spellNum).Duration > 0 Then
                AddDoT_Player victim, spellNum, attacker
            End If
        End If
        
        ' change target if need be
        If TempPlayer(attacker).target = 0 Then
            TempPlayer(attacker).targetType = TARGET_TYPE_PLAYER
            TempPlayer(attacker).target = victim
            SendTarget attacker
        End If
    End If

    ' Reset attack timer
    TempPlayer(attacker).AttackTimer = GetTickCount
End Sub

' ############
' ## Spells ##
' ############
Public Sub BufferSpell(ByVal Index As Long, ByVal spellSlot As Long)
    Dim spellNum As Long, mpCost As Long, LevelReq As Long, MapNum As Long, spellCastType As Long, ClassReq As Long
    Dim AccessReq As Long, Range As Long, HasBuffered As Boolean, targetType As Byte, target As Long
    
    ' Prevent subscript out of range
    If spellSlot <= 0 Or spellSlot > MAX_PLAYER_SPELLS Then Exit Sub
    
    spellNum = Player(Index).Spell(spellSlot).Spell
    MapNum = GetPlayerMap(Index)
    
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
    If Spell(spellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellNum).IsAoE Then
            spellCastType = 2 ' targetted
        Else
            spellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellNum).IsAoE Then
            spellCastType = 0 ' self-cast
        Else
            spellCastType = 1 ' self-cast AoE
        End If
    End If
    
    targetType = TempPlayer(Index).targetType
    target = TempPlayer(Index).target
    Range = Spell(spellNum).Range
    HasBuffered = False
    
    Select Case spellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not target > 0 Then
                PlayerMsg Index, "You do not have a target.", BrightRed
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(target), GetPlayerY(target)) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                Else
                    ' go through spell types
                    If Spell(spellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(Index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if beneficial magic then self-cast it instead
                If Spell(spellNum).Type = SPELL_TYPE_HEALHP Or Spell(spellNum).Type = SPELL_TYPE_HEALMP Then
                    target = Index
                    targetType = TARGET_TYPE_PLAYER
                    HasBuffered = True
                Else
                    ' if have target, check in range
                    If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), MapNpc(MapNum).Npc(target).x, MapNpc(MapNum).Npc(target).y) Then
                        PlayerMsg Index, "Target not in range.", BrightRed
                        HasBuffered = False
                    Else
                        ' go through spell types
                        If Spell(spellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                            HasBuffered = True
                        Else
                            If CanPlayerAttackNpc(Index, target, True) Then
                                HasBuffered = True
                            End If
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation MapNum, Spell(spellNum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, Index, 1
        TempPlayer(Index).spellBuffer.Spell = spellSlot
        TempPlayer(Index).spellBuffer.Timer = GetTickCount
        TempPlayer(Index).spellBuffer.target = target
        TempPlayer(Index).spellBuffer.tType = targetType
        Exit Sub
    Else
        SendClearSpellBuffer Index
    End If
End Sub

Public Sub CastSpell(ByVal Index As Long, ByVal spellSlot As Long, ByVal target As Long, ByVal targetType As Byte)
    Dim spellNum As Long, mpCost As Long, LevelReq As Long
    Dim MapNum As Long, Vital As Long, DidCast As Boolean, ClassReq As Long
    Dim AccessReq As Long, I As Long, AoE As Long, Range As Long
    Dim vitalType As Byte, increment As Boolean, x As Long, y As Long
    Dim Buffer As clsBuffer, spellCastType As Long
    
    DidCast = False

    ' Prevent subscript out of range
    If spellSlot <= 0 Or spellSlot > MAX_PLAYER_SPELLS Then Exit Sub

    spellNum = Player(Index).Spell(spellSlot).Spell
    MapNum = GetPlayerMap(Index)

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
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellNum).IsAoE Then
            spellCastType = 2 ' targetted
        Else
            spellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellNum).IsAoE Then
            spellCastType = 0 ' self-cast
        Else
            spellCastType = 1 ' self-cast AoE
        End If
    End If
    
    ' get damage
    Vital = GetPlayerSpellDamage(Index, spellNum)
    
    ' store data
    AoE = Spell(spellNum).AoE
    Range = Spell(spellNum).Range
    
    Select Case spellCastType
        Case 0 ' self-cast target
            Select Case Spell(spellNum).Type
                Case SPELL_TYPE_HEALHP
                    SpellPlayer_Effect Vitals.HP, True, Index, Vital, spellNum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, Index, Vital, spellNum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation MapNum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    PlayerWarp Index, Spell(spellNum).Map, Spell(spellNum).x, Spell(spellNum).y
                    SendAnimation GetPlayerMap(Index), Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If spellCastType = 1 Then
                x = GetPlayerX(Index)
                y = GetPlayerY(Index)
            ElseIf spellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
                
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(target)
                    y = GetPlayerY(target)
                Else
                    x = MapNpc(MapNum).Npc(target).x
                    y = MapNpc(MapNum).Npc(target).y
                End If
                
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), x, y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    SendClearSpellBuffer Index
                End If
            End If
            Select Case Spell(spellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    For I = 1 To Player_HighIndex
                        If IsPlaying(I) Then
                            If I <> Index Then
                                If GetPlayerMap(I) = GetPlayerMap(Index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(I), GetPlayerY(I)) Then
                                        If CanPlayerAttackPlayer(Index, I, True) Then
                                            SendAnimation MapNum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, I
                                            PlayerAttackPlayer Index, I, Vital, spellNum
                                            DidCast = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For I = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(I).Num > 0 Then
                            If MapNpc(MapNum).Npc(I).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(MapNum).Npc(I).x, MapNpc(MapNum).Npc(I).y) Then
                                    If CanPlayerAttackNpc(Index, I, True) Then
                                        SendAnimation MapNum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, I
                                        PlayerAttackNpc Index, I, Vital, spellNum
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
                    
                    For I = 1 To Player_HighIndex
                        If IsPlaying(I) Then
                            If GetPlayerMap(I) = GetPlayerMap(Index) Then
                                If isInRange(AoE, x, y, GetPlayerX(I), GetPlayerY(I)) Then
                                    SpellPlayer_Effect vitalType, increment, I, Vital, spellNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
                    
                    If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        For I = 1 To MAX_MAP_NPCS
                            If MapNpc(MapNum).Npc(I).Num > 0 Then
                                If MapNpc(MapNum).Npc(I).Vital(HP) > 0 Then
                                    If isInRange(AoE, x, y, MapNpc(MapNum).Npc(I).x, MapNpc(MapNum).Npc(I).y) Then
                                        SpellNpc_Effect vitalType, increment, I, Vital, spellNum, MapNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        Next
                    End If
            End Select
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If target = 0 Then Exit Sub
            
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(target)
                y = GetPlayerY(target)
            Else
                x = MapNpc(MapNum).Npc(target).x
                y = MapNpc(MapNum).Npc(target).y
            End If
                
            If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), x, y) Then
                PlayerMsg Index, "Target not in range.", BrightRed
                SendClearSpellBuffer Index
                Exit Sub
            End If
            
            Select Case Spell(spellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(Index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                                PlayerAttackPlayer Index, target, Vital, spellNum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(Index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, target
                                PlayerAttackNpc Index, target, Vital, spellNum
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
                            If CanPlayerAttackPlayer(Index, target, True) Then
                                SpellPlayer_Effect vitalType, increment, target, Vital, spellNum
                                DidCast = True
                            End If
                        Else
                            SpellPlayer_Effect vitalType, increment, target, Vital, spellNum
                            DidCast = True
                        End If
                    Else
                        If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(Index, target, True) Then
                                SpellNpc_Effect vitalType, increment, target, Vital, spellNum, MapNum
                                DidCast = True
                            End If
                        Else
                            SpellNpc_Effect vitalType, increment, target, Vital, spellNum, MapNum
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
    Dim spellNum As Long, I As Long
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
                For I = 1 To MAX_HOTBAR
                    If Player(Index).Hotbar(I).Slot > 0 Then
                        If Player(Index).Hotbar(I).sType = 2 Then ' spell
                            If Spell(Player(Index).Hotbar(I).Slot).UniqueIndex = Spell(Spell(spellNum).NextRank).UniqueIndex Then
                                Player(Index).Hotbar(I).Slot = Spell(spellNum).NextRank
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

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal damage As Long, ByVal spellNum As Long)
    Dim sSymbol As String * 1
    Dim colour As Long

    If damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then colour = BrightGreen
            If Vital = Vitals.MP Then colour = BrightBlue
        Else
            sSymbol = "-"
            colour = Blue
        End If
    
        SendAnimation GetPlayerMap(Index), Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        SendActionMsg GetPlayerMap(Index), sSymbol & damage, colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        
        ' send the sound
        SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, spellNum
        
        If increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) + damage
            If Spell(spellNum).Duration > 0 Then
                AddHoT_Player Index, spellNum
            End If
        ElseIf Not increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) - damage
        End If
        
        ' send update
        SendVital Index, Vital
    End If
End Sub

Public Sub AddDoT_Player(ByVal Index As Long, ByVal spellNum As Long, ByVal Caster As Long)
    Dim I As Long

    For I = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(I)
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
    Dim I As Long

    For I = 1 To MAX_DOTS
        With TempPlayer(Index).HoT(I)
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
                SendActionMsg Player(Index).Map, "+" & GetPlayerSpellDamage(.Caster, .Spell), BrightGreen, ACTIONMSG_SCROLL, Player(Index).x * 32, Player(Index).y * 32
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
