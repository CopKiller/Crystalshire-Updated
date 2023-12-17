Attribute VB_Name = "Player_Checks"
Public Function IsPlayerItemRequerimentsOK(ByVal PlayerIndex As Long, ByVal ItemNum As Long) As Boolean
    Dim Text As String
    IsPlayerItemRequerimentsOK = True
    
    ' stat requirement
    For i = 1 To Stats.Stat_Count - 1
        If Item(ItemNum).Stat_Req(i) > GetPlayerStat(PlayerIndex, i) Then
            Select Case i
                Case Stats.Intelligence
                    Call PlayerMsg(PlayerIndex, "Você não tem a Intelligence mínima necessária.", BrightRed)
                Case Stats.Intelligence
                    Call PlayerMsg(PlayerIndex, "Você não tem a Intelligence mínima necessária.", BrightRed)
                Case Stats.Agility
                    Call PlayerMsg(PlayerIndex, "Você não tem a Agility mínima necessária.", BrightRed)
                Case Stats.Endurance
                    Call PlayerMsg(PlayerIndex, "Você não tem o Endurance mínimo necessário.", BrightRed)
                Case Stats.Willpower
                    Call PlayerMsg(PlayerIndex, "Você não tem a Willpower mínima necessária.", BrightRed)
            End Select
            IsPlayerItemRequerimentsOK = False
        End If
    Next
    
    ' level requirement
    If GetPlayerLevel(PlayerIndex) < Item(ItemNum).LevelReq Then
        Call PlayerMsg(PlayerIndex, "Você precisa estar no level " & Item(ItemNum).LevelReq & " para usar este item.", BrightRed)
        IsPlayerItemRequerimentsOK = False
    End If
    
    ' access requirement
    If Not GetPlayerAccess(PlayerIndex) >= Item(ItemNum).AccessReq Then
        Call PlayerMsg(PlayerIndex, "Você não tem acesso para usar esse item.", BrightRed)
        IsPlayerItemRequerimentsOK = False
    End If
    
    ' prociency requirement
    If Not hasProficiency(PlayerIndex, Item(ItemNum).proficiency) Then
        Call PlayerMsg(PlayerIndex, "Você não tem a proficiência que este item requer.", BrightRed)
        IsPlayerItemRequerimentsOK = False
    End If
    
    ' class requirement
    If Item(ItemNum).ClassReq > 0 Then
        If Not GetPlayerClass(index) = Item(ItemNum).ClassReq Then
            Call PlayerMsg(PlayerIndex, "You do not meet the class requirement to equip this item.", BrightRed)
            IsPlayerItemRequerimentsOK = False
        End If
    End If
End Function

Public Function CanMove(index As Long, Dir As Long) As Byte
    Dim warped As Boolean, newMapX As Long, newMapY As Long

    CanMove = 1
    Select Case Dir
        Case DIR_UP
            ' Check to see if they are trying to go out of bounds
            If GetPlayerY(index) > 0 Then
                If CheckDirection(index, DIR_UP) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If Map(GetPlayerMap(index)).MapData.Up > 0 Then
                    newMapY = Map(Map(GetPlayerMap(index)).MapData.Up).MapData.MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.Up, GetPlayerX(index), newMapY)
                    warped = True
                    CanMove = 2
                End If
                CanMove = 0
                Exit Function
            End If
'#######################################################################################################################
'#######################################################################################################################
        Case DIR_DOWN
            ' Check to see if they are trying to go out of bounds
            If GetPlayerY(index) < Map(GetPlayerMap(index)).MapData.MaxY Then
                If CheckDirection(index, DIR_DOWN) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If Map(GetPlayerMap(index)).MapData.Down > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.Down, GetPlayerX(index), 0)
                    warped = True
                    CanMove = 2
                End If
                CanMove = 0
                Exit Function
            End If
'#######################################################################################################################
'#######################################################################################################################
        Case DIR_LEFT
            ' Check to see if they are trying to go out of bounds
            If GetPlayerX(index) > 0 Then
                If CheckDirection(index, DIR_LEFT) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If Map(GetPlayerMap(index)).MapData.left > 0 Then
                    newMapX = Map(Map(GetPlayerMap(index)).MapData.left).MapData.MaxX
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.left, newMapX, GetPlayerY(index))
                    warped = True
                    CanMove = 2
                End If
                CanMove = 0
                Exit Function
            End If
'#######################################################################################################################
'#######################################################################################################################
        Case DIR_RIGHT
            ' Check to see if they are trying to go out of bounds
            If GetPlayerX(index) < Map(GetPlayerMap(index)).MapData.MaxX Then
                If CheckDirection(index, DIR_RIGHT) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If Map(GetPlayerMap(index)).MapData.Right > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.Right, 0, GetPlayerY(index))
                    warped = True
                    CanMove = 2
                End If
                CanMove = 0
                Exit Function
            End If
'#######################################################################################################################
'#######################################################################################################################
        Case DIR_UP_LEFT
            ' Check to see if they are trying to go out of bounds
            If GetPlayerY(index) > 0 And GetPlayerX(index) > 0 Then
                If CheckDirection(index, DIR_UP_LEFT) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If GetPlayerY(index) = 0 Then
                    If Map(GetPlayerMap(index)).MapData.Up > 0 Then
                        newMapY = Map(Map(GetPlayerMap(index)).MapData.Up).MapData.MaxY
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.Up, GetPlayerX(index), newMapY)
                        warped = True
                        CanMove = 2
                    End If
                Else
                    If Map(GetPlayerMap(index)).MapData.left > 0 Then
                        newMapX = Map(Map(GetPlayerMap(index)).MapData.left).MapData.MaxX
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.left, newMapX, GetPlayerY(index))
                        warped = True
                        CanMove = 2
                    End If
                End If
                CanMove = 0
                Exit Function
            End If
'#######################################################################################################################
'#######################################################################################################################
        Case DIR_UP_RIGHT
            ' Check to see if they are trying to go out of bounds
            If GetPlayerY(index) > 0 And GetPlayerX(index) < Map(GetPlayerMap(index)).MapData.MaxX Then
                If CheckDirection(index, DIR_UP_RIGHT) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If GetPlayerY(index) = 0 Then
                    If Map(GetPlayerMap(index)).MapData.Up > 0 Then
                        newMapY = Map(Map(GetPlayerMap(index)).MapData.Up).MapData.MaxY
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.Up, GetPlayerX(index), newMapY)
                        warped = True
                        CanMove = 2
                    End If
                Else
                    If Map(GetPlayerMap(index)).MapData.Right > 0 Then
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.Right, 0, GetPlayerY(index))
                        warped = True
                        CanMove = 2
                    End If
                End If
                CanMove = 0
                Exit Function
            End If
'#######################################################################################################################
'#######################################################################################################################
        Case DIR_DOWN_LEFT
            ' Check to see if they are trying to go out of bounds
            If GetPlayerY(index) < Map(GetPlayerMap(index)).MapData.MaxY And GetPlayerX(index) > 0 Then
                If CheckDirection(index, DIR_DOWN_LEFT) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If GetPlayerY(index) = Map(GetPlayerMap(index)).MapData.MaxY Then
                    If Map(GetPlayerMap(index)).MapData.Down > 0 Then
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.Down, GetPlayerX(index), 0)
                        warped = True
                        CanMove = 2
                    End If
                Else
                    If Map(GetPlayerMap(index)).MapData.left > 0 Then
                        newMapX = Map(Map(GetPlayerMap(index)).MapData.left).MapData.MaxX
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.left, newMapX, GetPlayerY(index))
                        warped = True
                        CanMove = 2
                    End If
                End If
                CanMove = 0
                Exit Function
            End If
'#######################################################################################################################
'#######################################################################################################################
        Case DIR_DOWN_RIGHT
            ' Check to see if they are trying to go out of bounds
            If GetPlayerY(index) < Map(GetPlayerMap(index)).MapData.MaxY And GetPlayerX(index) < Map(GetPlayerMap(index)).MapData.MaxX Then
                If CheckDirection(index, DIR_DOWN_RIGHT) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If GetPlayerY(index) = Map(GetPlayerMap(index)).MapData.MaxY Then
                    If Map(GetPlayerMap(index)).MapData.Down > 0 Then
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.Down, GetPlayerX(index), 0)
                        warped = True
                        CanMove = 2
                    End If
                Else
                    If Map(GetPlayerMap(index)).MapData.Right > 0 Then
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.Right, 0, GetPlayerY(index))
                        warped = True
                        CanMove = 2
                    End If
                End If
                CanMove = 0
                Exit Function
            End If
    End Select
    ' check if we've warped
    If warped Then
        ' clear their target
        TempPlayer(index).target = 0
        TempPlayer(index).targetType = TARGET_TYPE_NONE
        SendTarget index
    End If
End Function

Public Function CheckDirection(index As Long, direction As Long) As Boolean
    Dim x As Long, y As Long, i As Long
    Dim EventCount As Long, mapnum As Long, page As Long

    CheckDirection = False
    
    Select Case direction
        Case DIR_UP
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1
        Case DIR_DOWN
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1
        Case DIR_LEFT
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)
        Case DIR_RIGHT
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)
        Case DIR_UP_LEFT
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index) - 1
        Case DIR_UP_RIGHT
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index) - 1
        Case DIR_DOWN_LEFT
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index) + 1
        Case DIR_DOWN_RIGHT
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index) + 1
    End Select
    
    ' Check to see if the map tile is blocked or not
    If Map(GetPlayerMap(index)).TileData.Tile(x, y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If Map(GetPlayerMap(index)).TileData.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        CheckDirection = True
        Exit Function
    End If
    
    ' Check to make sure that any events on that space aren't blocked
    mapnum = GetPlayerMap(index)
    ' Check to see if a player is already on that tile
    If Map(GetPlayerMap(index)).MapData.Moral = 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(index) Then
                If GetPlayerX(i) = x Then
                    If GetPlayerY(i) = y Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If

    ' Check to see if a npc is already on that tile
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(index)).Npc(i).Num > 0 Then
            If MapNpc(GetPlayerMap(index)).Npc(i).x = x Then
                If MapNpc(GetPlayerMap(index)).Npc(i).y = y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Sub CheckEquippedItems(ByVal index As Long)
    Dim Slot As Long
    Dim ItemNum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(index, i)

        If ItemNum > 0 Then

            Select Case i
                Case Equipment.Weapon

                    If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment index, 0, i
                Case Equipment.Armor

                    If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment index, 0, i
                Case Equipment.Helmet

                    If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment index, 0, i
                Case Equipment.Shield

                    If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment index, 0, i
                Case Equipment.Pants

                    If Item(ItemNum).Type <> ITEM_TYPE_PANTS Then SetPlayerEquipment index, 0, i
                Case Equipment.Feet

                    If Item(ItemNum).Type <> ITEM_TYPE_FEET Then SetPlayerEquipment index, 0, i
            End Select

        Else
            SetPlayerEquipment index, 0, i
        End If

    Next

End Sub

Public Function HasItem(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Public Function HasSpell(ByVal index As Long, ByVal spellNum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If Player(index).Spell(i).Spell = spellNum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Public Function CanPlayerPickupItem(ByVal index As Long, ByVal mapItemNum As Long)
    Dim mapnum As Long, tmpIndex As Long, i As Long

    mapnum = GetPlayerMap(index)
    
    ' no lock or locked to player?
    If MapItem(mapnum, mapItemNum).playerName = vbNullString Or MapItem(mapnum, mapItemNum).playerName = Trim$(GetPlayerName(index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    ' if in party show their party member's drops
    If TempPlayer(index).inParty > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            tmpIndex = Party(TempPlayer(index).inParty).Member(i)
            If tmpIndex > 0 Then
                If Trim$(GetPlayerName(tmpIndex)) = MapItem(mapnum, mapItemNum).playerName Then
                    If MapItem(mapnum, mapItemNum).Bound = 0 Then
                        CanPlayerPickupItem = True
                        Exit Function
                    End If
                End If
            End If
        Next
    End If
    
    ' exit out
    CanPlayerPickupItem = False
End Function

Public Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
    On Error Resume Next
    Dim i As Long
    Dim N As Long

    If GetPlayerEquipment(index, Weapon) > 0 Then
        N = (Rnd) * 2

        If N = 1 Then
            i = (GetPlayerStat(index, Stats.Strength) \ 2) + (GetPlayerLevel(index) \ 2)
            N = Int(Rnd * 100) + 1

            If N <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Public Function CanPlayerBlockHit(ByVal index As Long) As Boolean
    Dim i As Long
    Dim N As Long
    Dim ShieldSlot As Long
    ShieldSlot = GetPlayerEquipment(index, Shield)

    If ShieldSlot > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            i = (GetPlayerStat(index, Stats.Endurance) \ 2) + (GetPlayerLevel(index) \ 2)
            N = Int(Rnd * 100) + 1

            If N <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function

Public Sub CheckHas_Mission(ByVal index As Long, ByVal NpcIndex As Long)
    Dim MissionID As Long, NpcNum As Long
    Dim AlreadyOnMission As Boolean
    Dim HasCompletedMission As Boolean
    Dim AlreadyCompletedMission As Boolean
    Dim IsRepeatable As Boolean
    Dim ActualMissionID As Long
    Dim MissionSlot As Long
    Dim i As Long, N As Long
    Dim mapnum As Long
    
    mapnum = GetPlayerMap(index)
    
    If NpcIndex = 0 Then
        Exit Sub
    End If
    
    NpcNum = MapNpc(mapnum).Npc(NpcIndex).Num
'    'Before we bother checking the main Mission variables, if we are on a Mission to talk to this NPC, let's do that first!
'    For i = 1 To MAX_PLAYER_MISSIONS
'        MissionID = Player(index).Mission(i).ID
'        If MissionID > 0 Then
'            If Mission(MissionID).Type = Mission_TYPE_TALK Then
'                If Mission(MissionID).TalkNPC = NpcNum Then
'                    Call MissionChat(index, MissionID)
'                    Call CompleteMission(index, i)
'                    Exit Sub
'                End If
'            End If
'        End If
'    Next i
    
    'Exit out if it's not a Mission giver
    If Npc(NpcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY Then
        Exit Sub
    End If
    
    If Npc(NpcNum).Mission <> 0 Then
        MissionID = Npc(NpcNum).Mission
        For i = 1 To MAX_PLAYER_MISSIONS
            'Toggle flags to false, this will be modified further down as we run if statements
            AlreadyOnMission = False
            HasCompletedMission = False
            AlreadyCompletedMission = False
            IsRepeatable = False
            'See if we're already on the Mission,
            For N = 1 To MAX_MISSIONS
                If Player(index).CompletedMission(N) = MissionID Then
                    ' AlreadyOnMission = True
                    AlreadyCompletedMission = True
                    If Mission(MissionID).Repeatable = 1 Then
                       IsRepeatable = True
                    Else
                       IsRepeatable = False
                    End If
                End If
            Next N
            
            MissionSlot = 0
            
            For N = 1 To MAX_PLAYER_MISSIONS
                If Player(index).Mission(N).ID = MissionID Then
                    AlreadyOnMission = True
                    MissionSlot = N
                End If
            Next N
    
            'We're going to use a bunch of If statements to run some checks and figure out what state we're at in the Mission
            If AlreadyOnMission = True Then
            'Flag that we're already on this NPC's Mission
            ' AlreadyOnMission = True
            
                'if it's a kill Mission
                If Mission(MissionID).Type = MissionType.Mission_TypeKill Then
                'We have completed the objective...
                    If Player(index).Mission(MissionSlot).Count >= Mission(MissionID).KillNPCAmount Then
                        HasCompletedMission = True
                    Else
                        HasCompletedMission = False
                    End If
                'if it's a collect Mission
                ElseIf Mission(MissionID).Type = MissionType.Mission_TypeCollect Then
                    'We have completed the objective
                    If HasItem(index, Mission(MissionID).CollectItem) >= Mission(MissionID).CollectItemAmount Then
                        HasCompletedMission = True
                    Else
                        HasCompletedMission = False
                    End If
                'if it's a talk Mission
                ElseIf Mission(MissionID).Type = MissionType.Mission_TypeTalk Then
                    'We flag this as incomplete because we're not going to have a Mission to talk to the same person...or we shouldn't anyways, that'd bad dialogue..
                    HasCompletedMission = False
                End If
            End If
    
    
            If IsRepeatable = True And AlreadyOnMission = False And AlreadyCompletedMission = True Then
                Call OfferMission(index, MissionID)
                Exit Sub
            End If
            If IsRepeatable = True And AlreadyOnMission = True And AlreadyCompletedMission = True And HasCompletedMission = False Then
                Call SendChatBubble(GetPlayerMap(index), NpcNum, TARGET_TYPE_NPC, Mission(MissionID).Incomplete, White)
                Exit Sub
            End If
            If IsRepeatable = True And AlreadyOnMission = True And AlreadyCompletedMission = True And HasCompletedMission = True Then
                'Call MissionChat(index, MissionID)
                Call CompleteMission(index, MissionSlot)
                Exit Sub
            End If
            If AlreadyOnMission = True And HasCompletedMission = True And AlreadyCompletedMission = False Then
                'Call MissionChat(index, MissionID)
                Call CompleteMission(index, MissionSlot)
                Exit Sub
            End If
            If AlreadyOnMission = True And HasCompletedMission = False And AlreadyCompletedMission = False Then
                Call PlayerMsg(index, Mission(MissionID).Incomplete, BrightRed)
                Call SendChatBubble(GetPlayerMap(index), NpcNum, TARGET_TYPE_NPC, Mission(MissionID).Incomplete, White)
                Exit Sub
            End If
            If AlreadyOnMission = False And HasCompletedMission = False And AlreadyCompletedMission = False And Player(index).Mission(i).ID = 0 Then
                Call OfferMission(index, MissionID)
                Exit Sub
            End If
        Next i
    End If
End Sub


Public Sub Check_Mission(ByVal Player_Index As Long, ByVal Target_Index As Long, Optional ByVal ItemAmount As Long = 0)
    Dim Missin_ID As Long, i As Long
    Dim mapnum As Long
    
    mapnum = GetPlayerMap(Player_Index)
    
    For i = 1 To MAX_PLAYER_MISSIONS
        Mission_ID = Player(Player_Index).Mission(i).ID
        If Mission_ID > 0 Then
            If Mission(Mission_ID).Type = MissionType.Mission_TypeKill And Mission(Mission_ID).KillNPC = Target_Index Then
                If TempPlayer(Player_Index).inParty > 0 Then
                    'Party_ShareNPCKill TempPlayer(Player_Index).inParty, Player_Index, GetPlayerMap(Player_Index), Player(Player_Index).Mission(i).ID
                Else
                    If Player(Player_Index).Mission(i).Count < Mission(Mission_ID).KillNPCAmount Then
                        Player(Player_Index).Mission(i).Count = Player(Player_Index).Mission(i).Count + 1
                        If Player(Player_Index).Mission(i).Count Mod 5 = 0 Then
                            Call PlayerMsg(Player_Index, "(" & Player(Player_Index).Mission(i).Count & "/" & Mission(Mission_ID).KillNPCAmount & ") " & Npc(Target_Index).Name, Yellow)
                        End If
                        Call SendPlayerMission(Player_Index, Mission_ID)
                    Else
                        If Player(Player_Index).Mission(i).Count Mod 5 = 0 Then
                            Call PlayerMsg(Player_Index, "(" & Player(Player_Index).Mission(i).Count & "/" & Mission(Mission_ID).KillNPCAmount & ") " & Npc(Target_Index).Name, Yellow)
                        End If
                    End If
                End If
                
            ElseIf Mission(Mission_ID).Type = MissionType.Mission_TypeCollect And Mission(Mission_ID).CollectItem = MapItem(mapnum, Target_Index).Num Then
                If Player(Player_Index).Mission(i).Count < Mission(Mission_ID).CollectItemAmount Then
                    Player(Player_Index).Mission(i).Count = Player(Player_Index).Mission(i).Count + ItemAmount
                    If Player(Player_Index).Mission(i).Count Mod 5 = 0 Then
                        Call PlayerMsg(Player_Index, "(" & Player(Player_Index).Mission(i).Count & "/" & Mission(Mission_ID).CollectItemAmount & ") " & Item(MapItem(mapnum, Target_Index).Num).Name, Yellow)
                    End If
                    Call SendPlayerMission(Player_Index, Mission_ID)
                Else
                    If Player(Player_Index).Mission(i).Count Mod 5 = 0 Then
                        Call PlayerMsg(Player_Index, "(" & Player(Player_Index).Mission(i).Count & "/" & Mission(Mission_ID).CollectItemAmount & ") " & Item(MapItem(mapnum, Target_Index).Num).Name, Yellow)
                    End If
                End If
            ElseIf Mission(Mission_ID).Type = MissionType.Mission_TypeTalk And Mission(Mission_ID).TalkNPC = Target_Index Then
                If Player(Player_Index).Mission(i).Count = 0 Then
                    Player(Player_Index).Mission(i).Count = Player(Player_Index).Mission(i).Count + 1
                    If Player(Player_Index).Mission(i).Count <> 0 Then
                        Call PlayerMsg(Player_Index, "Mission conv " & Npc(Target_Index).Name, Yellow)
                    End If
                    Call SendPlayerMission(Player_Index, Mission_ID)
                End If
            End If
        End If
    Next i
End Sub
