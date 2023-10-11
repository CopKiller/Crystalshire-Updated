Attribute VB_Name = "Player_Checks"
Public Function IsPlayerItemRequerimentsOK(ByVal PlayerIndex As Long, ByVal ItemNum As Long) As Boolean
    Dim Text As String
    IsPlayerItemRequerimentsOK = True
    
    ' stat requirement
    For I = 1 To Stats.Stat_Count - 1
        If Item(ItemNum).Stat_Req(I) > GetPlayerStat(PlayerIndex, I) Then
            Select Case I
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
        If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
            PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
            IsPlayerItemRequerimentsOK = False
        End If
    End If
End Function

Public Function CanMove(Index As Long, Dir As Long) As Byte
    Dim warped As Boolean, newMapX As Long, newMapY As Long

    CanMove = 1
    Select Case Dir
        Case DIR_UP
            ' Check to see if they are trying to go out of bounds
            If GetPlayerY(Index) > 0 Then
                If CheckDirection(Index, DIR_UP) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If Map(GetPlayerMap(Index)).MapData.Up > 0 Then
                    newMapY = Map(Map(GetPlayerMap(Index)).MapData.Up).MapData.MaxY
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.Up, GetPlayerX(Index), newMapY)
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
            If GetPlayerY(Index) < Map(GetPlayerMap(Index)).MapData.MaxY Then
                If CheckDirection(Index, DIR_DOWN) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If Map(GetPlayerMap(Index)).MapData.Down > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.Down, GetPlayerX(Index), 0)
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
            If GetPlayerX(Index) > 0 Then
                If CheckDirection(Index, DIR_LEFT) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If Map(GetPlayerMap(Index)).MapData.left > 0 Then
                    newMapX = Map(Map(GetPlayerMap(Index)).MapData.left).MapData.MaxX
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.left, newMapX, GetPlayerY(Index))
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
            If GetPlayerX(Index) < Map(GetPlayerMap(Index)).MapData.MaxX Then
                If CheckDirection(Index, DIR_RIGHT) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If Map(GetPlayerMap(Index)).MapData.Right > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.Right, 0, GetPlayerY(Index))
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
            If GetPlayerY(Index) > 0 And GetPlayerX(Index) > 0 Then
                If CheckDirection(Index, DIR_UP_LEFT) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If GetPlayerY(Index) = 0 Then
                    If Map(GetPlayerMap(Index)).MapData.Up > 0 Then
                        newMapY = Map(Map(GetPlayerMap(Index)).MapData.Up).MapData.MaxY
                        Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.Up, GetPlayerX(Index), newMapY)
                        warped = True
                        CanMove = 2
                    End If
                Else
                    If Map(GetPlayerMap(Index)).MapData.left > 0 Then
                        newMapX = Map(Map(GetPlayerMap(Index)).MapData.left).MapData.MaxX
                        Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.left, newMapX, GetPlayerY(Index))
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
            If GetPlayerY(Index) > 0 And GetPlayerX(Index) < Map(GetPlayerMap(Index)).MapData.MaxX Then
                If CheckDirection(Index, DIR_UP_RIGHT) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If GetPlayerY(Index) = 0 Then
                    If Map(GetPlayerMap(Index)).MapData.Up > 0 Then
                        newMapY = Map(Map(GetPlayerMap(Index)).MapData.Up).MapData.MaxY
                        Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.Up, GetPlayerX(Index), newMapY)
                        warped = True
                        CanMove = 2
                    End If
                Else
                    If Map(GetPlayerMap(Index)).MapData.Right > 0 Then
                        Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.Right, 0, GetPlayerY(Index))
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
            If GetPlayerY(Index) < Map(GetPlayerMap(Index)).MapData.MaxY And GetPlayerX(Index) > 0 Then
                If CheckDirection(Index, DIR_DOWN_LEFT) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MapData.MaxY Then
                    If Map(GetPlayerMap(Index)).MapData.Down > 0 Then
                        Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.Down, GetPlayerX(Index), 0)
                        warped = True
                        CanMove = 2
                    End If
                Else
                    If Map(GetPlayerMap(Index)).MapData.left > 0 Then
                        newMapX = Map(Map(GetPlayerMap(Index)).MapData.left).MapData.MaxX
                        Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.left, newMapX, GetPlayerY(Index))
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
            If GetPlayerY(Index) < Map(GetPlayerMap(Index)).MapData.MaxY And GetPlayerX(Index) < Map(GetPlayerMap(Index)).MapData.MaxX Then
                If CheckDirection(Index, DIR_DOWN_RIGHT) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MapData.MaxY Then
                    If Map(GetPlayerMap(Index)).MapData.Down > 0 Then
                        Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.Down, GetPlayerX(Index), 0)
                        warped = True
                        CanMove = 2
                    End If
                Else
                    If Map(GetPlayerMap(Index)).MapData.Right > 0 Then
                        Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.Right, 0, GetPlayerY(Index))
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
        TempPlayer(Index).target = 0
        TempPlayer(Index).targetType = TARGET_TYPE_NONE
        SendTarget Index
    End If
End Function

Public Function CheckDirection(Index As Long, direction As Long) As Boolean
    Dim x As Long, y As Long, I As Long
    Dim EventCount As Long, MapNum As Long, page As Long

    CheckDirection = False
    
    Select Case direction
        Case DIR_UP
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) - 1
        Case DIR_DOWN
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) + 1
        Case DIR_LEFT
            x = GetPlayerX(Index) - 1
            y = GetPlayerY(Index)
        Case DIR_RIGHT
            x = GetPlayerX(Index) + 1
            y = GetPlayerY(Index)
        Case DIR_UP_LEFT
            x = GetPlayerX(Index) - 1
            y = GetPlayerY(Index) - 1
        Case DIR_UP_RIGHT
            x = GetPlayerX(Index) + 1
            y = GetPlayerY(Index) - 1
        Case DIR_DOWN_LEFT
            x = GetPlayerX(Index) - 1
            y = GetPlayerY(Index) + 1
        Case DIR_DOWN_RIGHT
            x = GetPlayerX(Index) + 1
            y = GetPlayerY(Index) + 1
    End Select
    
    ' Check to see if the map tile is blocked or not
    If Map(GetPlayerMap(Index)).TileData.Tile(x, y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If Map(GetPlayerMap(Index)).TileData.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        CheckDirection = True
        Exit Function
    End If
    
    ' Check to make sure that any events on that space aren't blocked
    MapNum = GetPlayerMap(Index)
    EventCount = Map(MapNum).TileData.EventCount
    For I = 1 To EventCount
        With Map(MapNum).TileData.Events(I)
            If .x = x And .y = y Then
                ' Get the active event page
                page = ActiveEventPage(Index, I)
                If page > 0 Then
                    If Map(MapNum).TileData.Events(I).EventPage(page).WalkThrough = 0 Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        End With
    Next

    ' Check to see if a player is already on that tile
    If Map(GetPlayerMap(Index)).MapData.Moral = 0 Then
        For I = 1 To Player_HighIndex
            If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(Index) Then
                If GetPlayerX(I) = x Then
                    If GetPlayerY(I) = y Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        Next I
    End If

    ' Check to see if a npc is already on that tile
    For I = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(Index)).Npc(I).Num > 0 Then
            If MapNpc(GetPlayerMap(Index)).Npc(I).x = x Then
                If MapNpc(GetPlayerMap(Index)).Npc(I).y = y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Sub CheckEquippedItems(ByVal Index As Long)
    Dim Slot As Long
    Dim ItemNum As Long
    Dim I As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For I = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(Index, I)

        If ItemNum > 0 Then

            Select Case I
                Case Equipment.Weapon

                    If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment Index, 0, I
                Case Equipment.Armor

                    If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment Index, 0, I
                Case Equipment.Helmet

                    If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment Index, 0, I
                Case Equipment.Shield

                    If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment Index, 0, I
            End Select

        Else
            SetPlayerEquipment Index, 0, I
        End If

    Next

End Sub

Public Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim I As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For I = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, I) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(Index, I)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Public Function HasSpell(ByVal Index As Long, ByVal spellNum As Long) As Boolean
    Dim I As Long

    For I = 1 To MAX_PLAYER_SPELLS

        If Player(Index).Spell(I).Spell = spellNum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Public Function CanPlayerPickupItem(ByVal Index As Long, ByVal mapItemNum As Long)
    Dim MapNum As Long, tmpIndex As Long, I As Long

    MapNum = GetPlayerMap(Index)
    
    ' no lock or locked to player?
    If MapItem(MapNum, mapItemNum).playerName = vbNullString Or MapItem(MapNum, mapItemNum).playerName = Trim$(GetPlayerName(Index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    ' if in party show their party member's drops
    If TempPlayer(Index).inParty > 0 Then
        For I = 1 To MAX_PARTY_MEMBERS
            tmpIndex = Party(TempPlayer(Index).inParty).Member(I)
            If tmpIndex > 0 Then
                If Trim$(GetPlayerName(tmpIndex)) = MapItem(MapNum, mapItemNum).playerName Then
                    If MapItem(MapNum, mapItemNum).Bound = 0 Then
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

Public Sub CheckPlayerEvent(Index As Long, eventNum As Long)
    Dim Count As Long, MapNum As Long, I As Long
    ' find the page to process
    MapNum = GetPlayerMap(Index)
    ' make sure it's in the same spot
    If Map(MapNum).TileData.Events(eventNum).x <> GetPlayerX(Index) Then Exit Sub
    If Map(MapNum).TileData.Events(eventNum).y <> GetPlayerY(Index) Then Exit Sub
    ' loop
    Count = Map(MapNum).TileData.Events(eventNum).PageCount
    ' get the active page
    I = ActiveEventPage(Index, eventNum)
    ' exit out early
    If I = 0 Then Exit Sub
    ' make sure the page has actual commands
    If Map(MapNum).TileData.Events(eventNum).EventPage(I).CommandCount = 0 Then Exit Sub
    ' set event
    TempPlayer(Index).inEvent = True
    TempPlayer(Index).eventNum = eventNum
    TempPlayer(Index).pageNum = I
    TempPlayer(Index).commandNum = 1
    ' send it to the player
    SendEvent Index
End Sub

Public Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
    On Error Resume Next
    Dim I As Long
    Dim N As Long

    If GetPlayerEquipment(Index, Weapon) > 0 Then
        N = (Rnd) * 2

        If N = 1 Then
            I = (GetPlayerStat(Index, Stats.Strength) \ 2) + (GetPlayerLevel(Index) \ 2)
            N = Int(Rnd * 100) + 1

            If N <= I Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Public Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
    Dim I As Long
    Dim N As Long
    Dim ShieldSlot As Long
    ShieldSlot = GetPlayerEquipment(Index, Shield)

    If ShieldSlot > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            I = (GetPlayerStat(Index, Stats.Endurance) \ 2) + (GetPlayerLevel(Index) \ 2)
            N = Int(Rnd * 100) + 1

            If N <= I Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function

Public Sub Check_Mission(ByVal Player_Index As Long, ByVal Target_Index As Long)
    Dim Missin_ID As Long, I As Long
    Dim MapNum As Long
    
    MapNum = GetPlayerMap(Player_Index)
    
    For I = 1 To MAX_PLAYER_MISSIONS
        Mission_ID = Player(Player_Index).Mission(I).ID
        If Mission_ID > 0 Then
            If Mission(Mission_ID).Type = MissionType.TypeKill And Mission(Mission_ID).KillNPC = Target_Index Then
                If TempPlayer(Player_Index).inParty > 0 Then
                    'Party_ShareNPCKill TempPlayer(Player_Index).inParty, Player_Index, GetPlayerMap(Player_Index), Player(Player_Index).Quest(i).ID
                Else
                    If Player(Player_Index).Mission(I).Count < Mission(Mission_ID).KillNPCAmount Then
                        Player(Player_Index).Mission(I).Count = Player(Player_Index).Mission(I).Count + 1
                        If Player(Player_Index).Mission(I).Count Mod 5 = 0 Then
                            Call PlayerMsg(Player_Index, Player(Player_Index).Mission(I).Count & "/" & Mission(Mission_ID).KillNPCAmount & " " & Npc(Target_Index).Name, Yellow)
                        End If
                        Call SendPlayerMission(Player_Index, Mission_ID)
                    End If
                End If
                
            ElseIf Mission(Mission_ID).Type = MissionType.TypeCollect And Mission(Mission_ID).CollectItem = MapItem(MapNum, Target_Index).Num Then
                If Player(Player_Index).Quest(I).Counter < Mission(Mission_ID).CollectItemAmount Then
                    Player(Player_Index).Mission(I).Count = Player(Player_Index).Mission(I).Count + 1
                    If Player(Player_Index).Mission(I).Count Mod 5 = 0 Then
                        Call PlayerMsg(Player_Index, Player(Player_Index).Mission(I).Count & "/" & Mission(Mission_ID).CollectItemAmount & " " & Item(MapItem(MapNum, Target_Index).Num).Name, Yellow)
                    End If
                    Call SendPlayerMission(Player_Index, Mission_ID)
                End If
            ElseIf Mission(Mission_ID).Type = MissionType.TypeTalk And Mission(Mission_ID).TalkNPC = Target_Index Then
                If Player(Player_Index).Mission(I).Count = 0 Then
                    Player(Player_Index).Mission(I).Count = Player(Player_Index).Mission(I).Count + 1
                    If Player(Player_Index).Mission(I).Count <> 5 = 0 Then
                        Call PlayerMsg(Player_Index, "Mission conv " & Npc(Target_Index).Name, Yellow)
                    End If
                    Call SendPlayerMission(Player_Index, Mission_ID)
                End If
            End If
        End If
    Next I
End Sub
