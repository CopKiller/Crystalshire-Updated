Attribute VB_Name = "Player_Checks"
Public Function CanMove(index As Long, dir As Long) As Byte
    Dim warped As Boolean, newMapX As Long, newMapY As Long

    CanMove = 1
    Select Case dir
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
                CanMove = False
                Exit Function
            End If
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
                CanMove = False
                Exit Function
            End If
        Case DIR_RIGHT
            ' Check to see if they are trying to go out of bounds
            If GetPlayerX(index) < Map(GetPlayerMap(index)).MapData.MaxX Then
                If CheckDirection(index, DIR_RIGHT) Then
                    CanMove = False
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If Map(GetPlayerMap(index)).MapData.Right > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.Right, 0, GetPlayerY(index))
                    warped = True
                    CanMove = 2
                End If
                CanMove = False
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
    EventCount = Map(mapnum).TileData.EventCount
    For i = 1 To EventCount
        With Map(mapnum).TileData.Events(i)
            If .x = x And .y = y Then
                ' Get the active event page
                page = ActiveEventPage(index, i)
                If page > 0 Then
                    If Map(mapnum).TileData.Events(i).EventPage(page).WalkThrough = 0 Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        End With
    Next

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
    Dim itemNum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        itemNum = GetPlayerEquipment(index, i)

        If itemNum > 0 Then

            Select Case i
                Case Equipment.Weapon

                    If Item(itemNum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment index, 0, i
                Case Equipment.Armor

                    If Item(itemNum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment index, 0, i
                Case Equipment.Helmet

                    If Item(itemNum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment index, 0, i
                Case Equipment.Shield

                    If Item(itemNum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment index, 0, i
            End Select

        Else
            SetPlayerEquipment index, 0, i
        End If

    Next

End Sub

Public Function HasItem(ByVal index As Long, ByVal itemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemNum Then
            If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then
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

Public Sub CheckPlayerEvent(index As Long, eventNum As Long)
    Dim Count As Long, mapnum As Long, i As Long
    ' find the page to process
    mapnum = GetPlayerMap(index)
    ' make sure it's in the same spot
    If Map(mapnum).TileData.Events(eventNum).x <> GetPlayerX(index) Then Exit Sub
    If Map(mapnum).TileData.Events(eventNum).y <> GetPlayerY(index) Then Exit Sub
    ' loop
    Count = Map(mapnum).TileData.Events(eventNum).PageCount
    ' get the active page
    i = ActiveEventPage(index, eventNum)
    ' exit out early
    If i = 0 Then Exit Sub
    ' make sure the page has actual commands
    If Map(mapnum).TileData.Events(eventNum).EventPage(i).CommandCount = 0 Then Exit Sub
    ' set event
    TempPlayer(index).inEvent = True
    TempPlayer(index).eventNum = eventNum
    TempPlayer(index).pageNum = i
    TempPlayer(index).commandNum = 1
    ' send it to the player
    SendEvent index
End Sub

Public Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
    On Error Resume Next
    Dim i As Long
    Dim n As Long

    If GetPlayerEquipment(index, Weapon) > 0 Then
        n = (Rnd) * 2

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.Strength) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Public Function CanPlayerBlockHit(ByVal index As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ShieldSlot As Long
    ShieldSlot = GetPlayerEquipment(index, Shield)

    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.Endurance) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function
