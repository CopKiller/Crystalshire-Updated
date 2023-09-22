Attribute VB_Name = "Player_Handle"
Public Sub HandleRequestPlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerData(Index)
End Sub

Public Sub HandleRequestLevelUp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If GetPlayerAccess(Index) < 4 Then Exit Sub
    Call SetPlayerExp(Index, GetPlayerNextLevel(Index))
    Call CheckPlayerLevelUp(Index)
End Sub
' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim name As String
    Dim i As Long
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    name = buffer.ReadString 'Parse(1)
    buffer.Flush: Set buffer = Nothing
    i = FindPlayer(name)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Public Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim dir As Long
    Dim movement As Long
    Dim buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    dir = buffer.ReadLong 'CLng(Parse(1))
    movement = buffer.ReadLong 'CLng(Parse(2))
    tmpX = buffer.ReadLong
    tmpY = buffer.ReadLong
    buffer.Flush: Set buffer = Nothing

    ' Prevent hacking
    If dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If

    ' Prevent hacking
    If movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    ' Prevent player from moving if they have casted a spell
    'If TempPlayer(index).spellBuffer.Spell > 0 Then
    '    Call SendPlayerXY(index)
    '    Exit Sub
    'End If
    
    'Cant move if in the bank!
    If TempPlayer(Index).InBank Then
        'Call SendPlayerXY(Index)
        'Exit Sub
        TempPlayer(Index).InBank = False
    End If

    ' if stunned, stop them moving
    If TempPlayer(Index).StunDuration > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If
    
    ' Prever player from moving if in shop
    If TempPlayer(Index).InShop > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(Index) <> tmpX Then
        SendPlayerXY (Index)
        Exit Sub
    End If
    
    If GetPlayerY(Index) <> tmpY Then
        SendPlayerXY (Index)
        Exit Sub
    End If
    
    ' cant move if chatting
    If TempPlayer(Index).inChatWith > 0 Then
        ClosePlayerChat Index
    End If
    
    Call PlayerMove(Index, dir, movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Public Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim dir As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    dir = buffer.ReadLong 'CLng(Parse(1))
    buffer.Flush: Set buffer = Nothing

    ' Prevent hacking
    If dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, dir)
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerDir
    buffer.WriteLong Index
    buffer.WriteLong GetPlayerDir(Index)
    SendDataToMapBut Index, GetPlayerMap(Index), buffer.ToArray()
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Public Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, n As Long, damage As Long, TempIndex As Long, x As Long, y As Long, mapnum As Long, dirReq As Long
    
    ' can't attack whilst casting
    If TempPlayer(Index).spellBuffer.Spell > 0 Then Exit Sub
    
    ' can't attack whilst stunned
    If TempPlayer(Index).StunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
    SendAttack Index

    ' Try to attack a player
    For i = 1 To Player_HighIndex
        TempIndex = i

        ' Make sure we dont try to attack ourselves
        If TempIndex <> Index Then
            TryPlayerAttackPlayer Index, i
        End If
    Next

    ' Try to attack a npc
    For i = 1 To MAX_MAP_NPCS
        TryPlayerAttackNpc Index, i
    Next
    
    ' check if we've got a remote chat tile
    mapnum = GetPlayerMap(Index)
    x = GetPlayerX(Index)
    y = GetPlayerY(Index)
    If Map(mapnum).TileData.Tile(x, y).Type = TILE_TYPE_CHAT Then
        dirReq = Map(mapnum).TileData.Tile(x, y).Data2
        If Player(Index).dir = dirReq Then
            InitChat Index, mapnum, Map(mapnum).TileData.Tile(x, y).Data1, True
            Exit Sub
        End If
    End If

    ' Check tradeskills
    Select Case GetPlayerDir(Index)
        Case DIR_UP

            If GetPlayerY(Index) = 0 Then Exit Sub
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) - 1
        Case DIR_DOWN

            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MapData.MaxY Then Exit Sub
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) + 1
        Case DIR_LEFT

            If GetPlayerX(Index) = 0 Then Exit Sub
            x = GetPlayerX(Index) - 1
            y = GetPlayerY(Index)
        Case DIR_RIGHT

            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MapData.MaxX Then Exit Sub
            x = GetPlayerX(Index) + 1
            y = GetPlayerY(Index)
    End Select
    
    CheckResource Index, x, y
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Public Sub HandleCast(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Spell slot
    n = buffer.ReadLong 'CLng(Parse(1))
    buffer.Flush: Set buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(Index, n)
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Public Sub HandleUseItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim invNum As Long
    Dim buffer As clsBuffer
    
    ' get inventory slot number
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    invNum = buffer.ReadLong
    buffer.Flush: Set buffer = Nothing

    UseItem Index, invNum
End Sub

Public Sub HandleUnequip(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PlayerUnequipItem Index, buffer.ReadLong
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim tmpItem As Long
    Dim tmpAmount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' item
    tmpItem = buffer.ReadLong
    tmpAmount = buffer.ReadLong
        
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then Exit Sub
    
    SpawnItem tmpItem, tmpAmount, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), GetPlayerName(Index)
    buffer.Flush: Set buffer = Nothing
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Public Sub HandleTarget(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, target As Long, targetType As Long

    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    
    target = buffer.ReadLong
    targetType = buffer.ReadLong
    
    buffer.Flush: Set buffer = Nothing
    
    ' set player's target - no need to send, it's client side
    TempPlayer(Index).target = target
    TempPlayer(Index).targetType = targetType
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Public Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(Index)
End Sub

Public Sub HandleForgetSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim spellSlot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    spellSlot = buffer.ReadLong
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If TempPlayer(Index).SpellCD(spellSlot) > GetTickCount Then
        PlayerMsg Index, "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(Index).spellBuffer.Spell = spellSlot Then
        PlayerMsg Index, "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    Player(Index).Spell(spellSlot).Spell = 0
    Player(Index).Spell(spellSlot).Uses = 0
    SendPlayerSpells Index
    
    buffer.Flush: Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Public Sub HandleUseStatPoint(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim PointType As Byte
    Dim buffer As clsBuffer
    Dim sMes As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PointType = buffer.ReadByte 'CLng(Parse(1))
    buffer.Flush: Set buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_Count) Then
        Exit Sub
    End If

    ' Make sure they have points
    If GetPlayerPOINTS(Index) > 0 Then
        ' make sure they're not maxed
        If GetPlayerRawStat(Index, PointType) >= 255 Then
            PlayerMsg Index, "You cannot spend any more points on that stat.", BrightRed
            Exit Sub
        End If
        
        ' make sure they're not spending too much
        If GetPlayerRawStat(Index, PointType) - Class(GetPlayerClass(Index)).Stat(PointType) >= (GetPlayerLevel(Index) * 2) - 1 Then
            PlayerMsg Index, "You cannot spend any more points on that stat.", BrightRed
            Exit Sub
        End If
        
        ' Take away a stat point
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)

        ' Everything is ok
        Select Case PointType
            Case Stats.Strength
                Call SetPlayerStat(Index, Stats.Strength, GetPlayerRawStat(Index, Stats.Strength) + 1)
                sMes = "Strength"
            Case Stats.Endurance
                Call SetPlayerStat(Index, Stats.Endurance, GetPlayerRawStat(Index, Stats.Endurance) + 1)
                sMes = "Endurance"
            Case Stats.Intelligence
                Call SetPlayerStat(Index, Stats.Intelligence, GetPlayerRawStat(Index, Stats.Intelligence) + 1)
                sMes = "Intelligence"
            Case Stats.Agility
                Call SetPlayerStat(Index, Stats.Agility, GetPlayerRawStat(Index, Stats.Agility) + 1)
                sMes = "Agility"
            Case Stats.Willpower
                Call SetPlayerStat(Index, Stats.Willpower, GetPlayerRawStat(Index, Stats.Willpower) + 1)
                sMes = "Willpower"
        End Select
        
        SendActionMsg GetPlayerMap(Index), "+1 " & sMes, White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
    Else
        Exit Sub
    End If

    ' Send the update
    SendPlayerData Index
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Public Sub HandleSwapInvSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    
    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = buffer.ReadLong
    newSlot = buffer.ReadLong
    buffer.Flush: Set buffer = Nothing
    PlayerSwitchInvSlots Index, oldSlot, newSlot
End Sub

Public Sub HandleSwapSpellSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long, n As Long
    
    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub
    
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        PlayerMsg Index, "You cannot swap spells whilst casting.", BrightRed
        Exit Sub
    End If
    
    For n = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(Index).SpellCD(n) > GetTickCount Then
            PlayerMsg Index, "You cannot swap spells whilst they're cooling down.", BrightRed
            Exit Sub
        End If
    Next
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = buffer.ReadLong
    newSlot = buffer.ReadLong
    buffer.Flush: Set buffer = Nothing
    PlayerSwitchSpellSlots Index, oldSlot, newSlot
End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Public Sub HandleWhosOnline(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(Index)
End Sub

' :::::::::::::::::::::::
' ::    Party packet   ::
' :::::::::::::::::::::::

Public Sub HandlePartyRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, targetIndex As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    targetIndex = buffer.ReadLong
    buffer.Flush: Set buffer = Nothing
    
    ' make sure it's a valid target
    If targetIndex = Index Then
        PlayerMsg Index, "You can't invite yourself. That would be weird.", BrightRed
        Exit Sub
    End If
    
    ' make sure they're connected and on the same map
    If Not IsConnected(targetIndex) Or Not IsPlaying(targetIndex) Then Exit Sub
    If GetPlayerMap(targetIndex) <> GetPlayerMap(Index) Then Exit Sub
    
    ' init the request
    Party_Invite Index, targetIndex
End Sub

Public Sub HandleAcceptParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteAccept TempPlayer(Index).partyInvite, Index
End Sub

Public Sub HandleDeclineParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteDecline TempPlayer(Index).partyInvite, Index
End Sub

Public Sub HandlePartyLeave(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_PlayerLeave Index
End Sub

' :::::::::::::::::::::::
' ::   HOTBAR packet   ::
' :::::::::::::::::::::::

Public Sub HandleHotbarChange(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim sType As Long
    Dim Slot As Long
    Dim hotbarNum As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    sType = buffer.ReadLong
    Slot = buffer.ReadLong
    hotbarNum = buffer.ReadLong
    
    Select Case sType
        Case 0 ' clear
            Player(Index).Hotbar(hotbarNum).Slot = 0
            Player(Index).Hotbar(hotbarNum).sType = 0
        Case 1 ' inventory
            If Slot > 0 And Slot <= MAX_INV Then
                If Player(Index).Inv(Slot).Num > 0 Then
                    If Len(Trim$(Item(GetPlayerInvItemNum(Index, Slot)).name)) > 0 Then
                        Player(Index).Hotbar(hotbarNum).Slot = Player(Index).Inv(Slot).Num
                        Player(Index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
        Case 2 ' spell
            If Slot > 0 And Slot <= MAX_PLAYER_SPELLS Then
                If Player(Index).Spell(Slot).Spell > 0 Then
                    If Len(Trim$(Spell(Player(Index).Spell(Slot).Spell).name)) > 0 Then
                        Player(Index).Hotbar(hotbarNum).Slot = Player(Index).Spell(Slot).Spell
                        Player(Index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
    End Select
    
    SendHotbar Index
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub HandleHotbarUse(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Slot As Long
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Slot = buffer.ReadLong
    
    Select Case Player(Index).Hotbar(Slot).sType
        Case 1 ' inventory
            For i = 1 To MAX_INV
                If Player(Index).Inv(i).Num > 0 Then
                    If Player(Index).Inv(i).Num = Player(Index).Hotbar(Slot).Slot Then
                        UseItem Index, i
                        Exit Sub
                    End If
                End If
            Next
        Case 2 ' spell
            For i = 1 To MAX_PLAYER_SPELLS
                If Player(Index).Spell(i).Spell > 0 Then
                    If Player(Index).Spell(i).Spell = Player(Index).Hotbar(Slot).Slot Then
                        BufferSpell Index, i
                        Exit Sub
                    End If
                End If
            Next
    End Select
    
    buffer.Flush: Set buffer = Nothing
End Sub

' :::::::::::::::::::::::
' ::    TRADE packet   ::
' :::::::::::::::::::::::

Public Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long, buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' find the target
    tradeTarget = buffer.ReadLong
    
    buffer.Flush: Set buffer = Nothing
    
    If Not IsConnected(Index) Or Not IsPlaying(Index) Then
        TempPlayer(tradeTarget).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If
    
    If Not IsConnected(tradeTarget) Or Not IsPlaying(tradeTarget) Then
        TempPlayer(tradeTarget).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If
    
    ' make sure we don't error
    If tradeTarget <= 0 Or tradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' can't trade with yourself..
    If tradeTarget = Index Then
        PlayerMsg Index, "You can't trade with yourself.", BrightRed
        Exit Sub
    End If
    
    ' make sure they're on the same map
    If Not Player(tradeTarget).Map = Player(Index).Map Then Exit Sub
    
    ' make sure they're stood next to each other
    tX = Player(tradeTarget).x
    tY = Player(tradeTarget).y
    sX = Player(Index).x
    sY = Player(Index).y
    
    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg Index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg Index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    
    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg Index, "This player is busy.", BrightRed
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = Index
    SendTradeRequest tradeTarget, Index
End Sub

Public Sub HandleAcceptTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long

    tradeTarget = TempPlayer(Index).TradeRequest
    
    If Not IsConnected(Index) Or Not IsPlaying(Index) Then
        TempPlayer(tradeTarget).TradeRequest = 0
        Exit Sub
    End If
    
    If Not IsConnected(tradeTarget) Or Not IsPlaying(tradeTarget) Then
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If
    
    If TempPlayer(Index).TradeRequest <= 0 Or TempPlayer(Index).TradeRequest > MAX_PLAYERS Then Exit Sub
    ' let them know they're trading
    PlayerMsg Index, "You have accepted " & Trim$(GetPlayerName(tradeTarget)) & "'s trade request.", BrightGreen
    PlayerMsg tradeTarget, Trim$(GetPlayerName(Index)) & " has accepted your trade request.", BrightGreen
    ' clear the tradeRequest server-side
    TempPlayer(Index).TradeRequest = 0
    TempPlayer(tradeTarget).TradeRequest = 0
    ' set that they're trading with each other
    TempPlayer(Index).InTrade = tradeTarget
    TempPlayer(tradeTarget).InTrade = Index
    ' clear out their trade offers
    For i = 1 To MAX_INV
        TempPlayer(Index).TradeOffer(i).Num = 0
        TempPlayer(Index).TradeOffer(i).Value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).Value = 0
    Next
    ' Used to init the trade window clientside
    SendTrade Index, tradeTarget
    SendTrade tradeTarget, Index
    ' Send the offer data - Used to clear their client
    SendTradeUpdate Index, 0
    SendTradeUpdate Index, 1
    SendTradeUpdate tradeTarget, 0
    SendTradeUpdate tradeTarget, 1
End Sub

Public Sub HandleDeclineTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PlayerMsg TempPlayer(Index).TradeRequest, GetPlayerName(Index) & " has declined your trade request.", BrightRed
    PlayerMsg Index, "You decline the trade request.", BrightRed
    ' clear the tradeRequest server-side
    TempPlayer(Index).TradeRequest = 0
End Sub

Public Sub HandleAcceptTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long, x As Long
    Dim tmpTradeItem(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(1 To MAX_INV) As PlayerInvRec
    Dim itemNum As Long
    Dim theirInvSpace As Long, yourInvSpace As Long
    Dim theirItemCount As Long, yourItemCount As Long
    
    If TempPlayer(Index).InTrade = 0 Then Exit Sub
    
    TempPlayer(Index).AcceptTrade = True
    tradeTarget = TempPlayer(Index).InTrade
    
    If Not IsConnected(Index) Or Not IsPlaying(Index) Then
        TempPlayer(tradeTarget).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If
    
    If Not IsConnected(tradeTarget) Or Not IsPlaying(tradeTarget) Then
        TempPlayer(tradeTarget).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If
    
    ' if not both of them accept, then exit
    If Not TempPlayer(tradeTarget).AcceptTrade Then
        SendTradeStatus Index, 2
        SendTradeStatus tradeTarget, 1
        Exit Sub
    End If
    
    ' get inventory spaces
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(Index, i) > 0 Then
            ' check if we're offering it
            For x = 1 To MAX_INV
                If TempPlayer(Index).TradeOffer(x).Num = i Then
                    itemNum = Player(Index).Inv(TempPlayer(Index).TradeOffer(x).Num).Num
                    ' if it's a currency then make sure we're offering all of it
                    If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then
                        If TempPlayer(Index).TradeOffer(x).Value = GetPlayerInvItemNum(Index, i) Then
                            yourInvSpace = yourInvSpace + 1
                        End If
                    Else
                        yourInvSpace = yourInvSpace + 1
                    End If
                End If
            Next
        Else
            yourInvSpace = yourInvSpace + 1
        End If
        If GetPlayerInvItemNum(tradeTarget, i) > 0 Then
            ' check if we're offering it
            For x = 1 To MAX_INV
                If TempPlayer(tradeTarget).TradeOffer(x).Num = i Then
                    itemNum = Player(tradeTarget).Inv(TempPlayer(tradeTarget).TradeOffer(x).Num).Num
                    ' if it's a currency then make sure we're offering all of it
                    If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then
                        If TempPlayer(tradeTarget).TradeOffer(x).Value = GetPlayerInvItemNum(tradeTarget, i) Then
                            theirInvSpace = theirInvSpace + 1
                        End If
                    Else
                        theirInvSpace = theirInvSpace + 1
                    End If
                End If
            Next
        Else
            theirInvSpace = theirInvSpace + 1
        End If
    Next
    
    ' get item count
    For i = 1 To MAX_INV
        If TempPlayer(Index).TradeOffer(i).Num > 0 Then
            itemNum = Player(Index).Inv(TempPlayer(Index).TradeOffer(i).Num).Num
            If itemNum > 0 Then
                If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then
                    ' check if the other player has the item
                    If HasItem(tradeTarget, itemNum) = 0 Then
                        yourItemCount = yourItemCount + 1
                    End If
                Else
                    yourItemCount = yourItemCount + 1
                End If
            End If
        End If
        If TempPlayer(tradeTarget).TradeOffer(i).Num > 0 Then
            itemNum = Player(tradeTarget).Inv(TempPlayer(tradeTarget).TradeOffer(i).Num).Num
            If itemNum > 0 Then
                If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then
                    ' check if the other player has the item
                    If HasItem(Index, itemNum) = 0 Then
                        theirItemCount = theirItemCount + 1
                    End If
                Else
                    theirItemCount = theirItemCount + 1
                End If
            End If
        End If
    Next
    
    ' make sure they have enough space
    If yourInvSpace < theirItemCount Then
        PlayerMsg Index, "You don't have enough inventory space.", BrightRed
        PlayerMsg tradeTarget, "They don't have enough inventory space.", BrightRed
        TempPlayer(Index).AcceptTrade = False
        TempPlayer(tradeTarget).AcceptTrade = False
        SendTradeUpdate Index, 0
        SendTradeUpdate tradeTarget, 0
        SendTradeStatus Index, 3
        SendTradeStatus tradeTarget, 3
        Exit Sub
    End If
    If theirInvSpace < yourItemCount Then
        PlayerMsg Index, "They don't have enough inventory space.", BrightRed
        PlayerMsg tradeTarget, "You don't have enough inventory space.", BrightRed
        TempPlayer(Index).AcceptTrade = False
        TempPlayer(tradeTarget).AcceptTrade = False
        SendTradeUpdate Index, 0
        SendTradeUpdate tradeTarget, 0
        SendTradeStatus Index, 3
        SendTradeStatus tradeTarget, 3
        Exit Sub
    End If
    
    ' take their items
    For i = 1 To MAX_INV
        ' player
        If TempPlayer(Index).TradeOffer(i).Num > 0 Then
            itemNum = Player(Index).Inv(TempPlayer(Index).TradeOffer(i).Num).Num
            If itemNum > 0 Then
                ' store temp
                tmpTradeItem(i).Num = itemNum
                tmpTradeItem(i).Value = TempPlayer(Index).TradeOffer(i).Value
                ' take item
                TakeInvSlot Index, TempPlayer(Index).TradeOffer(i).Num, tmpTradeItem(i).Value
            End If
        End If
        ' target
        If TempPlayer(tradeTarget).TradeOffer(i).Num > 0 Then
            itemNum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            If itemNum > 0 Then
                ' store temp
                tmpTradeItem2(i).Num = itemNum
                tmpTradeItem2(i).Value = TempPlayer(tradeTarget).TradeOffer(i).Value
                ' take item
                TakeInvSlot tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, tmpTradeItem2(i).Value
            End If
        End If
    Next
    
    ' taken all items. now they can't not get items because of no inventory space.
    For i = 1 To MAX_INV
        ' player
        If tmpTradeItem2(i).Num > 0 Then
            ' give away!
            GiveInvItem Index, tmpTradeItem2(i).Num, tmpTradeItem2(i).Value, False
        End If
        ' target
        If tmpTradeItem(i).Num > 0 Then
            ' give away!
            GiveInvItem tradeTarget, tmpTradeItem(i).Num, tmpTradeItem(i).Value, False
        End If
    Next
    
    SendInventory Index
    SendInventory tradeTarget
    
    ' they now have all the items. Clear out values + let them out of the trade.
    For i = 1 To MAX_INV
        TempPlayer(Index).TradeOffer(i).Num = 0
        TempPlayer(Index).TradeOffer(i).Value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).Value = 0
    Next

    TempPlayer(Index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0
    
    PlayerMsg Index, "Trade completed.", BrightGreen
    PlayerMsg tradeTarget, "Trade completed.", BrightGreen
    
    SendCloseTrade Index
    SendCloseTrade tradeTarget
End Sub

Public Sub HandleDeclineTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim tradeTarget As Long

    tradeTarget = TempPlayer(Index).InTrade
    
    If tradeTarget = 0 Then
        SendCloseTrade Index
        Exit Sub
    End If

    For i = 1 To MAX_INV
        TempPlayer(Index).TradeOffer(i).Num = 0
        TempPlayer(Index).TradeOffer(i).Value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).Value = 0
    Next

    TempPlayer(Index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0
    
    PlayerMsg Index, "You declined the trade.", BrightRed
    PlayerMsg tradeTarget, GetPlayerName(Index) & " has declined the trade.", BrightRed
    
    SendCloseTrade Index
    SendCloseTrade tradeTarget
End Sub

Public Sub HandleTradeItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim invSlot As Long
    Dim amount As Long
    Dim EmptySlot As Long
    Dim itemNum As Long
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    invSlot = buffer.ReadLong
    amount = buffer.ReadLong
    
    buffer.Flush: Set buffer = Nothing
    
    If invSlot <= 0 Or invSlot > MAX_INV Then Exit Sub
    
    itemNum = GetPlayerInvItemNum(Index, invSlot)
    If itemNum <= 0 Or itemNum > MAX_ITEMS Then Exit Sub
    
    If TempPlayer(Index).InTrade <= 0 Or TempPlayer(Index).InTrade > MAX_PLAYERS Then Exit Sub
    
    ' make sure they have the amount they offer
    If amount < 0 Or amount > GetPlayerInvItemValue(Index, invSlot) Then
        PlayerMsg Index, "You do not have that many.", BrightRed
        Exit Sub
    End If
    
    ' make sure it's not soulbound
    If Item(itemNum).BindType > 0 Then
        If Player(Index).Inv(invSlot).Bound > 0 Then
            PlayerMsg Index, "Cannot trade a soulbound item.", BrightRed
            Exit Sub
        End If
    End If

    If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then
        ' check if already offering same currency item
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).Num = invSlot Then
                ' add amount
                TempPlayer(Index).TradeOffer(i).Value = TempPlayer(Index).TradeOffer(i).Value + amount
                ' clamp to limits
                If TempPlayer(Index).TradeOffer(i).Value > GetPlayerInvItemValue(Index, invSlot) Then
                    TempPlayer(Index).TradeOffer(i).Value = GetPlayerInvItemValue(Index, invSlot)
                End If
                ' cancel any trade agreement
                TempPlayer(Index).AcceptTrade = False
                TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
                
                SendTradeStatus Index, 0
                SendTradeStatus TempPlayer(Index).InTrade, 0
                
                SendTradeUpdate Index, 0
                SendTradeUpdate TempPlayer(Index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they're not already offering it
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).Num = invSlot Then
                PlayerMsg Index, "You've already offered this item.", BrightRed
                Exit Sub
            End If
        Next
    End If
    
    ' not already offering - find earliest empty slot
    For i = 1 To MAX_INV
        If TempPlayer(Index).TradeOffer(i).Num = 0 Then
            EmptySlot = i
            Exit For
        End If
    Next
    TempPlayer(Index).TradeOffer(EmptySlot).Num = invSlot
    TempPlayer(Index).TradeOffer(EmptySlot).Value = amount
    
    ' cancel any trade agreement and send new data
    TempPlayer(Index).AcceptTrade = False
    TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
    
    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0
    
    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1
End Sub

Public Sub HandleUntradeItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim tradeSlot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    tradeSlot = buffer.ReadLong
    
    buffer.Flush: Set buffer = Nothing
    
    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(Index).TradeOffer(tradeSlot).Num <= 0 Then Exit Sub
    
    TempPlayer(Index).TradeOffer(tradeSlot).Num = 0
    TempPlayer(Index).TradeOffer(tradeSlot).Value = 0
    
    If TempPlayer(Index).AcceptTrade Then TempPlayer(Index).AcceptTrade = False
    If TempPlayer(TempPlayer(Index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
    
    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0
    
    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1
End Sub

' :::::::::::::::::::::::
' ::    SHOP packet    ::
' :::::::::::::::::::::::

Public Sub HandleCloseShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(Index).InShop = 0
End Sub

Public Sub HandleBuyItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim shopslot As Long
    Dim shopNum As Long
    Dim itemAmount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    shopslot = buffer.ReadLong
    
    ' not in shop, exit out
    shopNum = TempPlayer(Index).InShop
    If shopNum < 1 Or shopNum > MAX_SHOPS Then Exit Sub
    
    With Shop(shopNum).TradeItem(shopslot)
        ' check trade exists
        If .Item < 1 Then Exit Sub
        
        ' make sure they have inventory space
        If FindOpenInvSlot(Index, .Item) = 0 Then
            PlayerMsg Index, "You do not have enough inventory space.", BrightRed
            ResetShopAction Index
            Exit Sub
        End If
            
        ' check has the cost item
        itemAmount = HasItem(Index, .costitem)
        If itemAmount = 0 Or itemAmount < .costvalue Then
            PlayerMsg Index, "You do not have enough to buy this item.", BrightRed
            ResetShopAction Index
            Exit Sub
        End If
        
        ' it's fine, let's go ahead
        TakeInvItem Index, .costitem, .costvalue
        GiveInvItem Index, .Item, .ItemValue
        
        PlayerMsg Index, "You successfully bought " & Trim$(Item(.Item).name) & " for " & .costvalue & " " & Trim$(Item(.costitem).name) & ".", BrightGreen
    End With
    
    ' send confirmation message & reset their shop action
    'PlayerMsg index, "Trade successful.", BrightGreen
    
    ResetShopAction Index
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub HandleSellItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim invSlot As Long
    Dim itemNum As Long
    Dim price As Long
    Dim multiplier As Double
    Dim amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    invSlot = buffer.ReadLong
    
    If TempPlayer(Index).InShop = 0 Then Exit Sub
    
    ' if invalid, exit out
    If invSlot < 1 Or invSlot > MAX_INV Then Exit Sub
    
    ' has item?
    If GetPlayerInvItemNum(Index, invSlot) < 1 Or GetPlayerInvItemNum(Index, invSlot) > MAX_ITEMS Then Exit Sub
    
    ' seems to be valid
    itemNum = GetPlayerInvItemNum(Index, invSlot)
    
    ' work out price
    multiplier = Shop(TempPlayer(Index).InShop).BuyRate / 100
    price = Item(itemNum).price * multiplier
    
    ' item has cost?
    If price <= 0 Then
        PlayerMsg Index, "The shop doesn't want that item.", BrightRed
        ResetShopAction Index
        Exit Sub
    End If

    ' take item and give gold
    TakeInvItem Index, itemNum, 1
    GiveInvItem Index, 1, price
    
    ' send confirmation message & reset their shop action
    PlayerMsg Index, "Trade successful.", BrightGreen
    ResetShopAction Index
    
    buffer.Flush: Set buffer = Nothing
End Sub

' :::::::::::::::::::::::
' ::    BANK packet    ::
' :::::::::::::::::::::::

Public Sub HandleChangeBankSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim newSlot As Long
    Dim oldSlot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    oldSlot = buffer.ReadLong
    newSlot = buffer.ReadLong
    
    PlayerSwitchBankSlots Index, oldSlot, newSlot
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub HandleWithdrawItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim BankSlot As Long
    Dim amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    BankSlot = buffer.ReadLong
    amount = buffer.ReadLong
    
    TakeBankItem Index, BankSlot, amount
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub HandleDepositItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim invSlot As Long
    Dim amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    invSlot = buffer.ReadLong
    amount = buffer.ReadLong
    
    GiveBankItem Index, invSlot, amount
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub HandleCloseBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(Index) Then
        Exit Sub
    End If
    
    If TempPlayer(Index).InBank Then
        SavePlayer Index
    
        TempPlayer(Index).InBank = False
    End If
End Sub

Public Sub HandleFinishTutorial(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Player(Index).TutorialState = 1
    SavePlayer Index
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Public Sub HandleQuit(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(Index)
End Sub
