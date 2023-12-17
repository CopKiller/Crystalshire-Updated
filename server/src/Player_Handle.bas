Attribute VB_Name = "Player_Handle"
Public Sub HandleRequestPlayerData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerData(index)
End Sub

Public Sub HandleRequestLevelUp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If GetPlayerAccess(index) < 4 Then Exit Sub
    Call SetPlayerExp(index, GetPlayerNextLevel(index))
    Call CheckPlayerLevelUp(index)
End Sub
' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String
    Dim i As Long
    Dim N As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Name = Buffer.ReadString 'Parse(1)
    Buffer.Flush: Set Buffer = Nothing
    i = FindPlayer(Name)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Public Sub HandlePlayerMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim movement As Long
    Dim Buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong 'CLng(Parse(1))
    movement = Buffer.ReadLong 'CLng(Parse(2))
    tmpX = Buffer.ReadLong
    tmpY = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Then
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
    If TempPlayer(index).InBank Then
        'Call SendPlayerXY(Index)
        'Exit Sub
        TempPlayer(index).InBank = False
    End If

    ' if stunned, stop them moving
    If TempPlayer(index).StunDuration > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If
    
    ' Prever player from moving if in shop
    If TempPlayer(index).InShop > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(index) <> tmpX Then
        SendPlayerXY (index)
        Exit Sub
    End If
    
    If GetPlayerY(index) <> tmpY Then
        SendPlayerXY (index)
        Exit Sub
    End If
    
    ' cant move if chatting
    If TempPlayer(index).inChatWith > 0 Then
        ClosePlayerChat index
    End If
    
    Call PlayerMove(index, Dir, movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Public Sub HandlePlayerDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong 'CLng(Parse(1))
    Buffer.Flush: Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerDir
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerDir(index)
    
    SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Public Sub HandleAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, N As Long, damage As Long, TempIndex As Long, x As Long, y As Long, mapnum As Long, dirReq As Long
    
    ' can't attack whilst casting
    If TempPlayer(index).spellBuffer.Spell > 0 Then Exit Sub
    
    ' can't attack whilst stunned
    If TempPlayer(index).StunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
    SendAttack index

    ' Try to attack a player
    For i = 1 To Player_HighIndex
        TempIndex = i

        ' Make sure we dont try to attack ourselves
        If TempIndex <> index Then
            TryPlayerAttackPlayer index, i
        End If
    Next

    ' Try to attack a npc
    For i = 1 To MAX_MAP_NPCS
        TryPlayerAttackNpc index, i
    Next
    
    ' check if we've got a remote chat tile
    mapnum = GetPlayerMap(index)
    x = GetPlayerX(index)
    y = GetPlayerY(index)
    If Map(mapnum).TileData.Tile(x, y).Type = TILE_TYPE_CHAT Then
        dirReq = Map(mapnum).TileData.Tile(x, y).Data2
        If Player(index).Dir = dirReq Then
            InitChat index, mapnum, Map(mapnum).TileData.Tile(x, y).Data1, True
            Exit Sub
        End If
    End If

    ' Check tradeskills
    Select Case GetPlayerDir(index)
        Case DIR_UP

            If GetPlayerY(index) = 0 Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1
        Case DIR_DOWN

            If GetPlayerY(index) = Map(GetPlayerMap(index)).MapData.MaxY Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1
        Case DIR_LEFT

            If GetPlayerX(index) = 0 Then Exit Sub
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)
        Case DIR_RIGHT

            If GetPlayerX(index) = Map(GetPlayerMap(index)).MapData.MaxX Then Exit Sub
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)
    End Select
    
    CheckResource index, x, y
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Public Sub HandleCast(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Spell slot
    N = Buffer.ReadLong 'CLng(Parse(1))
    Buffer.Flush: Set Buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(index, N)
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Public Sub HandleUseItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim invNum As Long
    Dim Buffer As clsBuffer
    
    ' get inventory slot number
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    invNum = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing

    UseItem index, invNum
End Sub

Public Sub HandleUnequip(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PlayerUnequipItem index, Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub HandleSpawnItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tmpItem As Long
    Dim tmpAmount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' item
    tmpItem = Buffer.ReadLong
    tmpAmount = Buffer.ReadLong
        
    If GetPlayerAccess(index) < ADMIN_CREATOR Then Exit Sub
    
    SpawnItem tmpItem, tmpAmount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), GetPlayerName(index)
    Buffer.Flush: Set Buffer = Nothing
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Public Sub HandleTarget(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, target As Long, targetType As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    target = Buffer.ReadLong
    targetType = Buffer.ReadLong
    
    Buffer.Flush: Set Buffer = Nothing
    
    ' set player's target - no need to send, it's client side
    TempPlayer(index).target = target
    TempPlayer(index).targetType = targetType
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Public Sub HandleSpells(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(index)
End Sub

Public Sub HandleForgetSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim spellSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    spellSlot = Buffer.ReadLong
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If TempPlayer(index).SpellCD(spellSlot) > GetTickCount Then
        PlayerMsg index, "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(index).spellBuffer.Spell = spellSlot Then
        PlayerMsg index, "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    Player(index).Spell(spellSlot).Spell = 0
    Player(index).Spell(spellSlot).Uses = 0
    SendPlayerSpells index
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Public Sub HandleUseStatPoint(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim PointType As Byte
    Dim Buffer As clsBuffer
    Dim sMes As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PointType = Buffer.ReadByte 'CLng(Parse(1))
    Buffer.Flush: Set Buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_Count) Then
        Exit Sub
    End If

    ' Make sure they have points
    If GetPlayerPOINTS(index) > 0 Then
        ' make sure they're not maxed
        If GetPlayerRawStat(index, PointType) >= 255 Then
            PlayerMsg index, "You cannot spend any more points on that stat.", BrightRed
            Exit Sub
        End If
        
        ' make sure they're not spending too much
        If GetPlayerRawStat(index, PointType) - Class(GetPlayerClass(index)).Stat(PointType) >= (GetPlayerLevel(index) * 2) - 1 Then
            PlayerMsg index, "You cannot spend any more points on that stat.", BrightRed
            Exit Sub
        End If
        
        ' Take away a stat point
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)

        ' Everything is ok
        Select Case PointType
            Case Stats.Strength
                Call SetPlayerStat(index, Stats.Strength, GetPlayerRawStat(index, Stats.Strength) + 1)
                sMes = "Strength"
            Case Stats.Endurance
                Call SetPlayerStat(index, Stats.Endurance, GetPlayerRawStat(index, Stats.Endurance) + 1)
                sMes = "Endurance"
            Case Stats.Intelligence
                Call SetPlayerStat(index, Stats.Intelligence, GetPlayerRawStat(index, Stats.Intelligence) + 1)
                sMes = "Intelligence"
            Case Stats.Agility
                Call SetPlayerStat(index, Stats.Agility, GetPlayerRawStat(index, Stats.Agility) + 1)
                sMes = "Agility"
            Case Stats.Willpower
                Call SetPlayerStat(index, Stats.Willpower, GetPlayerRawStat(index, Stats.Willpower) + 1)
                sMes = "Willpower"
        End Select
        
        SendActionMsg GetPlayerMap(index), "+1 " & sMes, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
    Else
        Exit Sub
    End If

    ' Send the update
    SendPlayerData index
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Public Sub HandleSwapInvSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    
    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).InBank Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    PlayerSwitchInvSlots index, oldSlot, newSlot
End Sub

Public Sub HandleSwapSpellSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long, N As Long
    
    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub
    
    If TempPlayer(index).spellBuffer.Spell > 0 Then
        PlayerMsg index, "You cannot swap spells whilst casting.", BrightRed
        Exit Sub
    End If
    
    For N = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(index).SpellCD(N) > GetTickCount Then
            PlayerMsg index, "You cannot swap spells whilst they're cooling down.", BrightRed
            Exit Sub
        End If
    Next
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    PlayerSwitchSpellSlots index, oldSlot, newSlot
End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Public Sub HandleWhosOnline(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(index)
End Sub

' :::::::::::::::::::::::
' ::    Party packet   ::
' :::::::::::::::::::::::

Public Sub HandlePartyRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, targetIndex As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    targetIndex = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    
    ' make sure it's a valid target
    If targetIndex = index Then
        PlayerMsg index, "You can't invite yourself. That would be weird.", BrightRed
        Exit Sub
    End If
    
    ' make sure they're connected and on the same map
    If Not IsConnected(targetIndex) Or Not IsPlaying(targetIndex) Then Exit Sub
    If GetPlayerMap(targetIndex) <> GetPlayerMap(index) Then Exit Sub
    
    ' init the request
    Party_Invite index, targetIndex
End Sub

Public Sub HandleAcceptParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteAccept TempPlayer(index).partyInvite, index
End Sub

Public Sub HandleDeclineParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteDecline TempPlayer(index).partyInvite, index
End Sub

Public Sub HandlePartyLeave(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_PlayerLeave index
End Sub

' :::::::::::::::::::::::
' ::   HOTBAR packet   ::
' :::::::::::::::::::::::

Public Sub HandleHotbarChange(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim sType As Long
    Dim Slot As Long
    Dim hotbarNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    sType = Buffer.ReadLong
    Slot = Buffer.ReadLong
    hotbarNum = Buffer.ReadLong
    
    Select Case sType
        Case 0 ' clear
            Player(index).Hotbar(hotbarNum).Slot = 0
            Player(index).Hotbar(hotbarNum).sType = 0
        Case 1 ' inventory
            If Slot > 0 And Slot <= MAX_INV Then
                If Player(index).Inv(Slot).Num > 0 Then
                    If Len(Trim$(Item(GetPlayerInvItemNum(index, Slot)).Name)) > 0 Then
                        Player(index).Hotbar(hotbarNum).Slot = Player(index).Inv(Slot).Num
                        Player(index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
        Case 2 ' spell
            If Slot > 0 And Slot <= MAX_PLAYER_SPELLS Then
                If Player(index).Spell(Slot).Spell > 0 Then
                    If Len(Trim$(Spell(Player(index).Spell(Slot).Spell).Name)) > 0 Then
                        Player(index).Hotbar(hotbarNum).Slot = Player(index).Spell(Slot).Spell
                        Player(index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
    End Select
    
    SendHotbar index
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub HandleHotbarUse(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Slot As Long
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Slot = Buffer.ReadLong
    
    Select Case Player(index).Hotbar(Slot).sType
        Case 1 ' inventory
            For i = 1 To MAX_INV
                If Player(index).Inv(i).Num > 0 Then
                    If Player(index).Inv(i).Num = Player(index).Hotbar(Slot).Slot Then
                        UseItem index, i
                        Exit Sub
                    End If
                End If
            Next
        Case 2 ' spell
            For i = 1 To MAX_PLAYER_SPELLS
                If Player(index).Spell(i).Spell > 0 Then
                    If Player(index).Spell(i).Spell = Player(index).Hotbar(Slot).Slot Then
                        BufferSpell index, i
                        Exit Sub
                    End If
                End If
            Next
    End Select
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

' :::::::::::::::::::::::
' ::   MISSION packet  ::
' :::::::::::::::::::::::

Public Sub HandleAcceptMissionRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim MissionID As Long, i As Long
    Dim ItemNum As Long, ItemCount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MissionID = Buffer.ReadLong
    If MissionID <= 0 Or MissionID > MAX_MISSIONS Then Exit Sub
    
    For i = 1 To MAX_PLAYER_MISSIONS
        If Player(index).Mission(i).ID = 0 Then
            Player(index).Mission(i).ID = MissionID
            Player(index).Mission(i).Count = 0
            'Let's check for any existing items in their inventory and update their counter!
            If Mission(MissionID).Type = MissionType.Mission_TypeCollect Then
                ItemNum = Mission(MissionID).CollectItem
                ItemCount = HasItem(index, ItemNum)
                If ItemCount > 0 Then
                    Player(index).Mission(i).Count = ItemCount
                End If
            End If
            TempPlayer(index).MissionRequest = 0
            Call PlayerMsg(index, "Mission Accepted!", Yellow)
            Call SendPlayerData(index)
            Exit Sub
        End If
    Next i
    Set Buffer = Nothing
End Sub

Public Sub HandleDeclineMissionRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' clear the missionRequest server-side
    TempPlayer(index).MissionRequest = 0
End Sub

' :::::::::::::::::::::::
' ::    TRADE packet   ::
' :::::::::::::::::::::::

Public Sub HandleTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long, Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' find the target
    tradeTarget = Buffer.ReadLong
    
    Buffer.Flush: Set Buffer = Nothing
    
    If Not IsConnected(index) Or Not IsPlaying(index) Then
        TempPlayer(tradeTarget).TradeRequest = 0
        TempPlayer(index).TradeRequest = 0
        Exit Sub
    End If
    
    If Not IsConnected(tradeTarget) Or Not IsPlaying(tradeTarget) Then
        TempPlayer(tradeTarget).TradeRequest = 0
        TempPlayer(index).TradeRequest = 0
        Exit Sub
    End If
    
    ' make sure we don't error
    If tradeTarget <= 0 Or tradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' can't trade with yourself..
    If tradeTarget = index Then
        PlayerMsg index, "You can't trade with yourself.", BrightRed
        Exit Sub
    End If
    
    ' make sure they're on the same map
    If Not Player(tradeTarget).Map = Player(index).Map Then Exit Sub
    
    ' make sure they're stood next to each other
    tX = Player(tradeTarget).x
    tY = Player(tradeTarget).y
    sX = Player(index).x
    sY = Player(index).y
    
    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    
    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg index, "This player is busy.", BrightRed
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = index
    SendTradeRequest tradeTarget, index
End Sub

Public Sub HandleAcceptTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long

    tradeTarget = TempPlayer(index).TradeRequest
    
    If Not IsConnected(index) Or Not IsPlaying(index) Then
        TempPlayer(tradeTarget).TradeRequest = 0
        Exit Sub
    End If
    
    If Not IsConnected(tradeTarget) Or Not IsPlaying(tradeTarget) Then
        TempPlayer(index).TradeRequest = 0
        Exit Sub
    End If
    
    If TempPlayer(index).TradeRequest <= 0 Or TempPlayer(index).TradeRequest > MAX_PLAYERS Then Exit Sub
    ' let them know they're trading
    PlayerMsg index, "You have accepted " & Trim$(GetPlayerName(tradeTarget)) & "'s trade request.", BrightGreen
    PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " has accepted your trade request.", BrightGreen
    ' clear the tradeRequest server-side
    TempPlayer(index).TradeRequest = 0
    TempPlayer(tradeTarget).TradeRequest = 0
    ' set that they're trading with each other
    TempPlayer(index).InTrade = tradeTarget
    TempPlayer(tradeTarget).InTrade = index
    ' clear out their trade offers
    For i = 1 To MAX_INV
        TempPlayer(index).TradeOffer(i).Num = 0
        TempPlayer(index).TradeOffer(i).Value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).Value = 0
    Next
    ' Used to init the trade window clientside
    SendTrade index, tradeTarget
    SendTrade tradeTarget, index
    ' Send the offer data - Used to clear their client
    SendTradeUpdate index, 0
    SendTradeUpdate index, 1
    SendTradeUpdate tradeTarget, 0
    SendTradeUpdate tradeTarget, 1
End Sub

Public Sub HandleDeclineTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PlayerMsg TempPlayer(index).TradeRequest, GetPlayerName(index) & " has declined your trade request.", BrightRed
    PlayerMsg index, "You decline the trade request.", BrightRed
    ' clear the tradeRequest server-side
    TempPlayer(index).TradeRequest = 0
End Sub

Public Sub HandleAcceptTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long, x As Long
    Dim tmpTradeItem(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(1 To MAX_INV) As PlayerInvRec
    Dim ItemNum As Long
    Dim theirInvSpace As Long, yourInvSpace As Long
    Dim theirItemCount As Long, yourItemCount As Long
    
    If TempPlayer(index).InTrade = 0 Then Exit Sub
    
    TempPlayer(index).AcceptTrade = True
    tradeTarget = TempPlayer(index).InTrade
    
    If Not IsConnected(index) Or Not IsPlaying(index) Then
        TempPlayer(tradeTarget).TradeRequest = 0
        TempPlayer(index).TradeRequest = 0
        Exit Sub
    End If
    
    If Not IsConnected(tradeTarget) Or Not IsPlaying(tradeTarget) Then
        TempPlayer(tradeTarget).TradeRequest = 0
        TempPlayer(index).TradeRequest = 0
        Exit Sub
    End If
    
    ' if not both of them accept, then exit
    If Not TempPlayer(tradeTarget).AcceptTrade Then
        SendTradeStatus index, 2
        SendTradeStatus tradeTarget, 1
        Exit Sub
    End If
    
    ' get inventory spaces
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(index, i) > 0 Then
            ' check if we're offering it
            For x = 1 To MAX_INV
                If TempPlayer(index).TradeOffer(x).Num = i Then
                    ItemNum = Player(index).Inv(TempPlayer(index).TradeOffer(x).Num).Num
                    ' if it's a currency then make sure we're offering all of it
                    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                        If TempPlayer(index).TradeOffer(x).Value = GetPlayerInvItemNum(index, i) Then
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
                    ItemNum = Player(tradeTarget).Inv(TempPlayer(tradeTarget).TradeOffer(x).Num).Num
                    ' if it's a currency then make sure we're offering all of it
                    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
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
        If TempPlayer(index).TradeOffer(i).Num > 0 Then
            ItemNum = Player(index).Inv(TempPlayer(index).TradeOffer(i).Num).Num
            If ItemNum > 0 Then
                If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                    ' check if the other player has the item
                    If HasItem(tradeTarget, ItemNum) = 0 Then
                        yourItemCount = yourItemCount + 1
                    End If
                Else
                    yourItemCount = yourItemCount + 1
                End If
            End If
        End If
        If TempPlayer(tradeTarget).TradeOffer(i).Num > 0 Then
            ItemNum = Player(tradeTarget).Inv(TempPlayer(tradeTarget).TradeOffer(i).Num).Num
            If ItemNum > 0 Then
                If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                    ' check if the other player has the item
                    If HasItem(index, ItemNum) = 0 Then
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
        PlayerMsg index, "You don't have enough inventory space.", BrightRed
        PlayerMsg tradeTarget, "They don't have enough inventory space.", BrightRed
        TempPlayer(index).AcceptTrade = False
        TempPlayer(tradeTarget).AcceptTrade = False
        SendTradeUpdate index, 0
        SendTradeUpdate tradeTarget, 0
        SendTradeStatus index, 3
        SendTradeStatus tradeTarget, 3
        Exit Sub
    End If
    If theirInvSpace < yourItemCount Then
        PlayerMsg index, "They don't have enough inventory space.", BrightRed
        PlayerMsg tradeTarget, "You don't have enough inventory space.", BrightRed
        TempPlayer(index).AcceptTrade = False
        TempPlayer(tradeTarget).AcceptTrade = False
        SendTradeUpdate index, 0
        SendTradeUpdate tradeTarget, 0
        SendTradeStatus index, 3
        SendTradeStatus tradeTarget, 3
        Exit Sub
    End If
    
    ' take their items
    For i = 1 To MAX_INV
        ' player
        If TempPlayer(index).TradeOffer(i).Num > 0 Then
            ItemNum = Player(index).Inv(TempPlayer(index).TradeOffer(i).Num).Num
            If ItemNum > 0 Then
                ' store temp
                tmpTradeItem(i).Num = ItemNum
                tmpTradeItem(i).Value = TempPlayer(index).TradeOffer(i).Value
                ' take item
                TakeInvSlot index, TempPlayer(index).TradeOffer(i).Num, tmpTradeItem(i).Value
            End If
        End If
        ' target
        If TempPlayer(tradeTarget).TradeOffer(i).Num > 0 Then
            ItemNum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            If ItemNum > 0 Then
                ' store temp
                tmpTradeItem2(i).Num = ItemNum
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
            GiveInvItem index, tmpTradeItem2(i).Num, tmpTradeItem2(i).Value, False
        End If
        ' target
        If tmpTradeItem(i).Num > 0 Then
            ' give away!
            GiveInvItem tradeTarget, tmpTradeItem(i).Num, tmpTradeItem(i).Value, False
        End If
    Next
    
    SendInventory index
    SendInventory tradeTarget
    
    ' they now have all the items. Clear out values + let them out of the trade.
    For i = 1 To MAX_INV
        TempPlayer(index).TradeOffer(i).Num = 0
        TempPlayer(index).TradeOffer(i).Value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).Value = 0
    Next

    TempPlayer(index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0
    
    PlayerMsg index, "Trade completed.", BrightGreen
    PlayerMsg tradeTarget, "Trade completed.", BrightGreen
    
    SendCloseTrade index
    SendCloseTrade tradeTarget
End Sub

Public Sub HandleDeclineTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim tradeTarget As Long

    tradeTarget = TempPlayer(index).InTrade
    
    If tradeTarget = 0 Then
        SendCloseTrade index
        Exit Sub
    End If

    For i = 1 To MAX_INV
        TempPlayer(index).TradeOffer(i).Num = 0
        TempPlayer(index).TradeOffer(i).Value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).Value = 0
    Next

    TempPlayer(index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0
    
    PlayerMsg index, "You declined the trade.", BrightRed
    PlayerMsg tradeTarget, GetPlayerName(index) & " has declined the trade.", BrightRed
    
    SendCloseTrade index
    SendCloseTrade tradeTarget
End Sub

Public Sub HandleTradeItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim amount As Long
    Dim EmptySlot As Long
    Dim ItemNum As Long
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    amount = Buffer.ReadLong
    
    Buffer.Flush: Set Buffer = Nothing
    
    If invSlot <= 0 Or invSlot > MAX_INV Then Exit Sub
    
    ItemNum = GetPlayerInvItemNum(index, invSlot)
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    If TempPlayer(index).InTrade <= 0 Or TempPlayer(index).InTrade > MAX_PLAYERS Then Exit Sub
    
    ' make sure they have the amount they offer
    If amount < 0 Or amount > GetPlayerInvItemValue(index, invSlot) Then
        PlayerMsg index, "You do not have that many.", BrightRed
        Exit Sub
    End If
    
    ' make sure it's not soulbound
    If Item(ItemNum).BindType > 0 Then
        If Player(index).Inv(invSlot).Bound > 0 Then
            PlayerMsg index, "Cannot trade a soulbound item.", BrightRed
            Exit Sub
        End If
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        ' check if already offering same currency item
        For i = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(i).Num = invSlot Then
                ' add amount
                TempPlayer(index).TradeOffer(i).Value = TempPlayer(index).TradeOffer(i).Value + amount
                ' clamp to limits
                If TempPlayer(index).TradeOffer(i).Value > GetPlayerInvItemValue(index, invSlot) Then
                    TempPlayer(index).TradeOffer(i).Value = GetPlayerInvItemValue(index, invSlot)
                End If
                ' cancel any trade agreement
                TempPlayer(index).AcceptTrade = False
                TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
                
                SendTradeStatus index, 0
                SendTradeStatus TempPlayer(index).InTrade, 0
                
                SendTradeUpdate index, 0
                SendTradeUpdate TempPlayer(index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they're not already offering it
        For i = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(i).Num = invSlot Then
                PlayerMsg index, "You've already offered this item.", BrightRed
                Exit Sub
            End If
        Next
    End If
    
    ' not already offering - find earliest empty slot
    For i = 1 To MAX_INV
        If TempPlayer(index).TradeOffer(i).Num = 0 Then
            EmptySlot = i
            Exit For
        End If
    Next
    TempPlayer(index).TradeOffer(EmptySlot).Num = invSlot
    TempPlayer(index).TradeOffer(EmptySlot).Value = amount
    
    ' cancel any trade agreement and send new data
    TempPlayer(index).AcceptTrade = False
    TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

Public Sub HandleUntradeItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tradeSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    tradeSlot = Buffer.ReadLong
    
    Buffer.Flush: Set Buffer = Nothing
    
    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(index).TradeOffer(tradeSlot).Num <= 0 Then Exit Sub
    
    TempPlayer(index).TradeOffer(tradeSlot).Num = 0
    TempPlayer(index).TradeOffer(tradeSlot).Value = 0
    
    If TempPlayer(index).AcceptTrade Then TempPlayer(index).AcceptTrade = False
    If TempPlayer(TempPlayer(index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

' :::::::::::::::::::::::
' ::    SHOP packet    ::
' :::::::::::::::::::::::

Public Sub HandleCloseShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(index).InShop = 0
End Sub

Public Sub HandleBuyItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim shopslot As Long
    Dim shopNum As Long
    Dim ItemAmount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    shopslot = Buffer.ReadLong
    
    ' not in shop, exit out
    shopNum = TempPlayer(index).InShop
    If shopNum < 1 Or shopNum > MAX_SHOPS Then Exit Sub
    
    With Shop(shopNum).TradeItem(shopslot)
        ' check trade exists
        If .Item < 1 Then Exit Sub
        
        ' make sure they have inventory space
        If FindOpenInvSlot(index, .Item) = 0 Then
            PlayerMsg index, "You do not have enough inventory space.", BrightRed
            ResetShopAction index
            Exit Sub
        End If
            
        ' check has the cost item
        ItemAmount = HasItem(index, .costitem)
        If ItemAmount = 0 Or ItemAmount < .costvalue Then
            PlayerMsg index, "You do not have enough to buy this item.", BrightRed
            ResetShopAction index
            Exit Sub
        End If
        
        ' it's fine, let's go ahead
        TakeInvItem index, .costitem, .costvalue
        GiveInvItem index, .Item, .ItemValue
        
        PlayerMsg index, "You successfully bought " & Trim$(Item(.Item).Name) & " for " & .costvalue & " " & Trim$(Item(.costitem).Name) & ".", BrightGreen
    End With
    
    ' send confirmation message & reset their shop action
    'PlayerMsg index, "Trade successful.", BrightGreen
    
    ResetShopAction index
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub HandleSellItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim ItemNum As Long
    Dim price As Long
    Dim multiplier As Double
    Dim amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    
    If TempPlayer(index).InShop = 0 Then Exit Sub
    
    ' if invalid, exit out
    If invSlot < 1 Or invSlot > MAX_INV Then Exit Sub
    
    ' has item?
    If GetPlayerInvItemNum(index, invSlot) < 1 Or GetPlayerInvItemNum(index, invSlot) > MAX_ITEMS Then Exit Sub
    
    ' seems to be valid
    ItemNum = GetPlayerInvItemNum(index, invSlot)
    
    ' work out price
    multiplier = Shop(TempPlayer(index).InShop).BuyRate / 100
    price = Item(ItemNum).price * multiplier
    
    ' item has cost?
    If price <= 0 Then
        PlayerMsg index, "The shop doesn't want that item.", BrightRed
        ResetShopAction index
        Exit Sub
    End If

    ' take item and give gold
    TakeInvItem index, ItemNum, 1
    GiveInvItem index, 1, price
    
    ' send confirmation message & reset their shop action
    PlayerMsg index, "Trade successful.", BrightGreen
    ResetShopAction index
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

' :::::::::::::::::::::::
' ::    BANK packet    ::
' :::::::::::::::::::::::

Public Sub HandleChangeBankSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim newSlot As Long
    Dim oldSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    
    PlayerSwitchBankSlots index, oldSlot, newSlot
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub HandleWithdrawItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim BankSlot As Long
    Dim amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    BankSlot = Buffer.ReadLong
    amount = Buffer.ReadLong
    
    TakeBankItem index, BankSlot, amount
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub HandleDepositItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    amount = Buffer.ReadLong
    
    GiveBankItem index, invSlot, amount
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub HandleCloseBank(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(index) Then
        Exit Sub
    End If
    
    If TempPlayer(index).InBank Then
        SavePlayer index
    
        TempPlayer(index).InBank = False
    End If
End Sub

Public Sub HandleFinishTutorial(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Player(index).TutorialState = 1
    SavePlayer index
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Public Sub HandleQuit(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(index)
End Sub
