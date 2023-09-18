Attribute VB_Name = "Player_Handle"
' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Public Sub HandlePlayerMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim dir As Long
    Dim movement As Long
    Dim Buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then
        Exit Sub
    End If

    dir = Buffer.ReadLong 'CLng(Parse(1))
    movement = Buffer.ReadLong 'CLng(Parse(2))
    tmpX = Buffer.ReadLong
    tmpY = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing

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
    
    Call PlayerMove(index, dir, movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Public Sub HandlePlayerDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then
        Exit Sub
    End If

    dir = Buffer.ReadLong 'CLng(Parse(1))
    Buffer.Flush: Set Buffer = Nothing

    ' Prevent hacking
    If dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(index, dir)
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerDir
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerDir(index)
    SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Public Sub HandleAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, n As Long, damage As Long, TempIndex As Long, x As Long, y As Long, mapnum As Long, dirReq As Long
    
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
        If Player(index).dir = dirReq Then
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
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Spell slot
    n = Buffer.ReadLong 'CLng(Parse(1))
    Buffer.Flush: Set Buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(index, n)
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
    Dim n As Long
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
    Dim oldSlot As Long, newSlot As Long, n As Long
    
    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub
    
    If TempPlayer(index).spellBuffer.Spell > 0 Then
        PlayerMsg index, "You cannot swap spells whilst casting.", BrightRed
        Exit Sub
    End If
    
    For n = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(index).SpellCD(n) > GetTickCount Then
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

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Public Sub HandleQuit(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(index)
End Sub
