Attribute VB_Name = "Client_Handle"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SPlayerXYMap) = GetAddress(AddressOf HandlePlayerXYMap)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SMapKey) = GetAddress(AddressOf HandleMapKey)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SResourceEditor) = GetAddress(AddressOf HandleResourceEditor)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SDoorAnimation) = GetAddress(AddressOf HandleDoorAnimation)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SBlood) = GetAddress(AddressOf HandleBlood)
    HandleDataSub(SAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SMapNpcVitals) = GetAddress(AddressOf HandleMapNpcVitals)
    HandleDataSub(SCooldown) = GetAddress(AddressOf HandleCooldown)
    HandleDataSub(SClearSpellBuffer) = GetAddress(AddressOf HandleClearSpellBuffer)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SResetShopAction) = GetAddress(AddressOf HandleResetShopAction)
    HandleDataSub(SStunned) = GetAddress(AddressOf HandleStunned)
    HandleDataSub(SMapWornEq) = GetAddress(AddressOf HandleMapWornEq)
    HandleDataSub(SBank) = GetAddress(AddressOf HandleBank)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(STradeUpdate) = GetAddress(AddressOf HandleTradeUpdate)
    HandleDataSub(STradeStatus) = GetAddress(AddressOf HandleTradeStatus)
    HandleDataSub(STarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(SHotbar) = GetAddress(AddressOf HandleHotbar)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SSound) = GetAddress(AddressOf HandleSound)
    HandleDataSub(STradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(SPartyInvite) = GetAddress(AddressOf HandlePartyInvite)
    HandleDataSub(SPartyUpdate) = GetAddress(AddressOf HandlePartyUpdate)
    HandleDataSub(SPartyVitals) = GetAddress(AddressOf HandlePartyVitals)
    HandleDataSub(SChatUpdate) = GetAddress(AddressOf HandleChatUpdate)
    HandleDataSub(SConvEditor) = GetAddress(AddressOf HandleConvEditor)
    HandleDataSub(SUpdateConv) = GetAddress(AddressOf HandleUpdateConv)
    HandleDataSub(SStartTutorial) = GetAddress(AddressOf HandleStartTutorial)
    HandleDataSub(SChatBubble) = GetAddress(AddressOf HandleChatBubble)
    HandleDataSub(SPlayerChars) = GetAddress(AddressOf HandlePlayerChars)
    HandleDataSub(SCancelAnimation) = GetAddress(AddressOf HandleCancelAnimation)
    HandleDataSub(SPlayerVariables) = GetAddress(AddressOf HandlePlayerVariables)
    HandleDataSub(SMissionEditor) = GetAddress(AddressOf HandleMissionEditor)
    HandleDataSub(SUpdateMission) = GetAddress(AddressOf HandleUpdateMission)
    HandleDataSub(SUpdateMission) = GetAddress(AddressOf HandleUpdateMission)
    HandleDataSub(SOfferMission) = GetAddress(AddressOf HandleOfferMission)
    HandleDataSub(SPlayerMission) = GetAddress(AddressOf HandlePlayerUpdateDataMission)
    HandleDataSub(SProjectileAttack) = GetAddress(AddressOf HandleProjectile)
End Sub

Sub HandleData(ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Dim MsgType As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MsgType = buffer.ReadLong

    If MsgType < 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMsgCOUNT Then
        DestroyGame
        Exit Sub
    End If

    CallWindowProc HandleDataSub(MsgType), 1, buffer.ReadBytes(buffer.length), 0, 0
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, dialogue_index As Long, menuReset As Long, kick As Long
    
    SetStatus vbNullString
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    dialogue_index = buffer.ReadLong
    menuReset = buffer.ReadLong
    kick = buffer.ReadLong
    
    buffer.Flush: Set buffer = Nothing
    
    If menuReset > 0 Then
        HideWindows
        Select Case menuReset
            Case MenuCount.menuLogin
                ShowWindow GetWindowIndex("winLogin")
            Case MenuCount.menuChars
                ShowWindow GetWindowIndex("winCharacters")
            Case MenuCount.menuClass
                ShowWindow GetWindowIndex("winClasses")
            Case MenuCount.menuNewChar
                ShowWindow GetWindowIndex("winNewChar")
            Case MenuCount.menuMain
                ShowWindow GetWindowIndex("winLogin")
            Case MenuCount.menuRegister
                ShowWindow GetWindowIndex("winRegister")
        End Select
    Else
        If kick > 0 Or inMenu = True Then
            ShowWindow GetWindowIndex("winLogin")
            DialogueAlert dialogue_index
            logoutGame
            Exit Sub
        End If
    End If
    
    DialogueAlert dialogue_index
End Sub

Sub HandleLoginOk(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Now we can receive game data
    MyIndex = buffer.ReadLong
    ' player high index
    Player_HighIndex = buffer.ReadLong
    buffer.Flush: Set buffer = Nothing
    Call SetStatus("Receiving game data.")
End Sub

Sub HandleNewCharClasses(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim z As Long, X As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = buffer.ReadString
            .Vital(Vitals.HP) = buffer.ReadLong
            .Vital(Vitals.MP) = buffer.ReadLong
            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = buffer.ReadLong
            Next

            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = buffer.ReadLong
            Next

            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = buffer.ReadLong
            Next

        End With

        n = n + 10
    Next

    buffer.Flush: Set buffer = Nothing
End Sub

Sub HandleClassesData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim z As Long, X As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong 'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = buffer.ReadString 'Trim$(Parse(n))
            .Vital(Vitals.HP) = buffer.ReadLong 'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = buffer.ReadLong 'CLng(Parse(n + 2))
            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = buffer.ReadLong
            Next

            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = buffer.ReadLong
            Next

            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = buffer.ReadLong
            Next

        End With

        n = n + 10
    Next

    buffer.Flush: Set buffer = Nothing
End Sub

Sub HandleInGame(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    InGame = True
    inMenu = False
    SetStatus vbNullString
    ' show gui
    ShowWindow GetWindowIndex("winBars"), , False
    ShowWindow GetWindowIndex("winMenu"), , False
    ShowWindow GetWindowIndex("winHotbar"), , False
    ShowWindow GetWindowIndex("winChatSmall"), , False
    ' enter loop
    GameLoop
End Sub

Sub HandlePlayerInv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, buffer.ReadLong)
        PlayerInv(i).bound = buffer.ReadByte
    Next
    
    SetGoldLabel

    buffer.Flush: Set buffer = Nothing
End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, buffer.ReadLong) 'CLng(Parse(3)))
    PlayerInv(n).bound = buffer.ReadByte
    buffer.Flush: Set buffer = Nothing
    SetGoldLabel
End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Armor)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Shield)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Pants)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Feet)
    buffer.Flush: Set buffer = Nothing
End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim playerNum As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    playerNum = buffer.ReadLong
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Armor)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Shield)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Pants)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Feet)
    buffer.Flush: Set buffer = Nothing
End Sub

Private Sub HandlePlayerHp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    If MyIndex = 0 Then Exit Sub
    buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.HP) = buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, buffer.ReadLong)
    ' set max width
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 Then
        BarWidth_GuiHP_Max = ((GetPlayerVital(MyIndex, Vitals.HP) / 209) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / 209)) * 209
    Else
        BarWidth_GuiHP_Max = 0
    End If
    ' Update GUI
    UpdateStats_UI
End Sub

Private Sub HandlePlayerMp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.MP) = buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, buffer.ReadLong)
    ' set max width
    If GetPlayerVital(MyIndex, Vitals.MP) > 0 Then
        BarWidth_GuiSP_Max = ((GetPlayerVital(MyIndex, Vitals.MP) / 209) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / 209)) * 209
    Else
        BarWidth_GuiSP_Max = 0
    End If
    ' Update GUI
    UpdateStats_UI
End Sub

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To Stats.Stat_Count - 1
        SetPlayerStat Index, i, buffer.ReadLong
    Next
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    SetPlayerExp MyIndex, buffer.ReadLong
    TNL = buffer.ReadLong
    ' set max width
    If GetPlayerLevel(MyIndex) <= MAX_LEVELS Then
        If GetPlayerExp(MyIndex) > 0 Then
            BarWidth_GuiEXP_Max = ((GetPlayerExp(MyIndex) / 209) / (TNL / 209)) * 209
        Else
            BarWidth_GuiEXP_Max = 0
        End If
    Else
        BarWidth_GuiEXP_Max = 209
    End If
    ' Update GUI
    UpdateStats_UI
End Sub

Private Sub HandlePlayerUpdateDataMission(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, MissionID As Long, X As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    MissionID = buffer.ReadLong
    Player(i).Mission(MissionID).Count = buffer.ReadLong
    
    Call Window_QuestLabelUpdate
End Sub

Private Sub HandlePlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, X As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    Call SetPlayerName(i, buffer.ReadString)
    Call SetPlayerLevel(i, buffer.ReadLong)
    Call SetPlayerPOINTS(i, buffer.ReadLong)
    Call SetPlayerSprite(i, buffer.ReadLong)
    Call SetPlayerMap(i, buffer.ReadLong)
    Call SetPlayerX(i, buffer.ReadLong)
    Call SetPlayerY(i, buffer.ReadLong)
    Call SetPlayerDir(i, buffer.ReadLong)
    Call SetPlayerAccess(i, buffer.ReadLong)
    Call SetPlayerPK(i, buffer.ReadLong)
    Call SetPlayerClass(i, buffer.ReadLong)

    For X = 1 To Stats.Stat_Count - 1
        SetPlayerStat i, X, buffer.ReadLong
    Next
    
    For X = 1 To MAX_PLAYER_MISSIONS
        Player(i).Mission(X).ID = buffer.ReadLong
        Player(i).Mission(X).Count = buffer.ReadLong
    Next X
    
    For X = 1 To MAX_MISSIONS
        Player(i).CompletedMission(X) = buffer.ReadLong
    Next X

    ' Check if the player is the client player
    If i = MyIndex Then
        ' Reset directions
        DirUp = False
        DirLeft = False
        DirDown = False
        DirRight = False
        ' set form
        With Windows(GetWindowIndex("winCharacter"))
            .Controls(GetControlIndex("winCharacter", "lblName")).text = "Name: " & Trim$(GetPlayerName(MyIndex))
            .Controls(GetControlIndex("winCharacter", "lblClass")).text = "Class: " & Trim$(Class(GetPlayerClass(MyIndex)).Name)
            .Controls(GetControlIndex("winCharacter", "lblLevel")).text = "Level: " & GetPlayerLevel(MyIndex)
            .Controls(GetControlIndex("winCharacter", "lblGuild")).text = "Guild: " & "None"
            .Controls(GetControlIndex("winCharacter", "lblHealth")).text = "Health: " & GetPlayerVital(MyIndex, HP) & "/" & GetPlayerMaxVital(MyIndex, HP)
            .Controls(GetControlIndex("winCharacter", "lblSpirit")).text = "Spirit: " & GetPlayerVital(MyIndex, MP) & "/" & GetPlayerMaxVital(MyIndex, MP)
            .Controls(GetControlIndex("winCharacter", "lblExperience")).text = "Experience: " & Player(MyIndex).EXP & "/" & TNL
            ' stats
            For X = 1 To Stats.Stat_Count - 1
                .Controls(GetControlIndex("winCharacter", "lblStat_" & X)).text = GetPlayerStat(MyIndex, X)
            Next
            ' points
            .Controls(GetControlIndex("winCharacter", "lblPoints")).text = GetPlayerPOINTS(MyIndex)
            ' grey out buttons
            If GetPlayerPOINTS(MyIndex) = 0 Then
                For X = 1 To Stats.Stat_Count - 1
                    .Controls(GetControlIndex("winCharacter", "btnGreyStat_" & X)).visible = True
                Next
            Else
                For X = 1 To Stats.Stat_Count - 1
                    .Controls(GetControlIndex("winCharacter", "btnGreyStat_" & X)).visible = False
                Next
            End If
        End With
        
        ' Mission
        Call Window_QuestButtonUpdate
    End If

    ' Make sure they aren't walking
    Player(i).Moving = 0
    Player(i).xOffset = 0
    Player(i).yOffset = 0
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim Dir As Long
    Dim n As Byte
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    n = buffer.ReadLong
    Call SetPlayerX(i, X)
    Call SetPlayerY(i, Y)
    Call SetPlayerDir(i, Dir)
    Player(i).xOffset = 0
    Player(i).yOffset = 0
    Player(i).Moving = n

    Select Case GetPlayerDir(i)

        Case DIR_UP
            Player(i).yOffset = PIC_Y

        Case DIR_DOWN
            Player(i).yOffset = PIC_Y * -1

        Case DIR_LEFT
            Player(i).xOffset = PIC_X

        Case DIR_RIGHT
            Player(i).xOffset = PIC_X * -1
        
        Case DIR_UP_LEFT
            Player(i).yOffset = PIC_Y
            Player(i).xOffset = PIC_X
            
        Case DIR_UP_RIGHT
            Player(i).yOffset = PIC_Y
            Player(i).xOffset = PIC_X * -1

        Case DIR_DOWN_LEFT
            Player(i).yOffset = PIC_Y * -1
            Player(i).xOffset = PIC_X

        Case DIR_DOWN_RIGHT
            Player(i).yOffset = PIC_Y * -1
            Player(i).xOffset = PIC_X * -1
    End Select
End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Dir As Byte
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerDir(i, Dir)

    With Player(i)
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim Dir As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerX(MyIndex, X)
    Call SetPlayerY(MyIndex, Y)
    Call SetPlayerDir(MyIndex, Dir)
    ' Make sure they aren't walking
    Player(MyIndex).Moving = 0
    Player(MyIndex).xOffset = 0
    Player(MyIndex).yOffset = 0
End Sub

Private Sub HandlePlayerXYMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim Dir As Long
    Dim buffer As clsBuffer
    Dim thePlayer As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    thePlayer = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerX(thePlayer, X)
    Call SetPlayerY(thePlayer, Y)
    Call SetPlayerDir(thePlayer, Dir)
    ' Make sure they aren't walking
    Player(thePlayer).Moving = 0
    Player(thePlayer).xOffset = 0
    Player(thePlayer).yOffset = 0
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    ' Set player to attacking
    Player(i).Attacking = 1
    Player(i).AttackTimer = getTime
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, NeedMap As Byte, buffer As clsBuffer, MapDataCRC As Long, MapTileCRC As Long, mapNum As Long
    
    GettingMap = True
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Erase all players except self
    For i = 1 To Player_HighIndex
        If i <> MyIndex Then
            Call SetPlayerMap(i, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearTempTile
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap

    ' clear the blood
    For i = 1 To MAX_BYTE
        Blood(i).X = 0
        Blood(i).Y = 0
        Blood(i).sprite = 0
        Blood(i).timer = 0
    Next

    ' Get map num
    mapNum = buffer.ReadLong
    MapDataCRC = buffer.ReadLong
    MapTileCRC = buffer.ReadLong
    
    ' check against our own CRC32s
    NeedMap = 0
    If MapDataCRC <> MapCRC32(mapNum).MapDataCRC Then
        NeedMap = 1
    End If
    If MapTileCRC <> MapCRC32(mapNum).MapTileCRC Then
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set buffer = New clsBuffer
    buffer.WriteLong CNeedMap
    buffer.WriteLong NeedMap
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If Not applyingMap Then
        If InMapEditor Then
            InMapEditor = False
            frmEditor_Map.visible = False
            ClearAttributeDialogue
    
            If frmEditor_MapProperties.visible Then
                frmEditor_MapProperties.visible = False
            End If
        End If
    End If
    
    ' load the map if we don't need it
    If NeedMap = 0 Then
        LoadMap mapNum
        applyingMap = False
        CacheNewMapSounds
    End If
End Sub

Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, mapNum As Long, i As Long, X As Long, Y As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    'zlib
    buffer.DecompressBuffer
    
    mapNum = buffer.ReadLong
    
    With Map.MapData
        .Name = buffer.ReadString
        .Music = buffer.ReadString
        .Moral = buffer.ReadByte
        .Up = buffer.ReadLong
        .Down = buffer.ReadLong
        .Left = buffer.ReadLong
        .Right = buffer.ReadLong
        .BootMap = buffer.ReadLong
        .BootX = buffer.ReadByte
        .BootY = buffer.ReadByte
        .MaxX = buffer.ReadByte
        .MaxY = buffer.ReadByte
        
        .Weather = buffer.ReadLong
        .WeatherIntensity = buffer.ReadLong
        
        .Fog = buffer.ReadLong
        .FogSpeed = buffer.ReadLong
        .FogOpacity = buffer.ReadLong
        
        .Red = buffer.ReadLong
        .Green = buffer.ReadLong
        .Blue = buffer.ReadLong
        .alpha = buffer.ReadLong
        
        .BossNpc = buffer.ReadLong
        For i = 1 To MAX_MAP_NPCS
            .Npc(i) = buffer.ReadLong
        Next
    End With
    
    ReDim Map.TileData.Tile(0 To Map.MapData.MaxX, 0 To Map.MapData.MaxY)

    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map.TileData.Tile(X, Y).Layer(i).X = buffer.ReadLong
                Map.TileData.Tile(X, Y).Layer(i).Y = buffer.ReadLong
                Map.TileData.Tile(X, Y).Layer(i).tileSet = buffer.ReadLong
                Map.TileData.Tile(X, Y).Autotile(i) = buffer.ReadByte
            Next
            Map.TileData.Tile(X, Y).Type = buffer.ReadByte
            Map.TileData.Tile(X, Y).Data1 = buffer.ReadLong
            Map.TileData.Tile(X, Y).Data2 = buffer.ReadLong
            Map.TileData.Tile(X, Y).Data3 = buffer.ReadLong
            Map.TileData.Tile(X, Y).Data4 = buffer.ReadLong
            Map.TileData.Tile(X, Y).Data5 = buffer.ReadLong
            Map.TileData.Tile(X, Y).DirBlock = buffer.ReadByte
        Next
    Next

    ClearTempTile
    initAutotiles
    CacheNewMapSounds
    buffer.Flush: Set buffer = Nothing
    ' Save the map
    Call SaveMap(mapNum)
    GetMapCRC32 mapNum
    AddText "Downloaded new map.", BrightGreen

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If Not applyingMap Then
        If InMapEditor Then
            InMapEditor = False
            frmEditor_Map.visible = False
            ClearAttributeDialogue
            If frmEditor_MapProperties.visible Then
                frmEditor_MapProperties.visible = False
            End If
        End If
    End If
    applyingMap = False

End Sub

Private Sub HandleMapItemData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer, tmpLong As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_ITEMS

        With MapItem(i)
            .playerName = buffer.ReadString
            .Num = buffer.ReadLong
            .Value = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
            tmpLong = buffer.ReadLong

            If tmpLong = 0 Then
                .bound = False
            Else
                .bound = True
            End If

        End With

    Next

End Sub

Private Sub HandleMapDone()
    Dim i As Long
    Dim musicFile As String

    ' clear the action msgs
    For i = 1 To MAX_BYTE
        ClearActionMsg (i)
    Next i

    Action_HighIndex = 1

    ' player music
    If InGame Then
        musicFile = Trim$(Map.MapData.Music)

        If Not musicFile = "None." Then
            Play_Music musicFile
        Else
            Stop_Music
        End If
    End If

    ' get the npc high index
    For i = MAX_MAP_NPCS To 1 Step -1

        If MapNpc(i).Num > 0 Then
            Npc_HighIndex = i + 1
            Exit For
        End If

    Next

    ' make sure we're not overflowing
    If Npc_HighIndex > MAX_MAP_NPCS Then Npc_HighIndex = MAX_MAP_NPCS
    ' now cache the positions
    initAutotiles
    CurrentWeather = Map.MapData.Weather
    CurrentWeatherIntensity = Map.MapData.WeatherIntensity
    CurrentFog = Map.MapData.Fog
    CurrentFogSpeed = Map.MapData.FogSpeed
    CurrentFogOpacity = Map.MapData.FogOpacity
    CurrentTintR = Map.MapData.Red
    CurrentTintG = Map.MapData.Green
    CurrentTintB = Map.MapData.Blue
    CurrentTintA = Map.MapData.alpha
    GettingMap = False
    CanMoveNow = True
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer, tmpLong As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong

    With MapItem(n)
        .playerName = buffer.ReadString
        .Num = buffer.ReadLong
        .Value = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        tmpLong = buffer.ReadLong

        If tmpLong = 0 Then
            .bound = False
        Else
            .bound = True
        End If

    End With

End Sub

Private Sub HandleItemEditor()
    Dim i As Long

    With frmEditor_Item
        Editor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With

End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    buffer.Flush: Set buffer = Nothing
End Sub

Private Sub HandleMapKey(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim X As Long
    Dim Y As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    n = buffer.ReadByte
    TempTile(X, Y).DoorOpen = n

    ' re-cache rendering
    If Not GettingMap Then cacheRenderState X, Y, MapLayer.Mask
End Sub

Private Sub HandleEditMap()
    Call MapEditorInit
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Call ClearPlayer(buffer.ReadLong)
    buffer.Flush: Set buffer = Nothing
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PingEnd = getTime
    Ping = PingEnd - PingStart
End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Long, Y As Long, message As String, Color As Long, tmpType As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    message = buffer.ReadString
    Color = buffer.ReadLong
    tmpType = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    buffer.Flush: Set buffer = Nothing
    CreateActionMsg message, Color, tmpType, X, Y
End Sub

Private Sub HandleBlood(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Long, Y As Long, sprite As Long, i As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    buffer.Flush: Set buffer = Nothing
    ' randomise sprite
    sprite = Rand(1, BloodCount)

    ' make sure tile doesn't already have blood
    For i = 1 To MAX_BYTE

        If Blood(i).X = X And Blood(i).Y = Y Then
            ' already have blood :(
            Exit Sub
        End If

    Next

    ' carry on with the set
    BloodIndex = BloodIndex + 1

    If BloodIndex >= MAX_BYTE Then BloodIndex = 1

    With Blood(BloodIndex)
        .X = X
        .Y = Y
        .sprite = sprite
        .timer = getTime
    End With

End Sub

Private Sub HandleCooldown(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Slot As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Slot = buffer.ReadLong
    SpellCD(Slot) = getTime
    buffer.Flush: Set buffer = Nothing
End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SpellBuffer = 0
    SpellBufferTimer = 0
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, Access As Long, Name As String, message As String, Colour As Long, header As String, PK As Long, saycolour As Long
    Dim Channel As Byte, colStr As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Name = buffer.ReadString
    Access = buffer.ReadLong
    PK = buffer.ReadLong
    message = buffer.ReadString
    header = buffer.ReadString
    saycolour = buffer.ReadLong
    buffer.Flush: Set buffer = Nothing
    
    ' Check access level
    Colour = White

    If Access > 0 Then Colour = Pink
    If PK > 0 Then Colour = BrightRed
    
    ' find channel
    Channel = 0
    Select Case header
        Case "[Map] "
            Channel = ChatChannel.chMap
        Case "[Global] "
            Channel = ChatChannel.chGlobal
    End Select
    
    ' remove the colour char from the message
    message = Replace$(message, ColourChar, vbNullString)
    ' add to the chat box
    AddText ColourChar & GetColStr(Colour) & header & Name & ": " & ColourChar & GetColStr(Grey) & message, Grey, , Channel
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim shopNum As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    shopNum = buffer.ReadLong
    OpenShop shopNum
    buffer.Flush: Set buffer = Nothing
End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    StunDuration = buffer.ReadLong
    buffer.Flush: Set buffer = Nothing
End Sub

Private Sub HandleBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_BANK
        Bank.Item(i).Num = buffer.ReadLong
        Bank.Item(i).Value = buffer.ReadLong
    Next

    InBank = True
    buffer.Flush: Set buffer = Nothing
    
    If Not Windows(GetWindowIndex("winBank")).Window.visible Then
        ShowWindow GetWindowIndex("winBank"), , False
    End If
End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InTrade = buffer.ReadLong
    buffer.Flush: Set buffer = Nothing
    
    ShowTrade
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    InTrade = 0
    HideWindow GetWindowIndex("winTrade")
End Sub

Private Sub HandleTradeUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, dataType As Byte, i As Long, yourWorth As Long, theirWorth As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    dataType = buffer.ReadByte

    If dataType = 0 Then ' ours!
        For i = 1 To MAX_INV
            TradeYourOffer(i).Num = buffer.ReadLong
            TradeYourOffer(i).Value = buffer.ReadLong
        Next
        yourWorth = buffer.ReadLong
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblYourValue")).text = yourWorth & "g"
    ElseIf dataType = 1 Then 'theirs
        For i = 1 To MAX_INV
            TradeTheirOffer(i).Num = buffer.ReadLong
            TradeTheirOffer(i).Value = buffer.ReadLong
        Next
        theirWorth = buffer.ReadLong
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblTheirValue")).text = theirWorth & "g"
    End If

    buffer.Flush: Set buffer = Nothing
End Sub

Private Sub HandleTradeStatus(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim tradeStatus As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    tradeStatus = buffer.ReadByte
    buffer.Flush: Set buffer = Nothing

    Select Case tradeStatus
        Case 0 ' clear
            Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Choose items to offer."
        Case 1 ' they've accepted
            Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Other player has accepted."
        Case 2 ' you've accepted
            Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Waiting for other player to accept."
        Case 3 ' no room
            Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Not enough inventory space."
    End Select
End Sub

Private Sub HandleTarget(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    myTarget = buffer.ReadLong
    myTargetType = buffer.ReadLong
    buffer.Flush: Set buffer = Nothing
End Sub

Private Sub HandleHotbar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_HOTBAR
        Hotbar(i).Slot = buffer.ReadLong
        Hotbar(i).sType = buffer.ReadByte
    Next
End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Player_HighIndex = buffer.ReadLong
End Sub

Private Sub HandleResetShopAction(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    UpdateShop
End Sub

Private Sub HandleSound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Long, Y As Long, entityType As Long, entityNum As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    entityType = buffer.ReadLong
    entityNum = buffer.ReadLong
    PlayMapSound X, Y, entityType, entityNum
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Index_Offer As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Index_Offer = FindOpenOfferSlot
    
    If Index_Offer <> 0 Then
        inOfferInvite(Index_Offer) = buffer.ReadString
        inOfferType(Index_Offer) = Offers.Offer_Type_Trade
    End If
    buffer.Flush: Set buffer = Nothing
    
    Call UpdateWindowOffer(Index_Offer)
'    Dim buffer As clsBuffer, theName As String, Top As Long
'
'    Set buffer = New clsBuffer
'    buffer.WriteBytes Data()
'    theName = buffer.ReadString
'    ' cache name and show invitation
'    diaDataString = theName
'    ShowWindow GetWindowIndex("winInvite_Trade")
'    Windows(GetWindowIndex("winInvite_Trade")).Controls(GetControlIndex("winInvite_Trade", "btnInvite")).text = ColourChar & White & theName & ColourChar & "-1" & " has invited you to trade."
'    AddText Trim$(theName) & " has invited you to trade.", White
'    ' loop through
'    Top = ScreenHeight - 80
'    If Windows(GetWindowIndex("winInvite_Party")).Window.visible Then
'        Top = Top - 37
'    End If
'    Windows(GetWindowIndex("winInvite_Trade")).Window.Top = Top
End Sub

Private Sub HandlePartyInvite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, Top As Long
    Dim Index_Offer As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Index_Offer = FindOpenOfferSlot
    
    If Index_Offer <> 0 Then
        inOfferInvite(Index_Offer) = buffer.ReadString
        inOfferType(Index_Offer) = Offers.Offer_Type_Party
    End If
    buffer.Flush: Set buffer = Nothing
    
    Call UpdateWindowOffer(Index_Offer)
End Sub

Private Sub HandlePartyUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, i As Long, inParty As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    inParty = buffer.ReadByte

    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        UpdatePartyInterface
        ' exit out early
        Exit Sub
    End If

    ' carry on otherwise
    Party.Leader = buffer.ReadLong

    For i = 1 To MAX_PARTY_MEMBERS
        Party.Member(i) = buffer.ReadLong
    Next

    Party.MemberCount = buffer.ReadLong
    
    ' update the party interface
    UpdatePartyInterface
End Sub

Private Sub HandlePartyVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim playerNum As Long
    Dim buffer As clsBuffer, i As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' which player?
    playerNum = buffer.ReadLong

    ' set vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(i) = buffer.ReadLong
        Player(playerNum).Vital(i) = buffer.ReadLong
    Next

    ' update the party interface
    UpdatePartyBars
End Sub

Private Sub HandleChatUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, NpcNum As Long, mT As String, o(1 To 4) As String, i As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    
    NpcNum = buffer.ReadLong
    mT = buffer.ReadString
    For i = 1 To 4
        o(i) = buffer.ReadString
    Next

    buffer.Flush: Set buffer = Nothing

    ' if npcNum is 0, exit the chat system
    If NpcNum = 0 Then
        inChat = False
        HideWindow GetWindowIndex("winNpcChat")
        Exit Sub
    End If

    ' set chat going
    OpenNpcChat NpcNum, mT, o
End Sub

Private Sub HandleStartTutorial(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    'inTutorial = True
    ' set the first message
    'SetTutorialState 1
End Sub

Private Sub HandleChatBubble(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, TargetType As Long, target As Long, message As String, Colour As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    target = buffer.ReadLong
    TargetType = buffer.ReadLong
    message = buffer.ReadString
    Colour = buffer.ReadLong
    AddChatBubble target, TargetType, message, Colour
    buffer.Flush: Set buffer = Nothing
End Sub

Private Sub HandlePlayerChars(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, i As Long, winNum As Long, conNum As Long, isSlotEmpty(1 To MAX_CHARS) As Boolean, X As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()

    For i = 1 To MAX_CHARS
        CharName(i) = Trim$(buffer.ReadString)
        CharSprite(i) = buffer.ReadLong
        CharAccess(i) = buffer.ReadLong
        CharClass(i) = buffer.ReadLong
        ' set as empty or not
        If Not Len(Trim$(CharName(i))) > 0 Then isSlotEmpty(i) = True
    Next

    buffer.Flush: Set buffer = Nothing
    
    HideWindows
    ShowWindow GetWindowIndex("winCharacters")
    
    ' set GUI window up
    winNum = GetWindowIndex("winCharacters")
    For i = 1 To MAX_CHARS
        conNum = GetControlIndex("winCharacters", "lblCharName_" & i)
        With Windows(winNum).Controls(conNum)
            If Not isSlotEmpty(i) Then
                .text = CharName(i)
            Else
                .text = "Blank Slot"
            End If
        End With
        ' hide/show buttons
        If isSlotEmpty(i) Then
            ' create button
            conNum = GetControlIndex("winCharacters", "btnCreateChar_" & i)
            Windows(winNum).Controls(conNum).visible = True
            ' select button
            conNum = GetControlIndex("winCharacters", "btnSelectChar_" & i)
            Windows(winNum).Controls(conNum).visible = False
            ' delete button
            conNum = GetControlIndex("winCharacters", "btnDelChar_" & i)
            Windows(winNum).Controls(conNum).visible = False
        Else
            ' create button
            conNum = GetControlIndex("winCharacters", "btnCreateChar_" & i)
            Windows(winNum).Controls(conNum).visible = False
            ' select button
            conNum = GetControlIndex("winCharacters", "btnSelectChar_" & i)
            Windows(winNum).Controls(conNum).visible = True
            ' delete button
            conNum = GetControlIndex("winCharacters", "btnDelChar_" & i)
            Windows(winNum).Controls(conNum).visible = True
        End If
    Next
End Sub

Private Sub HandlePlayerVariables(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For i = 1 To MAX_BYTE
        Player(MyIndex).Variable(i) = buffer.ReadLong
    Next
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub HandleProjectile(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ProjectileSlot As Long, i As Long, Angle As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer

    ' read bytes from data()
    buffer.WriteBytes Data()

    ' recieve projectile data
    ProjectileSlot = buffer.ReadLong
    LastProjectile = buffer.ReadLong
    
    ReDim Preserve MapProjectile(1 To LastProjectile)
        
    ' populate the values
    With MapProjectile(ProjectileSlot)
        .Owner = buffer.ReadLong
        .OwnerType = buffer.ReadLong
        .Direction = buffer.ReadLong
        .Graphic = buffer.ReadLong
        .IsAoE = buffer.ReadByte
        .Rotate = buffer.ReadLong
        .RotateSpeed = buffer.ReadLong
        .speed = buffer.ReadLong
        .Duration = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .tx = buffer.ReadLong
        .ty = buffer.ReadLong
        .IsDirectional = buffer.ReadByte
        
        If .speed >= 5000 Then
            .Duration = Tick + .Duration
        End If
        
        For i = 1 To 4
            .ProjectileOffset(i).X = buffer.ReadLong
            .ProjectileOffset(i).Y = buffer.ReadLong
        Next
        
    End With

End Sub

Sub HandleClearProjectile(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ProjectileSlot As Long, i As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ProjectileSlot = buffer.ReadLong
    ClearProjectile ProjectileSlot
    
    buffer.Flush
    Set buffer = Nothing
End Sub
