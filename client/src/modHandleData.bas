Attribute VB_Name = "modHandleData"
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
    HandleDataSub(SEvent) = GetAddress(AddressOf HandleEvent)
End Sub

Sub HandleData(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim MsgType As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MsgType = Buffer.ReadLong

    If MsgType < 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMsgCOUNT Then
        DestroyGame
        Exit Sub
    End If

    CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.length), 0, 0
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, dialogue_index As Long, menuReset As Long, kick As Long
    
    SetStatus vbNullString
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    dialogue_index = Buffer.ReadLong
    menuReset = Buffer.ReadLong
    kick = Buffer.ReadLong
    
    Buffer.Flush: Set Buffer = Nothing
    
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
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Now we can receive game data
    MyIndex = Buffer.ReadLong
    ' player high index
    Player_HighIndex = MAX_PLAYERS 'Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    Call SetStatus("Receiving game data.")
End Sub

Sub HandleNewCharClasses(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim I As Long
    Dim z As Long, X As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For I = 1 To Max_Classes

        With Class(I)
            .name = Buffer.ReadString
            .Vital(Vitals.HP) = Buffer.ReadLong
            .Vital(Vitals.MP) = Buffer.ReadLong
            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = Buffer.ReadLong
            Next

            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = Buffer.ReadLong
            Next

            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = Buffer.ReadLong
            Next

        End With

        n = n + 10
    Next

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleClassesData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim I As Long
    Dim z As Long, X As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong 'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For I = 1 To Max_Classes

        With Class(I)
            .name = Buffer.ReadString 'Trim$(Parse(n))
            .Vital(Vitals.HP) = Buffer.ReadLong 'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = Buffer.ReadLong 'CLng(Parse(n + 2))
            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = Buffer.ReadLong
            Next

            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = Buffer.ReadLong
            Next

            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = Buffer.ReadLong
            Next

        End With

        n = n + 10
    Next

    Buffer.Flush: Set Buffer = Nothing
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
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For I = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, I, Buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, I, Buffer.ReadLong)
        PlayerInv(I).bound = Buffer.ReadByte
    Next
    
    SetGoldLabel

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(3)))
    PlayerInv(n).bound = Buffer.ReadByte
    Buffer.Flush: Set Buffer = Nothing
    SetGoldLabel
End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Shield)
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim playerNum As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    playerNum = Buffer.ReadLong
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Shield)
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandlePlayerHp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    If MyIndex = 0 Then Exit Sub
    Buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.HP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, Buffer.ReadLong)
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
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.MP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, Buffer.ReadLong)
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
    Dim Buffer As clsBuffer
    Dim I As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For I = 1 To Stats.Stat_Count - 1
        SetPlayerStat Index, I, Buffer.ReadLong
    Next
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    SetPlayerExp MyIndex, Buffer.ReadLong
    TNL = Buffer.ReadLong
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

Private Sub HandlePlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long, X As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    I = Buffer.ReadLong
    Call SetPlayerName(I, Buffer.ReadString)
    Call SetPlayerLevel(I, Buffer.ReadLong)
    Call SetPlayerPOINTS(I, Buffer.ReadLong)
    Call SetPlayerSprite(I, Buffer.ReadLong)
    Call SetPlayerMap(I, Buffer.ReadLong)
    Call SetPlayerX(I, Buffer.ReadLong)
    Call SetPlayerY(I, Buffer.ReadLong)
    Call SetPlayerDir(I, Buffer.ReadLong)
    Call SetPlayerAccess(I, Buffer.ReadLong)
    Call SetPlayerPK(I, Buffer.ReadLong)
    Call SetPlayerClass(I, Buffer.ReadLong)

    For X = 1 To Stats.Stat_Count - 1
        SetPlayerStat I, X, Buffer.ReadLong
    Next

    ' Check if the player is the client player
    If I = MyIndex Then
        ' Reset directions
        DirUp = False
        DirLeft = False
        DirDown = False
        DirRight = False
        ' set form
        With Windows(GetWindowIndex("winCharacter"))
            .Controls(GetControlIndex("winCharacter", "lblName")).text = "Name: " & Trim$(GetPlayerName(MyIndex))
            .Controls(GetControlIndex("winCharacter", "lblClass")).text = "Class: " & Trim$(Class(GetPlayerClass(MyIndex)).name)
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
    End If

    ' Make sure they aren't walking
    Player(I).Moving = 0
    Player(I).xOffset = 0
    Player(I).yOffset = 0
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long
    Dim X As Long
    Dim Y As Long
    Dim Dir As Long
    Dim n As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    I = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    n = Buffer.ReadLong
    Call SetPlayerX(I, X)
    Call SetPlayerY(I, Y)
    Call SetPlayerDir(I, Dir)
    Player(I).xOffset = 0
    Player(I).yOffset = 0
    Player(I).Moving = n

    Select Case GetPlayerDir(I)

        Case DIR_UP
            Player(I).yOffset = PIC_Y

        Case DIR_DOWN
            Player(I).yOffset = PIC_Y * -1

        Case DIR_LEFT
            Player(I).xOffset = PIC_X

        Case DIR_RIGHT
            Player(I).xOffset = PIC_X * -1
        
        Case DIR_UP_LEFT
            Player(I).yOffset = PIC_Y
            Player(I).xOffset = PIC_X
            
        Case DIR_UP_RIGHT
            Player(I).yOffset = PIC_Y
            Player(I).xOffset = PIC_X * -1

        Case DIR_DOWN_LEFT
            Player(I).yOffset = PIC_Y * -1
            Player(I).xOffset = PIC_X

        Case DIR_DOWN_RIGHT
            Player(I).yOffset = PIC_Y * -1
            Player(I).xOffset = PIC_X * -1
    End Select
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim MapNpcNum As Long
    Dim X As Long
    Dim Y As Long
    Dim Dir As Long
    Dim Movement As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MapNpcNum = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Movement = Buffer.ReadLong

    With MapNpc(MapNpcNum)
        .X = X
        .Y = Y
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = Movement

        Select Case .Dir

            Case DIR_UP
                .yOffset = PIC_Y

            Case DIR_DOWN
                .yOffset = PIC_Y * -1

            Case DIR_LEFT
                .xOffset = PIC_X

            Case DIR_RIGHT
                .xOffset = PIC_X * -1
            
            Case DIR_UP_LEFT
                .yOffset = PIC_Y
                .xOffset = PIC_X

            Case DIR_UP_RIGHT
                .yOffset = PIC_Y
                .xOffset = PIC_X * -1
                
            Case DIR_DOWN_LEFT
                .yOffset = PIC_Y * -1
                .xOffset = PIC_X
                
            Case DIR_DOWN_RIGHT
                .yOffset = PIC_Y * -1
                .xOffset = PIC_X * -1
        End Select

    End With

End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long
    Dim Dir As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    I = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Call SetPlayerDir(I, Dir)

    With Player(I)
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long
    Dim Dir As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    I = Buffer.ReadLong
    Dir = Buffer.ReadLong

    With MapNpc(I)
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim Dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
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
    Dim Buffer As clsBuffer
    Dim thePlayer As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    thePlayer = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Call SetPlayerX(thePlayer, X)
    Call SetPlayerY(thePlayer, Y)
    Call SetPlayerDir(thePlayer, Dir)
    ' Make sure they aren't walking
    Player(thePlayer).Moving = 0
    Player(thePlayer).xOffset = 0
    Player(thePlayer).yOffset = 0
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    I = Buffer.ReadLong
    ' Set player to attacking
    Player(I).Attacking = 1
    Player(I).AttackTimer = GetTickCount
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    I = Buffer.ReadLong
    ' Set player to attacking
    MapNpc(I).Attacking = 1
    MapNpc(I).AttackTimer = GetTickCount
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long, NeedMap As Byte, Buffer As clsBuffer, MapDataCRC As Long, MapTileCRC As Long, mapNum As Long
    
    GettingMap = True
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Erase all players except self
    For I = 1 To MAX_PLAYERS
        If I <> MyIndex Then
            Call SetPlayerMap(I, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearTempTile
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap

    ' clear the blood
    For I = 1 To MAX_BYTE
        Blood(I).X = 0
        Blood(I).Y = 0
        Blood(I).sprite = 0
        Blood(I).timer = 0
    Next

    ' Get map num
    mapNum = Buffer.ReadLong
    MapDataCRC = Buffer.ReadLong
    MapTileCRC = Buffer.ReadLong
    
    ' check against our own CRC32s
    NeedMap = 0
    If MapDataCRC <> MapCRC32(mapNum).MapDataCRC Then
        NeedMap = 1
    End If
    If MapTileCRC <> MapCRC32(mapNum).MapTileCRC Then
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteLong NeedMap
    SendData Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing

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
    Dim Buffer As clsBuffer, mapNum As Long, I As Long, X As Long, Y As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    mapNum = Buffer.ReadLong
    
    With Map.MapData
        .name = Buffer.ReadString
        .Music = Buffer.ReadString
        .Moral = Buffer.ReadByte
        .Up = Buffer.ReadLong
        .Down = Buffer.ReadLong
        .Left = Buffer.ReadLong
        .Right = Buffer.ReadLong
        .BootMap = Buffer.ReadLong
        .BootX = Buffer.ReadByte
        .BootY = Buffer.ReadByte
        .MaxX = Buffer.ReadByte
        .MaxY = Buffer.ReadByte
        
        .Weather = Buffer.ReadLong
        .WeatherIntensity = Buffer.ReadLong
        
        .Fog = Buffer.ReadLong
        .FogSpeed = Buffer.ReadLong
        .FogOpacity = Buffer.ReadLong
        
        .Red = Buffer.ReadLong
        .Green = Buffer.ReadLong
        .Blue = Buffer.ReadLong
        .alpha = Buffer.ReadLong
        
        .BossNpc = Buffer.ReadLong
        For I = 1 To MAX_MAP_NPCS
            .Npc(I) = Buffer.ReadLong
        Next
    End With
    
    ReDim Map.TileData.Tile(0 To Map.MapData.MaxX, 0 To Map.MapData.MaxY)

    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            For I = 1 To MapLayer.Layer_Count - 1
                Map.TileData.Tile(X, Y).Layer(I).X = Buffer.ReadLong
                Map.TileData.Tile(X, Y).Layer(I).Y = Buffer.ReadLong
                Map.TileData.Tile(X, Y).Layer(I).tileSet = Buffer.ReadLong
                Map.TileData.Tile(X, Y).Autotile(I) = Buffer.ReadByte
            Next
            Map.TileData.Tile(X, Y).Type = Buffer.ReadByte
            Map.TileData.Tile(X, Y).Data1 = Buffer.ReadLong
            Map.TileData.Tile(X, Y).Data2 = Buffer.ReadLong
            Map.TileData.Tile(X, Y).Data3 = Buffer.ReadLong
            Map.TileData.Tile(X, Y).Data4 = Buffer.ReadLong
            Map.TileData.Tile(X, Y).Data5 = Buffer.ReadLong
            Map.TileData.Tile(X, Y).DirBlock = Buffer.ReadByte
        Next
    Next

    ClearTempTile
    initAutotiles
    CacheNewMapSounds
    Buffer.Flush: Set Buffer = Nothing
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
    Dim I As Long
    Dim Buffer As clsBuffer, tmpLong As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For I = 1 To MAX_MAP_ITEMS

        With MapItem(I)
            .playerName = Buffer.ReadString
            .num = Buffer.ReadLong
            .value = Buffer.ReadLong
            .X = Buffer.ReadLong
            .Y = Buffer.ReadLong
            tmpLong = Buffer.ReadLong

            If tmpLong = 0 Then
                .bound = False
            Else
                .bound = True
            End If

        End With

    Next

End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For I = 1 To MAX_MAP_NPCS

        With MapNpc(I)
            .num = Buffer.ReadLong
            .X = Buffer.ReadLong
            .Y = Buffer.ReadLong
            .Dir = Buffer.ReadLong
            .Vital(HP) = Buffer.ReadLong
        End With

    Next

End Sub

Private Sub HandleMapDone()
    Dim I As Long
    Dim musicFile As String

    ' clear the action msgs
    For I = 1 To MAX_BYTE
        ClearActionMsg (I)
    Next I

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
    For I = MAX_MAP_NPCS To 1 Step -1

        If MapNpc(I).num > 0 Then
            Npc_HighIndex = I + 1
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
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer, tmpLong As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong

    With MapItem(n)
        .playerName = Buffer.ReadString
        .num = Buffer.ReadLong
        .value = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        tmpLong = Buffer.ReadLong

        If tmpLong = 0 Then
            .bound = False
        Else
            .bound = True
        End If

    End With

End Sub

Private Sub HandleItemEditor()
    Dim I As Long

    With frmEditor_Item
        Editor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_ITEMS
            .lstIndex.AddItem I & ": " & Trim$(Item(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With

End Sub

Private Sub HandleAnimationEditor()
    Dim I As Long

    With frmEditor_Animation
        Editor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem I & ": " & Trim$(Animation(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With

End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = Buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong

    With MapNpc(n)
        .num = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .Dir = Buffer.ReadLong
        ' Client use only
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

End Sub

Private Sub HandleNpcDead(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    Call ClearMapNpc(n)
End Sub

Private Sub HandleNpcEditor()
    Dim I As Long

    With frmEditor_NPC
        Editor = EDITOR_NPC
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_NPCS
            .lstIndex.AddItem I & ": " & Trim$(Npc(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With

End Sub

Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    NpcSize = LenB(Npc(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = Buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(Npc(n)), ByVal VarPtr(NpcData(0)), NpcSize
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleResourceEditor()
    Dim I As Long

    With frmEditor_Resource
        Editor = EDITOR_RESOURCE
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_RESOURCES
            .lstIndex.AddItem I & ": " & Trim$(Resource(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ResourceEditorInit
    End With

End Sub

Private Sub HandleUpdateResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ResourceNum = Buffer.ReadLong
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = Buffer.ReadBytes(ResourceSize)
    ClearResource ResourceNum
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleMapKey(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim X As Long
    Dim Y As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    n = Buffer.ReadByte
    TempTile(X, Y).DoorOpen = n

    ' re-cache rendering
    If Not GettingMap Then cacheRenderState X, Y, MapLayer.Mask
End Sub

Private Sub HandleEditMap()
    Call MapEditorInit
End Sub

Private Sub HandleShopEditor()
    Dim I As Long

    With frmEditor_Shop
        Editor = EDITOR_SHOP
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_SHOPS
            .lstIndex.AddItem I & ": " & Trim$(Shop(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ShopEditorInit
    End With

End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim shopNum As Long
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    shopNum = Buffer.ReadLong
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopNum)), ByVal VarPtr(ShopData(0)), ShopSize
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleSpellEditor()
    Dim I As Long

    With frmEditor_Spell
        Editor = EDITOR_SPELL
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_SPELLS
            .lstIndex.AddItem I & ": " & Trim$(Spell(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SpellEditorInit
    End With

End Sub

Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim spellnum As Long
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    spellnum = Buffer.ReadLong
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For I = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(I).Spell = Buffer.ReadLong
        PlayerSpells(I).Uses = Buffer.ReadLong
    Next

    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call ClearPlayer(Buffer.ReadLong)
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleResourceCache(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim I As Long

    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Resource_Index = Buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        For I = 0 To Resource_Index
            MapResource(I).ResourceState = Buffer.ReadByte
            MapResource(I).X = Buffer.ReadLong
            MapResource(I).Y = Buffer.ReadLong
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PingEnd = GetTickCount
    Ping = PingEnd - PingStart
End Sub

Private Sub HandleDoorAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim X As Long, Y As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong

    With TempTile(X, Y)
        .DoorFrame = 1
        .DoorAnimate = 1 ' 0 = nothing| 1 = opening | 2 = closing
        .DoorTimer = GetTickCount
    End With

    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim X As Long, Y As Long, message As String, Color As Long, tmpType As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    message = Buffer.ReadString
    Color = Buffer.ReadLong
    tmpType = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    CreateActionMsg message, Color, tmpType, X, Y
End Sub

Private Sub HandleBlood(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim X As Long, Y As Long, sprite As Long, I As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    ' randomise sprite
    sprite = Rand(1, BloodCount)

    ' make sure tile doesn't already have blood
    For I = 1 To MAX_BYTE

        If Blood(I).X = X And Blood(I).Y = Y Then
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
        .timer = GetTickCount
    End With

End Sub

Private Sub HandleAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, X As Long, Y As Long, isCasting As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    AnimationIndex = AnimationIndex + 1

    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1

    With AnimInstance(AnimationIndex)
        .Animation = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .LockType = Buffer.ReadByte
        .lockindex = Buffer.ReadLong
        .isCasting = Buffer.ReadByte
        .Used(0) = True
        .Used(1) = True
    End With

    Buffer.Flush: Set Buffer = Nothing

    ' play the sound if we've got one
    With AnimInstance(AnimationIndex)

        If .LockType = 0 Then
            X = AnimInstance(AnimationIndex).X
            Y = AnimInstance(AnimationIndex).Y
        ElseIf .LockType = TARGET_TYPE_PLAYER Then
            X = GetPlayerX(.lockindex)
            Y = GetPlayerY(.lockindex)
        ElseIf .LockType = TARGET_TYPE_NPC Then
            X = MapNpc(.lockindex).X
            Y = MapNpc(.lockindex).Y
        End If

    End With

    PlayMapSound X, Y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation
End Sub

Private Sub HandleMapNpcVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim I As Long
    Dim MapNpcNum As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MapNpcNum = Buffer.ReadLong

    For I = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(I) = Buffer.ReadLong
    Next

    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleCooldown(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Slot As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Slot = Buffer.ReadLong
    SpellCD(Slot) = GetTickCount
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SpellBuffer = 0
    SpellBufferTimer = 0
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, Access As Long, name As String, message As String, Colour As Long, header As String, PK As Long, saycolour As Long
    Dim Channel As Byte, colStr As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    name = Buffer.ReadString
    Access = Buffer.ReadLong
    PK = Buffer.ReadLong
    message = Buffer.ReadString
    header = Buffer.ReadString
    saycolour = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    
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
    AddText ColourChar & GetColStr(Colour) & header & name & ": " & ColourChar & GetColStr(Grey) & message, Grey, , Channel
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim shopNum As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    shopNum = Buffer.ReadLong
    OpenShop shopNum
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    StunDuration = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim I As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For I = 1 To MAX_BANK
        Bank.Item(I).num = Buffer.ReadLong
        Bank.Item(I).value = Buffer.ReadLong
    Next

    InBank = True
    Buffer.Flush: Set Buffer = Nothing
    
    If Not Windows(GetWindowIndex("winBank")).Window.visible Then
        ShowWindow GetWindowIndex("winBank"), , False
    End If
End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    InTrade = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    
    ShowTrade
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    InTrade = 0
    HideWindow GetWindowIndex("winTrade")
End Sub

Private Sub HandleTradeUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, dataType As Byte, I As Long, yourWorth As Long, theirWorth As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    dataType = Buffer.ReadByte

    If dataType = 0 Then ' ours!
        For I = 1 To MAX_INV
            TradeYourOffer(I).num = Buffer.ReadLong
            TradeYourOffer(I).value = Buffer.ReadLong
        Next
        yourWorth = Buffer.ReadLong
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblYourValue")).text = yourWorth & "g"
    ElseIf dataType = 1 Then 'theirs
        For I = 1 To MAX_INV
            TradeTheirOffer(I).num = Buffer.ReadLong
            TradeTheirOffer(I).value = Buffer.ReadLong
        Next
        theirWorth = Buffer.ReadLong
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblTheirValue")).text = theirWorth & "g"
    End If

    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleTradeStatus(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tradeStatus As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    tradeStatus = Buffer.ReadByte
    Buffer.Flush: Set Buffer = Nothing

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
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    myTarget = Buffer.ReadLong
    myTargetType = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleHotbar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim I As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For I = 1 To MAX_HOTBAR
        Hotbar(I).Slot = Buffer.ReadLong
        Hotbar(I).sType = Buffer.ReadByte
    Next
End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player_HighIndex = MAX_PLAYERS 'Buffer.ReadLong
End Sub

Private Sub HandleResetShopAction(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    UpdateShop
End Sub

Private Sub HandleSound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim X As Long, Y As Long, entityType As Long, entityNum As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    entityType = Buffer.ReadLong
    entityNum = Buffer.ReadLong
    PlayMapSound X, Y, entityType, entityNum
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, theName As String, Top As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    theName = Buffer.ReadString
    ' cache name and show invitation
    diaDataString = theName
    ShowWindow GetWindowIndex("winInvite_Trade")
    Windows(GetWindowIndex("winInvite_Trade")).Controls(GetControlIndex("winInvite_Trade", "btnInvite")).text = ColourChar & White & theName & ColourChar & "-1" & " has invited you to trade."
    AddText Trim$(theName) & " has invited you to trade.", White
    ' loop through
    Top = ScreenHeight - 80
    If Windows(GetWindowIndex("winInvite_Party")).Window.visible Then
        Top = Top - 37
    End If
    Windows(GetWindowIndex("winInvite_Trade")).Window.Top = Top
End Sub

Private Sub HandlePartyInvite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, theName As String, Top As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    theName = Buffer.ReadString
    ' cache name and show invitation popup
    diaDataString = theName
    ShowWindow GetWindowIndex("winInvite_Party")
    Windows(GetWindowIndex("winInvite_Party")).Controls(GetControlIndex("winInvite_Party", "btnInvite")).text = ColourChar & White & theName & ColourChar & "-1" & " has invited you to a party."
    AddText Trim$(theName) & " has invited you to a party.", White
    ' loop through
    Top = ScreenHeight - 80
    If Windows(GetWindowIndex("winInvite_Trade")).Window.visible Then
        Top = Top - 37
    End If
    Windows(GetWindowIndex("winInvite_Party")).Window.Top = Top
End Sub

Private Sub HandlePartyUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, I As Long, inParty As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    inParty = Buffer.ReadByte

    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        UpdatePartyInterface
        ' exit out early
        Exit Sub
    End If

    ' carry on otherwise
    Party.Leader = Buffer.ReadLong

    For I = 1 To MAX_PARTY_MEMBERS
        Party.Member(I) = Buffer.ReadLong
    Next

    Party.MemberCount = Buffer.ReadLong
    
    ' update the party interface
    UpdatePartyInterface
End Sub

Private Sub HandlePartyVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim playerNum As Long
    Dim Buffer As clsBuffer, I As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' which player?
    playerNum = Buffer.ReadLong

    ' set vitals
    For I = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(I) = Buffer.ReadLong
        Player(playerNum).Vital(I) = Buffer.ReadLong
    Next

    ' update the party interface
    UpdatePartyBars
End Sub

Private Sub HandleConvEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long

    With frmEditor_Conv
        Editor = EDITOR_CONV
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_CONVS
            .lstIndex.AddItem I & ": " & Trim$(Conv(I).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ConvEditorInit
    End With

End Sub

Private Sub HandleUpdateConv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Convnum As Long
    Dim Buffer As clsBuffer
    Dim I As Long
    Dim X As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Convnum = Buffer.ReadLong

    With Conv(Convnum)
        .name = Buffer.ReadString
        .chatCount = Buffer.ReadLong
        ReDim Conv(Convnum).Conv(1 To .chatCount)

        For I = 1 To .chatCount
            .Conv(I).Conv = Buffer.ReadString

            For X = 1 To 4
                .Conv(I).rText(X) = Buffer.ReadString
                .Conv(I).rTarget(X) = Buffer.ReadLong
            Next

            .Conv(I).Event = Buffer.ReadLong
            .Conv(I).Data1 = Buffer.ReadLong
            .Conv(I).Data2 = Buffer.ReadLong
            .Conv(I).Data3 = Buffer.ReadLong
        Next

    End With

    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleChatUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, NpcNum As Long, mT As String, o(1 To 4) As String, I As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    NpcNum = Buffer.ReadLong
    mT = Buffer.ReadString
    For I = 1 To 4
        o(I) = Buffer.ReadString
    Next

    Buffer.Flush: Set Buffer = Nothing

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
    Dim Buffer As clsBuffer, TargetType As Long, target As Long, message As String, Colour As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    target = Buffer.ReadLong
    TargetType = Buffer.ReadLong
    message = Buffer.ReadString
    Colour = Buffer.ReadLong
    AddChatBubble target, TargetType, message, Colour
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandlePlayerChars(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, I As Long, winNum As Long, conNum As Long, isSlotEmpty(1 To MAX_CHARS) As Boolean, X As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()

    For I = 1 To MAX_CHARS
        CharName(I) = Trim$(Buffer.ReadString)
        CharSprite(I) = Buffer.ReadLong
        CharAccess(I) = Buffer.ReadLong
        CharClass(I) = Buffer.ReadLong
        ' set as empty or not
        If Not Len(Trim$(CharName(I))) > 0 Then isSlotEmpty(I) = True
    Next

    Buffer.Flush: Set Buffer = Nothing
    
    HideWindows
    ShowWindow GetWindowIndex("winCharacters")
    
    ' set GUI window up
    winNum = GetWindowIndex("winCharacters")
    For I = 1 To MAX_CHARS
        conNum = GetControlIndex("winCharacters", "lblCharName_" & I)
        With Windows(winNum).Controls(conNum)
            If Not isSlotEmpty(I) Then
                .text = CharName(I)
            Else
                .text = "Blank Slot"
            End If
        End With
        ' hide/show buttons
        If isSlotEmpty(I) Then
            ' create button
            conNum = GetControlIndex("winCharacters", "btnCreateChar_" & I)
            Windows(winNum).Controls(conNum).visible = True
            ' select button
            conNum = GetControlIndex("winCharacters", "btnSelectChar_" & I)
            Windows(winNum).Controls(conNum).visible = False
            ' delete button
            conNum = GetControlIndex("winCharacters", "btnDelChar_" & I)
            Windows(winNum).Controls(conNum).visible = False
        Else
            ' create button
            conNum = GetControlIndex("winCharacters", "btnCreateChar_" & I)
            Windows(winNum).Controls(conNum).visible = False
            ' select button
            conNum = GetControlIndex("winCharacters", "btnSelectChar_" & I)
            Windows(winNum).Controls(conNum).visible = True
            ' delete button
            conNum = GetControlIndex("winCharacters", "btnDelChar_" & I)
            Windows(winNum).Controls(conNum).visible = True
        End If
    Next
End Sub

Private Sub HandleCancelAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim theIndex As Long, Buffer As clsBuffer, I As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    theIndex = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    ' find the casting animation
    For I = 1 To MAX_BYTE
        If AnimInstance(I).LockType = TARGET_TYPE_PLAYER Then
            If AnimInstance(I).lockindex = theIndex Then
                If AnimInstance(I).isCasting = 1 Then
                    ' clear it
                    ClearAnimInstance I
                End If
            End If
        End If
    Next
End Sub

Private Sub HandlePlayerVariables(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, I As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For I = 1 To MAX_BYTE
        Player(MyIndex).Variable(I) = Buffer.ReadLong
    Next
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleEvent(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    If Buffer.ReadLong = 1 Then
        inEvent = True
    Else
        inEvent = False
    End If
    eventNum = Buffer.ReadLong
    eventPageNum = Buffer.ReadLong
    eventCommandNum = Buffer.ReadLong
    
    Buffer.Flush: Set Buffer = Nothing
End Sub
