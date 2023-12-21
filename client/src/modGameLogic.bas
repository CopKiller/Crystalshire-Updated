Attribute VB_Name = "modGameLogic"
Option Explicit

Public Sub GameLoop()
    Dim i As Long, X As Long, Y As Long
    Dim barDifference As Long
    On Error GoTo retry
    
    ' *** Start GameLoop ***
    Do While InGame
retry:
        Loops = 0
        ' *** Start GameLoop ***
        Do While InGame And frmMain.WindowState <> vbMinimized And Loops < MAX_FRAME_SKIP
            Tick = getTime                            ' Set the inital tick
            ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
            FrameTime = Tick                               ' Set the time second loop time to the first.
            
            If Thread = False Then
                GameLooptmr = Tick + 25
            End If
    
            ' handle input
            If GetForegroundWindow() = frmMain.hWnd Then
                HandleMouseInput
            End If
    
            ' * Check surface timers *
            ' Sprites
            If tmr10000 < Tick Then
                ' check ping
                Call GetPing
                tmr10000 = Tick + 10000
            End If
    
            If tmr25 < Tick Then
                InGame = IsConnected
                Call CheckKeys ' Check to make sure they aren't trying to auto do anything
    
                If GetForegroundWindow() = frmMain.hWnd Then
                    Call CheckInputKeys ' Check which keys were pressed
                End If
    
                ' check if we need to end the CD icon
                If CountSpellicon > 0 Then
                    For i = 1 To MAX_PLAYER_SPELLS
                        If PlayerSpells(i).Spell > 0 Then
                            If SpellCD(i) > 0 Then
                                If SpellCD(i) + (Spell(PlayerSpells(i).Spell).CDTime * 1000) < Tick Then
                                    SpellCD(i) = 0
                                End If
                            End If
                        End If
                    Next
                End If
    
                ' check if we need to unlock the player's spell casting restriction
                If SpellBuffer > 0 Then
                    If SpellBufferTimer + (Spell(PlayerSpells(SpellBuffer).Spell).CastTime * 1000) < Tick Then
                        SpellBuffer = 0
                        SpellBufferTimer = 0
                    End If
                End If
    
                If CanMoveNow Then
                    Call CheckMovement ' Check if player is trying to move
                    Call CheckAttack   ' Check to see if player is trying to attack
                End If
    
                For i = 1 To MAX_BYTE
                    CheckAnimInstance i
                Next
                
                ' appear tile logic
                AppearTileFadeLogic
                CheckAppearTiles
    
                tmr25 = Tick + 25
            End If
    
            ' targetting
            If targetTmr < Tick Then
                If tabDown Then
                    FindNearestTarget
                End If
    
                targetTmr = Tick + 50
            End If
            
            ' chat timer
            If chatTmr < Tick Then
                ' scrolling
                If ChatButtonUp Then
                    ScrollChatBox 0
                End If
    
                If ChatButtonDown Then
                    ScrollChatBox 1
                End If
                
                ' remove messages
                If chatLastRemove + CHAT_DIFFERENCE_TIMER < getTime Then
                    ' remove timed out messages from chat
                    For i = Chat_HighIndex To 1 Step -1
                        If Len(Chat(i).text) > 0 Then
                            If Chat(i).visible Then
                                If Chat(i).timer + CHAT_TIMER < Tick Then
                                    Chat(i).visible = False
                                    chatLastRemove = getTime
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
    
                chatTmr = Tick + 50
            End If
            
            If tmr45 <= Tick Then
                For i = 1 To LastProjectile
                    If MapProjectile(i).Owner > 0 Then
                        If MapProjectile(i).curAnim = 8 Then
                            MapProjectile(i).curAnim = 3
                        ElseIf MapProjectile(i).curAnim < 8 Then
                            MapProjectile(i).curAnim = MapProjectile(i).curAnim + 1
                        ElseIf MapProjectile(i).curAnim > 8 And MapProjectile(i).curAnim < 11 Then
                            MapProjectile(i).curAnim = MapProjectile(i).curAnim + 1
                        End If
                    End If
                Next
                
                tmr45 = Tick + 45
            End If
    
            ' fog scrolling
            If fogTmr < Tick Then
                If CurrentFogSpeed > 0 Then
                    ' move
                    fogOffsetX = fogOffsetX - 1
                    fogOffsetY = fogOffsetY - 1
        
                    ' reset
                    If fogOffsetX < -256 Then fogOffsetX = 0
                    If fogOffsetY < -256 Then fogOffsetY = 0
                    
                    ' reset timer
                    fogTmr = Tick + 255 - CurrentFogSpeed
                End If
            End If
    
            ' elastic bars
            If barTmr < Tick Then
                SetBarWidth BarWidth_GuiHP_Max, BarWidth_GuiHP
                SetBarWidth BarWidth_GuiSP_Max, BarWidth_GuiSP
                SetBarWidth BarWidth_GuiEXP_Max, BarWidth_GuiEXP
                For i = 1 To MAX_MAP_NPCS
                    If MapNpc(i).Num > 0 Then
                        SetBarWidth BarWidth_NpcHP_Max(i), BarWidth_NpcHP(i)
                    End If
                Next
    
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        SetBarWidth BarWidth_PlayerHP_Max(i), BarWidth_PlayerHP(i)
                    End If
                Next
    
                ' reset timer
                barTmr = Tick + 10
            End If
    
            ' Animations!
            If mapTimer < Tick Then
    
                ' animate waterfalls
                Select Case waterfallFrame
    
                    Case 0
                        waterfallFrame = 1
    
                    Case 1
                        waterfallFrame = 2
    
                    Case 2
                        waterfallFrame = 0
                End Select
    
                ' animate autotiles
                Select Case autoTileFrame
    
                    Case 0
                        autoTileFrame = 1
    
                    Case 1
                        autoTileFrame = 2
    
                    Case 2
                        autoTileFrame = 0
                End Select
    
                ' animate textbox
                If chatShowLine = "|" Then
                    chatShowLine = vbNullString
                Else
                    chatShowLine = "|"
                End If
                
                ' re-set timer
                mapTimer = Tick + 500
            End If
            
            Call ProcessWeather
    
            ' Process input before rendering, otherwise input will be behind by 1 frame
            If WalkTimer < Tick Then
    
                For i = 1 To Player_HighIndex
    
                    If IsPlaying(i) Then
                        Call ProcessMovement(i)
                    End If
    
                Next i
    
                ' Process npc movements (actually move them)
                For i = 1 To Npc_HighIndex
    
                    If Map.MapData.Npc(i) > 0 Then
                        Call ProcessNpcMovement(i)
                    End If
    
                Next i
    
                WalkTimer = Tick + 30 ' edit this value to change WalkTimer
            End If
    
            ' *********************
            ' ** Render Graphics **
            ' *********************
            If Thread = False Then
                Call Render_Graphics
                Call UpdateSounds
            
                If Options.FPSLock And FPS > 60 Then
                    Tick = Tick + SKIP_TICKS
                    Loops = Loops + 1
                End If
                
                ' Calculate fps
                If TickFPS <= Tick Then
                    GameFPS = FPS
                    TickFPS = Tick + 1000
                    FPS = 0
                Else
                    FPS = FPS + 1
                End If
            
                If Options.FPSLock And FPS > 60 Then
                    Sleep SKIP_TICKS
                End If
            End If
            
            DoEvents
            
            If Thread And GameLooptmr > Tick Then
                Thread = False
                Exit Sub
            End If
    
        Loop
        ' Mute everything but still keep everything playing
        If frmMain.WindowState = vbMinimized Then
            Stop_Music
        End If
        
        Sleep MAX_FRAME_SKIP
        DoEvents
    Loop

    If InGame Then GoTo retry

    If isLogging Then
        isLogging = False
        MenuLoop
        GettingMap = True
        Stop_Music
        Play_Music MenuMusic
    Else
        ' Shutdown the game
        Call SetStatus("Destroying game data.")
        Call DestroyGame
    End If

End Sub

Public Sub MenuLoop()
    Dim FrameTime As Long, Tick As Long, TickFPS As Long, FPS As Long, tmr500 As Long, fadeTmr As Long

    ' *** Start MenuLoop ***
    Do While inMenu
retry:
        Loops = 0
        ' *** Start GameLoop ***
        Do While inMenu And frmMain.WindowState <> vbMinimized And Loops < MAX_FRAME_SKIP
            Tick = getTime                            ' Set the inital tick
            ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
            FrameTime = Tick                               ' Set the time second loop time to the first.
    
            ' handle input
            If GetForegroundWindow() = frmMain.hWnd Then
                HandleMouseInput
            End If
            
            ' Animations!
            If tmr500 < Tick Then
                ' animate textbox
                If chatShowLine = "|" Then
                    chatShowLine = vbNullString
                Else
                    chatShowLine = "|"
                End If
    
                ' re-set timer
                tmr500 = Tick + 500
            End If
            
            ' trailer
            If videoPlaying Then VideoLoop
            
            ' fading
            If fadeTmr < Tick Then
                If Not videoPlaying Then
                    If fadeAlpha > 5 Then
                        ' lower fade
                        fadeAlpha = fadeAlpha - 5
                    Else
                        fadeAlpha = 0
                    End If
                End If
                fadeTmr = Tick + 1
            End If
    
            ' *********************
            ' ** Render Graphics **
            ' *********************
            Call Render_Menu
            
            If Options.FPSLock And FPS > 60 Then
                Tick = Tick + SKIP_TICKS
                Loops = Loops + 1
            End If
            
            ' Calculate fps
            If TickFPS <= Tick Then
                GameFPS = FPS
                TickFPS = Tick + 1000
                FPS = 0
            Else
                FPS = FPS + 1
            End If
            
            If Options.FPSLock And FPS > 60 Then
                Sleep SKIP_TICKS
            End If
            
            DoEvents
    
        Loop
        
        ' Mute everything but still keep everything playing
        If frmMain.WindowState = vbMinimized Then
            Stop_Music
        End If
        
        Sleep MAX_FRAME_SKIP
        DoEvents
    Loop

End Sub

Public Sub ProcessMovement(ByVal Index As Long)
    Dim MovementSpeed As Long
    
    ' Check if player is walking, and if so process moving them over
    Select Case Player(Index).Moving
            Case MOVING_RUNNING: MovementSpeed = RUN_SPEED
            Case MOVING_WALKING: MovementSpeed = WALK_SPEED
        Case Else: Exit Sub
    End Select
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            Player(Index).yOffset = Player(Index).yOffset - MovementSpeed
            If Player(Index).yOffset < 0 Then Player(Index).yOffset = 0
        Case DIR_DOWN
            Player(Index).yOffset = Player(Index).yOffset + MovementSpeed
            If Player(Index).yOffset > 0 Then Player(Index).yOffset = 0
        Case DIR_LEFT
            Player(Index).xOffset = Player(Index).xOffset - MovementSpeed
            If Player(Index).xOffset < 0 Then Player(Index).xOffset = 0
        Case DIR_RIGHT
            Player(Index).xOffset = Player(Index).xOffset + MovementSpeed
            If Player(Index).xOffset > 0 Then Player(Index).xOffset = 0
        Case DIR_UP_LEFT
            Player(Index).yOffset = Player(Index).yOffset - MovementSpeed
            If Player(Index).yOffset < 0 Then Player(Index).yOffset = 0
            Player(Index).xOffset = Player(Index).xOffset - MovementSpeed
            If Player(Index).xOffset < 0 Then Player(Index).xOffset = 0
        
        Case DIR_UP_RIGHT
            Player(Index).yOffset = Player(Index).yOffset - MovementSpeed
            If Player(Index).yOffset < 0 Then Player(Index).yOffset = 0
            Player(Index).xOffset = Player(Index).xOffset + MovementSpeed
            If Player(Index).xOffset > 0 Then Player(Index).xOffset = 0

        Case DIR_DOWN_LEFT
            Player(Index).yOffset = Player(Index).yOffset + MovementSpeed
            If Player(Index).yOffset > 0 Then Player(Index).yOffset = 0
            Player(Index).xOffset = Player(Index).xOffset - MovementSpeed
            If Player(Index).xOffset < 0 Then Player(Index).xOffset = 0
        
        Case DIR_DOWN_RIGHT
            Player(Index).yOffset = Player(Index).yOffset + MovementSpeed
            If Player(Index).yOffset > 0 Then Player(Index).yOffset = 0
            Player(Index).xOffset = Player(Index).xOffset + MovementSpeed
            If Player(Index).xOffset > 0 Then Player(Index).xOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If Player(Index).Moving > 0 Then
        If GetPlayerDir(Index) = DIR_RIGHT Or GetPlayerDir(Index) = DIR_DOWN Or GetPlayerDir(Index) = DIR_DOWN_RIGHT Then
            If (Player(Index).xOffset >= 0) And (Player(Index).yOffset >= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 0 Then
                    Player(Index).Step = 2
                Else
                    Player(Index).Step = 0
                End If
            End If
        Else
            If (Player(Index).xOffset <= 0) And (Player(Index).yOffset <= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 0 Then
                    Player(Index).Step = 2
                Else
                    Player(Index).Step = 0
                End If
            End If
        End If
    End If
    
End Sub

Public Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
    Dim MovementSpeed As Long

    ' Check if NPC is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        MovementSpeed = RUN_SPEED
    Else
        Exit Sub
    End If

    Select Case MapNpc(MapNpcNum).Dir

        Case DIR_UP
            MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset - MovementSpeed

            If MapNpc(MapNpcNum).yOffset < 0 Then MapNpc(MapNpcNum).yOffset = 0

        Case DIR_DOWN
            MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset + MovementSpeed

            If MapNpc(MapNpcNum).yOffset > 0 Then MapNpc(MapNpcNum).yOffset = 0

        Case DIR_LEFT
            MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset - MovementSpeed

            If MapNpc(MapNpcNum).xOffset < 0 Then MapNpc(MapNpcNum).xOffset = 0

        Case DIR_RIGHT
            MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset + MovementSpeed

            If MapNpc(MapNpcNum).xOffset > 0 Then MapNpc(MapNpcNum).xOffset = 0
        
        Case DIR_UP_LEFT
                MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset - MovementSpeed
                If MapNpc(MapNpcNum).yOffset < 0 Then MapNpc(MapNpcNum).yOffset = 0
                MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset - MovementSpeed
                If MapNpc(MapNpcNum).xOffset < 0 Then MapNpc(MapNpcNum).xOffset = 0
            
            Case DIR_UP_RIGHT
                MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset - MovementSpeed
                If MapNpc(MapNpcNum).yOffset < 0 Then MapNpc(MapNpcNum).yOffset = 0
                MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset + MovementSpeed
                If MapNpc(MapNpcNum).xOffset > 0 Then MapNpc(MapNpcNum).xOffset = 0
    
            Case DIR_DOWN_LEFT
                MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset + MovementSpeed
                If MapNpc(MapNpcNum).yOffset > 0 Then MapNpc(MapNpcNum).yOffset = 0
                MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset - MovementSpeed
                If MapNpc(MapNpcNum).xOffset < 0 Then MapNpc(MapNpcNum).xOffset = 0
            
            Case DIR_DOWN_RIGHT
                MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset + MovementSpeed
                If MapNpc(MapNpcNum).yOffset > 0 Then MapNpc(MapNpcNum).yOffset = 0
                MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset + MovementSpeed
                If MapNpc(MapNpcNum).xOffset > 0 Then MapNpc(MapNpcNum).xOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If MapNpc(MapNpcNum).Moving > 0 Then
        If MapNpc(MapNpcNum).Dir = DIR_RIGHT Or MapNpc(MapNpcNum).Dir = DIR_DOWN Or MapNpc(MapNpcNum).Dir = DIR_DOWN_RIGHT Then
            If (MapNpc(MapNpcNum).xOffset >= 0) And (MapNpc(MapNpcNum).yOffset >= 0) Then
                MapNpc(MapNpcNum).Moving = 0

                If MapNpc(MapNpcNum).Step = 0 Then
                    MapNpc(MapNpcNum).Step = 2
                Else
                    MapNpc(MapNpcNum).Step = 0
                End If
            End If

        Else

            If (MapNpc(MapNpcNum).xOffset <= 0) And (MapNpc(MapNpcNum).yOffset <= 0) Then
                MapNpc(MapNpcNum).Moving = 0

                If MapNpc(MapNpcNum).Step = 0 Then
                    MapNpc(MapNpcNum).Step = 2
                Else
                    MapNpc(MapNpcNum).Step = 0
                End If
            End If
        End If
    End If

End Sub

Sub CheckMapGetItem()
    Dim buffer As New clsBuffer, tmpIndex As Long, i As Long, X As Long
    Set buffer = New clsBuffer

    If getTime > Player(MyIndex).MapGetTimer + 250 Then

        ' find out if we want to pick it up
        For i = 1 To MAX_MAP_ITEMS


            If MapItem(i).X = Player(MyIndex).X And MapItem(i).Y = Player(MyIndex).Y Then
                If MapItem(i).Num > 0 Then
                    If Item(MapItem(i).Num).BindType = 1 Then

                        ' make sure it's not a party drop
                        If Party.Leader > 0 Then

                            For X = 1 To MAX_PARTY_MEMBERS
                                tmpIndex = Party.Member(X)

                                If tmpIndex > 0 Then
                                    If Trim$(GetPlayerName(tmpIndex)) = Trim$(MapItem(i).playerName) Then
                                        If Item(MapItem(i).Num).ClassReq > 0 Then
                                            If Item(MapItem(i).Num).ClassReq <> Player(MyIndex).Class Then
                                                Dialogue "Loot Check", "This item is BoP and is not for your class.", "Are you sure you want to pick it up?", TypeLOOTITEM, StyleYESNO
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If

                            Next

                        End If

                    Else
                        'not bound
                        Exit For
                    End If
                End If
            End If

        Next

        ' nevermind, pick it up
        Player(MyIndex).MapGetTimer = getTime
        buffer.WriteLong CMapGetItem
        SendData buffer.ToArray()
    End If

    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub CheckAttack()
    Dim buffer As clsBuffer
    Dim attackspeed As Long

    If ControlDown Then
        If SpellBuffer > 0 Then Exit Sub ' currently casting a spell, can't attack
        If StunDuration > 0 Then Exit Sub ' stunned, can't attack

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(MyIndex, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If Player(MyIndex).AttackTimer + attackspeed < getTime Then
            If Player(MyIndex).Attacking = 0 Then

                With Player(MyIndex)
                    .Attacking = 1
                    .AttackTimer = getTime
                End With

                Set buffer = New clsBuffer
                buffer.WriteLong CAttack
                SendData buffer.ToArray()
                buffer.Flush: Set buffer = Nothing
            End If
        End If
    End If

End Sub

Function IsTryingToMove() As Boolean

    'If DirUp Or DirDown Or DirLeft Or DirRight Then
    If DirUp Or DirLeft Or DirDown Or DirRight Then
        IsTryingToMove = True
    End If

End Function

Function CanMove() As Boolean
    Dim d As Long
    CanMove = True

    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they haven't just casted a spell
    'If SpellBuffer > 0 Then
    '    CanMove = False
    '    Exit Function
    'End If

    ' make sure they're not stunned
    If StunDuration > 0 Then
        CanMove = False
        Exit Function
    End If

    ' make sure they're not in a shop
    If InShop > 0 Then
        CanMove = False
        Exit Function
    End If

    ' not in bank
    If InBank Then
        CanMove = False
        Exit Function
    End If

    If inTutorial Then
        CanMove = False
        Exit Function
    End If

    d = GetPlayerDir(MyIndex)

    If DirUp And DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_UP_LEFT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 And GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_UP_LEFT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP_LEFT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If GetPlayerY(MyIndex) <= 0 Then
                If Map.MapData.Up > 0 Then
                    Call MapEditorLeaveMap
                    Call SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
            ElseIf GetPlayerX(MyIndex) <= 0 Then
                If Map.MapData.Left > 0 Then
                    Call MapEditorLeaveMap
                    Call SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
            End If
            
            CanMove = False
            Exit Function
        End If
    End If
'#######################################################################################################################
'#######################################################################################################################
    If DirUp And DirRight Then
        Call SetPlayerDir(MyIndex, DIR_UP_RIGHT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 And GetPlayerX(MyIndex) < Map.MapData.MaxX Then
            If CheckDirection(DIR_UP_RIGHT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP_RIGHT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If GetPlayerY(MyIndex) <= 0 Then
                If Map.MapData.Up > 0 Then
                    Call MapEditorLeaveMap
                    Call SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
            ElseIf GetPlayerX(MyIndex) >= Map.MapData.MaxX Then
                If Map.MapData.Right > 0 Then
                    Call MapEditorLeaveMap
                    Call SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
            End If
            
            CanMove = False
            Exit Function
        End If
    End If
'#######################################################################################################################
'#######################################################################################################################
    If DirDown And DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_DOWN_LEFT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.MapData.MaxY And GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_DOWN_LEFT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN_LEFT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If GetPlayerY(MyIndex) >= Map.MapData.MaxY Then
                If Map.MapData.Down > 0 Then
                    Call MapEditorLeaveMap
                    Call SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
            ElseIf GetPlayerX(MyIndex) <= 0 Then
                If Map.MapData.Left > 0 Then
                    Call MapEditorLeaveMap
                    Call SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
            End If
            
            CanMove = False
            Exit Function
        End If
    End If
'#######################################################################################################################
'#######################################################################################################################
    If DirDown And DirRight Then
        Call SetPlayerDir(MyIndex, DIR_DOWN_RIGHT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.MapData.MaxY And GetPlayerX(MyIndex) < Map.MapData.MaxX Then
            If CheckDirection(DIR_DOWN_RIGHT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN_RIGHT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If GetPlayerY(MyIndex) >= Map.MapData.MaxX Then
                If Map.MapData.Down > 0 Then
                    Call MapEditorLeaveMap
                    Call SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
            ElseIf GetPlayerX(MyIndex) >= Map.MapData.MaxX Then
                If Map.MapData.Right > 0 Then
                    Call MapEditorLeaveMap
                    Call SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
            End If
            
            CanMove = False
            Exit Function
        End If
        Exit Function
    End If
'#######################################################################################################################
'#######################################################################################################################
    If DirUp And Not DirLeft And Not DirRight Then
        Call SetPlayerDir(MyIndex, DIR_UP)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.MapData.Up > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If
'#######################################################################################################################
'#######################################################################################################################
    If DirDown And Not DirLeft And Not DirRight Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.MapData.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.MapData.Down > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If
'#######################################################################################################################
'#######################################################################################################################
    If DirLeft And Not DirDown And Not DirUp Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.MapData.Left > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If
'#######################################################################################################################
'#######################################################################################################################
    If DirRight And Not DirDown And Not DirUp Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < Map.MapData.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.MapData.Right > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

End Function

Function CheckDirection(ByVal Direction As Byte) As Boolean

    Dim X As Long, Y As Long, i As Long
    
    CheckDirection = False

    If GettingMap Then Exit Function

    ' check directional blocking
    If Direction <= DIR_RIGHT Then
        If isDirBlocked(Map.TileData.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, Direction + 1) Then
            CheckDirection = True
            Exit Function
        End If
    Else
        Select Case Direction
            Case DIR_UP_LEFT, DIR_DOWN_LEFT
                If isDirBlocked(Map.TileData.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, DIR_LEFT + 1) Then
                    CheckDirection = True
                    Exit Function
                End If
'#######################################################################################################################
            Case DIR_UP_RIGHT, DIR_DOWN_RIGHT
                If isDirBlocked(Map.TileData.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, DIR_RIGHT + 1) Then
                    CheckDirection = True
                    Exit Function
                End If
        End Select
    End If

    Select Case Direction
        Case DIR_UP
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) - 1
'#######################################################################################################################
        Case DIR_DOWN
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) + 1
'#######################################################################################################################
        Case DIR_LEFT
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex)
'#######################################################################################################################
        Case DIR_RIGHT
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex)
'#######################################################################################################################
        Case DIR_UP_LEFT
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex) - 1
'#######################################################################################################################
        Case DIR_UP_RIGHT
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex) - 1
'#######################################################################################################################
        Case DIR_DOWN_LEFT
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex) + 1
'#######################################################################################################################
        Case DIR_DOWN_RIGHT
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex) + 1
    End Select

    ' Check to see if the map tile is blocked or not
    If Map.TileData.Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If Map.TileData.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the key door is open or not
    If Map.TileData.Tile(X, Y).Type = TILE_TYPE_KEY Then
        ' This actually checks if its open or not
        If TempTile(X, Y).DoorOpen = 0 Then
            CheckDirection = True
            Exit Function
        End If
    End If

    ' Check to see if a player is already on that tile
    If Map.MapData.Moral = 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then

                If GetPlayerX(i) = X Then
                    If GetPlayerY(i) = Y Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If

    ' Check to see if a npc is already on that tile
    For i = 1 To Npc_HighIndex
        If MapNpc(i).Num > 0 Then
            If MapNpc(i).X = X Then
                If MapNpc(i).Y = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' check if it's a drop warp - avoid if walking
    If ShiftDown Then
        If Map.TileData.Tile(X, Y).Type = TILE_TYPE_WARP Then
            If Map.TileData.Tile(X, Y).Data4 Then
                CheckDirection = True
                Exit Function
            End If
        End If
    End If

End Function

Sub CheckMovement()
    Dim X As Long, Y As Long
    With Player(MyIndex)
        If Not GettingMap Then
            If IsTryingToMove Then
                If CanMove Then
                    X = GetPlayerX(MyIndex)
                    Y = GetPlayerY(MyIndex)
                    ' Check if player has the shift key down for running
                    If ShiftDown Then
                        .Moving = MOVING_RUNNING
                    Else
                        .Moving = MOVING_WALKING
                    End If
                    
                    Call SendPlayerMove
        
                    Select Case GetPlayerDir(MyIndex)
                        Case DIR_UP
                            Y = Y - 1: .yOffset = PIC_Y
                        Case DIR_DOWN
                            Y = Y + 1: .yOffset = PIC_Y * -1
                        Case DIR_LEFT
                            X = X - 1: .xOffset = PIC_X
                        Case DIR_RIGHT
                            X = X + 1: .xOffset = PIC_X * -1
                        Case DIR_UP_LEFT
                            Y = Y - 1: X = X - 1
                            .yOffset = PIC_Y: .xOffset = PIC_X
                        Case DIR_UP_RIGHT
                            Y = Y - 1: X = X + 1
                            .yOffset = PIC_Y: .xOffset = PIC_X * -1
                        Case DIR_DOWN_LEFT
                            Y = Y + 1: X = X - 1
                            .yOffset = PIC_Y * -1: .xOffset = PIC_X
                        Case DIR_DOWN_RIGHT
                            Y = Y + 1: X = X + 1
                            .yOffset = PIC_Y * -1: .xOffset = PIC_X * -1
                    End Select
                    
                    ' Check map boundaries
                    If X < 0 Or X > Map.MapData.MaxX Then Exit Sub
                    If Y < 0 Or Y > Map.MapData.MaxY Then Exit Sub
                    
                    Call SetPlayerY(MyIndex, Y)
                    Call SetPlayerX(MyIndex, X)
        
                    If Map.TileData.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                        GettingMap = True
                    End If
                End If
            End If
        End If
    End With

End Sub

Public Function isInBounds()

    If (CurX >= 0) Then
        If (CurX <= Map.MapData.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.MapData.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MapData.MaxX Then Exit Function
    If Y > Map.MapData.MaxY Then Exit Function
    IsValidMapPoint = True
End Function

Public Function IsItem(StartX As Long, StartY As Long) As Long
Dim tempRec As RECT
Dim i As Long
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) Then
            With tempRec
                .Top = StartY + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = StartX + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
                If currMouseY >= tempRec.Top And currMouseY <= tempRec.Bottom Then
                    IsItem = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function IsTrade(StartX As Long, StartY As Long) As Long
Dim tempRec As RECT
Dim i As Long

    For i = 1 To MAX_INV
        With tempRec
            .Top = StartY + TradeTop + ((TradeOffsetY + 32) * ((i - 1) \ TradeColumns))
            .Bottom = .Top + PIC_Y
            .Left = StartX + TradeLeft + ((TradeOffsetX + 32) * (((i - 1) Mod TradeColumns)))
            .Right = .Left + PIC_X
        End With

        If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
            If currMouseY >= tempRec.Top And currMouseY <= tempRec.Bottom Then
                IsTrade = i
                Exit Function
            End If
        End If
    Next
End Function

Public Function IsOffer(StartX As Long, StartY As Long) As Long
    Dim tempRec As RECT
    Dim i As Long
    For i = 1 To MAX_OFFER
    
        If inOffer(i) > 0 Then

            With tempRec
                .Top = StartY + OfferTop + ((OfferOffsetY + 45) * ((i - 1) \ OfferColumns))
                .Bottom = .Top + 45
                .Left = StartX + OfferLeft + ((OfferOffsetX + 485) * (((i - 1) Mod OfferColumns)))
                .Right = .Left + 485
            End With
            
            If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
                RenderTexture TextureDesign(7), ConvertMapX(OfferTop), ConvertMapY(OfferLeft), 0, 0, 32, 32, 32, 32

                If currMouseY >= tempRec.Top And currMouseY <= tempRec.Bottom Then
                    IsOffer = i
                    Exit Function
                End If
            End If
        End If
        
    Next

End Function

Public Function IsBankItem(StartX As Long, StartY As Long) As Long
    Dim tempRec As RECT
    Dim i As Long
    For i = 1 To MAX_BANK
    
        If Bank.Item(i).Num > 0 Then

            With tempRec
                .Top = StartY + BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = StartX + BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With

            If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
                If currMouseY >= tempRec.Top And currMouseY <= tempRec.Bottom Then
                    IsBankItem = i
                    Exit Function
                End If
            End If
        End If
        
    Next

End Function

Public Function IsEqItem(StartX As Long, StartY As Long) As Long
Dim tempRec As RECT
Dim i As Long
    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(MyIndex, i) Then
            With tempRec
                .Top = StartY + EqTop + (32 * ((i - 1) \ EqColumns))
                .Bottom = .Top + PIC_Y
                .Left = StartX + EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
                If currMouseY >= tempRec.Top And currMouseY <= tempRec.Bottom Then
                    IsEqItem = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function IsSkill(StartX As Long, StartY As Long) As Long
Dim tempRec As RECT
Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS
        If PlayerSpells(i).Spell Then
            With tempRec
                .Top = StartY + SkillTop + ((SkillOffsetY + 32) * ((i - 1) \ SkillColumns))
                .Bottom = .Top + PIC_Y
                .Left = StartX + SkillLeft + ((SkillOffsetX + 32) * (((i - 1) Mod SkillColumns)))
                .Right = .Left + PIC_X
            End With

            If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
                If currMouseY >= tempRec.Top And currMouseY <= tempRec.Bottom Then
                    IsSkill = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function IsHotbar(StartX As Long, StartY As Long) As Long
Dim tempRec As RECT
Dim i As Long

    For i = 1 To MAX_HOTBAR
        If Hotbar(i).Slot Then
            With tempRec
                .Top = StartY + HotbarTop
                .Bottom = .Top + PIC_Y
                .Left = StartX + HotbarLeft + ((i - 1) * HotbarOffsetX)
                .Right = .Left + PIC_X
            End With

            If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
                If currMouseY >= tempRec.Top And currMouseY <= tempRec.Bottom Then
                    IsHotbar = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Sub UseItem()

    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)
End Sub

Public Sub ForgetSpell(ByVal spellSlot As Long)
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If

    ' dont let them forget a spell which is in CD
    If SpellCD(spellSlot) > 0 Then
        AddText "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If

    ' dont let them forget a spell which is buffered
    If SpellBuffer = spellSlot Then
        AddText "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If

    If PlayerSpells(spellSlot).Spell > 0 Then
        Set buffer = New clsBuffer
        buffer.WriteLong CForgetSpell
        buffer.WriteLong spellSlot
        SendData buffer.ToArray()
        buffer.Flush: Set buffer = Nothing
    Else
        AddText "No spell here.", BrightRed
    End If

End Sub

Public Sub CastSpell(ByVal spellSlot As Long)
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If

    If SpellCD(spellSlot) > 0 Then
        AddText "Spell has not cooled down yet!", BrightRed
        Exit Sub
    End If
    
    ' make sure we're not casting same spell
    If SpellBuffer > 0 Then
        If SpellBuffer = spellSlot Then
            ' stop them
            Exit Sub
        End If
    End If

    If PlayerSpells(spellSlot).Spell = 0 Then Exit Sub

    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < Spell(PlayerSpells(spellSlot).Spell).MPCost Then
        Call AddText("Not enough MP to cast " & Trim$(Spell(PlayerSpells(spellSlot).Spell).Name) & ".", BrightRed)
        Exit Sub
    End If

    If PlayerSpells(spellSlot).Spell > 0 Then
        If getTime > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Set buffer = New clsBuffer
                buffer.WriteLong CCast
                buffer.WriteLong spellSlot
                SendData buffer.ToArray()
                buffer.Flush: Set buffer = Nothing
                SpellBuffer = spellSlot
                SpellBufferTimer = getTime
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If

    Else
        Call AddText("No spell here.", BrightRed)
    End If

End Sub

Sub ClearTempTile()
    Dim X As Long
    Dim Y As Long
    ReDim TempTile(0 To Map.MapData.MaxX, 0 To Map.MapData.MaxY)

    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            TempTile(X, Y).DoorOpen = 0

            If Not GettingMap Then cacheRenderState X, Y, MapLayer.Mask
        Next
    Next

End Sub

Public Sub DevMsg(ByVal text As String, ByVal Color As Byte)

    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(text, Color)
        End If
    End If
    
End Sub

Public Function TwipsToPixels(ByVal twip_val As Long, ByVal XorY As Byte) As Long

    If XorY = 0 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelY
    End If

End Function

Public Function PixelsToTwips(ByVal pixel_val As Long, ByVal XorY As Byte) As Long

    If XorY = 0 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelY
    End If

End Function

Public Function ConvertCurrency(ByVal Amount As Long) As String

    If Int(Amount) < 10000 Then
        ConvertCurrency = Amount
    ElseIf Int(Amount) < 999999 Then
        ConvertCurrency = Int(Amount / 1000) & "k"
    ElseIf Int(Amount) < 999999999 Then
        ConvertCurrency = Int(Amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(Amount / 1000000000) & "b"
    End If

End Function

Public Sub CacheResources()
    Dim X As Long, Y As Long, Resource_Count As Long
    Resource_Count = 0

    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY

            If Map.TileData.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).X = X
                MapResource(Resource_Count).Y = Y
            End If

        Next
    Next

    Resource_Index = Resource_Count
End Sub

Public Sub CreateActionMsg(ByVal message As String, ByVal Color As Integer, ByVal MsgType As Byte, ByVal X As Long, ByVal Y As Long)
    Dim i As Long
    ActionMsgIndex = ActionMsgIndex + 1

    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .message = message
        .Color = Color
        .Type = MsgType
        .Created = getTime
        .Scroll = 1
        .X = X
        .Y = Y
        .alpha = 255
    End With

    If ActionMsg(ActionMsgIndex).Type = ACTIONMsgSCROLL Then
        ActionMsg(ActionMsgIndex).Y = ActionMsg(ActionMsgIndex).Y + Rand(-2, 6)
        ActionMsg(ActionMsgIndex).X = ActionMsg(ActionMsgIndex).X + Rand(-8, 8)
    End If

    ' find the new high index
    For i = MAX_BYTE To 1 Step -1

        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If

    Next

    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
End Sub

Public Sub ClearActionMsg(ByVal Index As Byte)
    Dim i As Long
    ActionMsg(Index) = EmptyActionMsg
    ActionMsg(Index).message = vbNullString

    ' find the new high index
    For i = MAX_BYTE To 1 Step -1

        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If

    Next

    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
End Sub

Public Sub CheckAnimInstance(ByVal Index As Long)
    Dim looptime As Long
    Dim Layer As Long
    Dim FrameCount As Long

    ' if doesn't exist then exit sub
    If AnimInstance(Index).Animation <= 0 Then Exit Sub
    If AnimInstance(Index).Animation >= MAX_ANIMATIONS Then Exit Sub

    For Layer = 0 To 1

        If AnimInstance(Index).Used(Layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(Layer)

            FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)

            ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).frameIndex(Layer) = 0 Then AnimInstance(Index).frameIndex(Layer) = 1
            If AnimInstance(Index).LoopIndex(Layer) = 0 Then AnimInstance(Index).LoopIndex(Layer) = 1

            ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).timer(Layer) + looptime <= getTime Then

                ' check if out of range
                If AnimInstance(Index).frameIndex(Layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(Layer) = AnimInstance(Index).LoopIndex(Layer) + 1

                    If AnimInstance(Index).LoopIndex(Layer) > Animation(AnimInstance(Index).Animation).LoopCount(Layer) Then
                        AnimInstance(Index).Used(Layer) = False
                    Else
                        AnimInstance(Index).frameIndex(Layer) = 1
                    End If

                Else
                    AnimInstance(Index).frameIndex(Layer) = AnimInstance(Index).frameIndex(Layer) + 1
                End If

                AnimInstance(Index).timer(Layer) = getTime
            End If
        End If

    Next

    ' if neither layer is used, clear
    If AnimInstance(Index).Used(0) = False And AnimInstance(Index).Used(1) = False Then ClearAnimInstance (Index)
End Sub

Public Function GetBankItemNum(ByVal BankSlot As Long) As Long

    If BankSlot = 0 Then
        GetBankItemNum = 0
        Exit Function
    End If

    If BankSlot > MAX_BANK Then
        GetBankItemNum = 0
        Exit Function
    End If

    GetBankItemNum = Bank.Item(BankSlot).Num
End Function

Public Sub SetBankItemNum(ByVal BankSlot As Long, ByVal ItemNum As Long)
    Bank.Item(BankSlot).Num = ItemNum
End Sub

Public Function GetBankItemValue(ByVal BankSlot As Long) As Long
    GetBankItemValue = Bank.Item(BankSlot).Value
End Function

Public Sub SetBankItemValue(ByVal BankSlot As Long, ByVal ItemValue As Long)
    Bank.Item(BankSlot).Value = ItemValue
End Sub

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockvar As Byte, ByRef Dir As Byte, ByVal block As Boolean)

    If block Then
        blockvar = blockvar Or (2 ^ Dir)
    Else
        blockvar = blockvar And Not (2 ^ Dir)
    End If

End Sub

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean

    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If

End Function

Public Sub PlayMapSound(ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)

    Dim soundName As String

    If entityNum <= 0 Then Exit Sub

    ' find the sound
    Select Case entityType

            ' animations
        Case SoundEntity.seAnimation

            If entityNum > MAX_ANIMATIONS Then Exit Sub
            soundName = Trim$(Animation(entityNum).sound)

            ' items
        Case SoundEntity.seItem

            If entityNum > MAX_ITEMS Then Exit Sub
            soundName = Trim$(Item(entityNum).sound)

            ' npcs
        Case SoundEntity.seNpc

            If entityNum > MAX_NPCS Then Exit Sub
            soundName = Trim$(Npc(entityNum).sound)

            ' resources
        Case SoundEntity.seResource

            If entityNum > MAX_RESOURCES Then Exit Sub
            soundName = Trim$(Resource(entityNum).sound)

            ' spells
        Case SoundEntity.seSpell

            If entityNum > MAX_SPELLS Then Exit Sub
            soundName = Trim$(Spell(entityNum).sound)

            ' other
        Case Else
            Exit Sub
    End Select

    ' exit out if it's not set
    If Trim$(soundName) = "None." Then Exit Sub

    ' play the sound
    If X > 0 And Y > 0 Then Play_Sound soundName, X, Y
End Sub

Public Sub CloseDialogue()
    diaIndex = 0
    HideWindow GetWindowIndex("winBlank")
    HideWindow GetWindowIndex("winDialogue")
End Sub

Public Sub Dialogue(ByVal header As String, ByVal body As String, ByVal body2 As String, ByVal Index As Long, Optional ByVal Style As Byte = 1, Optional ByVal Data1 As Long = 0)

    ' exit out if we've already got a dialogue open
    If diaIndex > 0 Then Exit Sub
    
    ' set buttons
    With Windows(GetWindowIndex("winDialogue"))
        If Style = StyleYESNO Then
           .Controls(GetControlIndex("winDialogue", "btnYes")).visible = True
            .Controls(GetControlIndex("winDialogue", "btnNo")).visible = True
            .Controls(GetControlIndex("winDialogue", "btnOkay")).visible = False
            .Controls(GetControlIndex("winDialogue", "txtInput")).visible = False
            .Controls(GetControlIndex("winDialogue", "lblBody_2")).visible = True
        ElseIf Style = StyleOKAY Then
            .Controls(GetControlIndex("winDialogue", "btnYes")).visible = False
            .Controls(GetControlIndex("winDialogue", "btnNo")).visible = False
            .Controls(GetControlIndex("winDialogue", "btnOkay")).visible = True
            .Controls(GetControlIndex("winDialogue", "txtInput")).visible = False
            .Controls(GetControlIndex("winDialogue", "lblBody_2")).visible = True
        ElseIf Style = StyleINPUT Then
            .Controls(GetControlIndex("winDialogue", "btnYes")).visible = False
            .Controls(GetControlIndex("winDialogue", "btnNo")).visible = False
            .Controls(GetControlIndex("winDialogue", "btnOkay")).visible = True
            .Controls(GetControlIndex("winDialogue", "txtInput")).visible = True
            .Controls(GetControlIndex("winDialogue", "lblBody_2")).visible = False
        End If
        
        ' set labels
        .Controls(GetControlIndex("winDialogue", "lblHeader")).text = header
        .Controls(GetControlIndex("winDialogue", "lblBody_1")).text = body
        .Controls(GetControlIndex("winDialogue", "lblBody_2")).text = body2
        .Controls(GetControlIndex("winDialogue", "txtInput")).text = vbNullString
    End With
    
    ' set it all up
    diaIndex = Index
    diaData1 = Data1
    diaStyle = Style
    
    ' make the windows visible
    ShowWindow GetWindowIndex("winBlank"), True
    ShowWindow GetWindowIndex("winDialogue"), True
End Sub

Public Sub dialogueHandler(ByVal Index As Long)
Dim Value As Long, diaInput As String

    Dim buffer As New clsBuffer
    Set buffer = New clsBuffer
    
    diaInput = Trim$(Windows(GetWindowIndex("winDialogue")).Controls(GetControlIndex("winDialogue", "txtInput")).text)

    ' find out which button
    If Index = 1 Then ' okay button

        ' dialogue index
        Select Case diaIndex
                Case TypeTRADEAMOUNT
                    Value = Val(diaInput)
                    TradeItem diaData1, Value
                Case TypeDEPOSITITEM
                    Value = Val(diaInput)
                    DepositItem diaData1, Value
                Case TypeWITHDRAWITEM
                    Value = Val(diaInput)
                    WithdrawItem diaData1, Value
                Case TypeDROPITEM
                    Value = Val(diaInput)
                    SendDropItem diaData1, Value
        End Select

    ElseIf Index = 2 Then ' yes button

        ' dialogue index
        Select Case diaIndex

            Case TypeTRADE
                SendAcceptTradeRequest

            Case TypeFORGET

                ForgetSpell diaData1

            Case TypePARTY
                SendAcceptParty

            Case TypeLOOTITEM
                ' send the packet
                Player(MyIndex).MapGetTimer = getTime
                buffer.WriteLong CMapGetItem
                SendData buffer.ToArray()

            Case TypeDELCHAR
                ' send the deletion
                SendDelChar diaData1
        End Select
    End If

    CloseDialogue
    diaIndex = 0
    diaInput = vbNullString
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long
    ConvertMapX = X - (TileView.Left * PIC_X) - Camera.Left
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    ConvertMapY = Y - (TileView.Top * PIC_Y) - Camera.Top

End Function

Public Sub UpdateCamera()
    Dim offsetX As Long, offsetY As Long, StartX As Long, StartY As Long, EndX As Long, EndY As Long
    
    offsetX = Player(MyIndex).xOffset + PIC_X
    offsetY = Player(MyIndex).yOffset + PIC_Y
    StartX = GetPlayerX(MyIndex) - ((TileWidth + 1) \ 2) - 1
    StartY = GetPlayerY(MyIndex) - ((TileHeight + 1) \ 2) - 1

    If TileWidth + 1 <= Map.MapData.MaxX Then
        If StartX < 0 Then
            offsetX = 0
    
            If StartX = -1 Then
                If Player(MyIndex).xOffset > 0 Then
                    offsetX = Player(MyIndex).xOffset
                End If
            End If
    
            StartX = 0
        End If
        
        EndX = StartX + (TileWidth + 1) + 1
        
        If EndX > Map.MapData.MaxX Then
            offsetX = 32
    
            If EndX = Map.MapData.MaxX + 1 Then
                If Player(MyIndex).xOffset < 0 Then
                    offsetX = Player(MyIndex).xOffset + PIC_X
                End If
            End If
    
            EndX = Map.MapData.MaxX
            StartX = EndX - TileWidth - 1
        End If
    Else
        EndX = StartX + (TileWidth + 1) + 1
    End If
    
    If TileHeight + 1 <= Map.MapData.MaxY Then
        If StartY < 0 Then
            offsetY = 0
    
            If StartY = -1 Then
                If Player(MyIndex).yOffset > 0 Then
                    offsetY = Player(MyIndex).yOffset
                End If
            End If
    
            StartY = 0
        End If
        
        EndY = StartY + (TileHeight + 1) + 1
        
        If EndY > Map.MapData.MaxY Then
            offsetY = 32
    
            If EndY = Map.MapData.MaxY + 1 Then
                If Player(MyIndex).yOffset < 0 Then
                    offsetY = Player(MyIndex).yOffset + PIC_Y
                End If
            End If
    
            EndY = Map.MapData.MaxY
            StartY = EndY - TileHeight - 1
        End If
    Else
        EndY = StartY + (TileHeight + 1) + 1
    End If
    
    If TileWidth + 1 = Map.MapData.MaxX Then
        offsetX = 0
    End If
    
    If TileHeight + 1 = Map.MapData.MaxY Then
        offsetY = 0
    End If

    With TileView
        .Top = StartY
        .Bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .Top = offsetY
        .Bottom = .Top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With

    CurX = TileView.Left + ((GlobalX + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((GlobalY + Camera.Top) \ PIC_Y)
    GlobalX_Map = GlobalX + (TileView.Left * PIC_X) + Camera.Left
    GlobalY_Map = GlobalY + (TileView.Top * PIC_Y) + Camera.Top
End Sub

Public Function CensorWord(ByVal sString As String) As String
    CensorWord = String$(Len(sString), "*")
End Function

Public Sub placeAutotile(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)

    With Autotile(X, Y).Layer(layernum).QuarterTile(tileQuarter)

        Select Case autoTileLetter

            Case "a"
                .X = autoInner(1).X
                .Y = autoInner(1).Y

            Case "b"
                .X = autoInner(2).X
                .Y = autoInner(2).Y

            Case "c"
                .X = autoInner(3).X
                .Y = autoInner(3).Y

            Case "d"
                .X = autoInner(4).X
                .Y = autoInner(4).Y

            Case "e"
                .X = autoNW(1).X
                .Y = autoNW(1).Y

            Case "f"
                .X = autoNW(2).X
                .Y = autoNW(2).Y

            Case "g"
                .X = autoNW(3).X
                .Y = autoNW(3).Y

            Case "h"
                .X = autoNW(4).X
                .Y = autoNW(4).Y

            Case "i"
                .X = autoNE(1).X
                .Y = autoNE(1).Y

            Case "j"
                .X = autoNE(2).X
                .Y = autoNE(2).Y

            Case "k"
                .X = autoNE(3).X
                .Y = autoNE(3).Y

            Case "l"
                .X = autoNE(4).X
                .Y = autoNE(4).Y

            Case "m"
                .X = autoSW(1).X
                .Y = autoSW(1).Y

            Case "n"
                .X = autoSW(2).X
                .Y = autoSW(2).Y

            Case "o"
                .X = autoSW(3).X
                .Y = autoSW(3).Y

            Case "p"
                .X = autoSW(4).X
                .Y = autoSW(4).Y

            Case "q"
                .X = autoSE(1).X
                .Y = autoSE(1).Y

            Case "r"
                .X = autoSE(2).X
                .Y = autoSE(2).Y

            Case "s"
                .X = autoSE(3).X
                .Y = autoSE(3).Y

            Case "t"
                .X = autoSE(4).X
                .Y = autoSE(4).Y
        End Select

    End With

End Sub

Public Sub initAutotiles()
    Dim X As Long, Y As Long, layernum As Long
    ' Procedure used to cache autotile positions. All positioning is
    ' independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    ' First, we need to re-size the array
    ReDim Autotile(0 To Map.MapData.MaxX, 0 To Map.MapData.MaxY)
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).X = 32
    autoInner(1).Y = 0
    ' NE - b
    autoInner(2).X = 48
    autoInner(2).Y = 0
    ' SW - c
    autoInner(3).X = 32
    autoInner(3).Y = 16
    ' SE - d
    autoInner(4).X = 48
    autoInner(4).Y = 16
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).X = 0
    autoNW(1).Y = 32
    ' NE - f
    autoNW(2).X = 16
    autoNW(2).Y = 32
    ' SW - g
    autoNW(3).X = 0
    autoNW(3).Y = 48
    ' SE - h
    autoNW(4).X = 16
    autoNW(4).Y = 48
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).X = 32
    autoNE(1).Y = 32
    ' NE - g
    autoNE(2).X = 48
    autoNE(2).Y = 32
    ' SW - k
    autoNE(3).X = 32
    autoNE(3).Y = 48
    ' SE - l
    autoNE(4).X = 48
    autoNE(4).Y = 48
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).X = 0
    autoSW(1).Y = 64
    ' NE - n
    autoSW(2).X = 16
    autoSW(2).Y = 64
    ' SW - o
    autoSW(3).X = 0
    autoSW(3).Y = 80
    ' SE - p
    autoSW(4).X = 16
    autoSW(4).Y = 80
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).X = 32
    autoSE(1).Y = 64
    ' NE - r
    autoSE(2).X = 48
    autoSE(2).Y = 64
    ' SW - s
    autoSE(3).X = 32
    autoSE(3).Y = 80
    ' SE - t
    autoSE(4).X = 48
    autoSE(4).Y = 80

    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            For layernum = 1 To MapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                calculateAutotile X, Y, layernum
                ' cache the rendering state of the tiles and set them
                cacheRenderState X, Y, layernum
            Next
        Next
    Next

End Sub

Public Sub cacheRenderState(ByVal X As Long, ByVal Y As Long, ByVal layernum As Long)
    Dim quarterNum As Long

    ' exit out early
    If X < 0 Or X > Map.MapData.MaxX Or Y < 0 Or Y > Map.MapData.MaxY Then Exit Sub

    With Map.TileData.Tile(X, Y)

        ' check if the tile can be rendered
        If .Layer(layernum).tileSet <= 0 Or .Layer(layernum).tileSet > CountTileset Then
            Autotile(X, Y).Layer(layernum).RenderState = RENDER_STATE_NONE
            Exit Sub
        End If
        
        ' check if we're a bottom
        If layernum = MapLayer.Ground Then
            ' check if bottom
            If Y > 0 Then
                If Map.TileData.Tile(X, Y - 1).Type = TILE_TYPE_APPEAR Then
                    If Map.TileData.Tile(X, Y - 1).Data2 Then
                        Autotile(X, Y).Layer(layernum).RenderState = RENDER_STATE_APPEAR
                        Exit Sub
                    End If
                End If
            End If
        End If

        ' check if it's a key - hide mask if key is closed
        If layernum = MapLayer.Mask Then
            If .Type = TILE_TYPE_KEY Then
                If TempTile(X, Y).DoorOpen = 0 Then
                    Autotile(X, Y).Layer(layernum).RenderState = RENDER_STATE_NONE
                    Exit Sub
                End If
            End If
            If .Type = TILE_TYPE_APPEAR Then
                Autotile(X, Y).Layer(layernum).RenderState = RENDER_STATE_APPEAR
                Exit Sub
            End If
        End If

        ' check if it needs to be rendered as an autotile
        If .Autotile(layernum) = AUTOTILE_NONE Or .Autotile(layernum) = AUTOTILE_FAKE Or Options.NoAuto = 1 Then
            ' default to... default
            Autotile(X, Y).Layer(layernum).RenderState = RENDER_STATE_normal
        Else
            Autotile(X, Y).Layer(layernum).RenderState = RENDER_STATE_AUTOTILE

            ' cache tileset positioning
            For quarterNum = 1 To 4
                Autotile(X, Y).Layer(layernum).srcX(quarterNum) = (Map.TileData.Tile(X, Y).Layer(layernum).X * 32) + Autotile(X, Y).Layer(layernum).QuarterTile(quarterNum).X
                Autotile(X, Y).Layer(layernum).srcY(quarterNum) = (Map.TileData.Tile(X, Y).Layer(layernum).Y * 32) + Autotile(X, Y).Layer(layernum).QuarterTile(quarterNum).Y
            Next

        End If

    End With

End Sub

Public Sub calculateAutotile(ByVal X As Long, ByVal Y As Long, ByVal layernum As Long)

    ' Right, so we've split the tile block in to an easy to remember
    ' collection of letters. We now need to do the calculations to find
    ' out which little lettered block needs to be rendered. We do this
    ' by reading the surrounding tiles to check for matches.
    ' First we check to make sure an autotile situation is actually there.
    ' Then we calculate exactly which situation has arisen.
    ' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
    ' Exit out if we don't have an auatotile
    If Map.TileData.Tile(X, Y).Autotile(layernum) = 0 Then Exit Sub

    ' Okay, we have autotiling but which one?
    Select Case Map.TileData.Tile(X, Y).Autotile(layernum)

            ' normal or animated - same difference
        Case AUTOTILE_normal, AUTOTILE_ANIM
            ' North West Quarter
            CalculateNW_normal layernum, X, Y
            ' North East Quarter
            CalculateNE_normal layernum, X, Y
            ' South West Quarter
            CalculateSW_normal layernum, X, Y
            ' South East Quarter
            CalculateSE_normal layernum, X, Y

            ' Cliff
        Case AUTOTILE_CLIFF
            ' North West Quarter
            CalculateNW_Cliff layernum, X, Y
            ' North East Quarter
            CalculateNE_Cliff layernum, X, Y
            ' South West Quarter
            CalculateSW_Cliff layernum, X, Y
            ' South East Quarter
            CalculateSE_Cliff layernum, X, Y

            ' Waterfalls
        Case AUTOTILE_WATERFALL
            ' North West Quarter
            CalculateNW_Waterfall layernum, X, Y
            ' North East Quarter
            CalculateNE_Waterfall layernum, X, Y
            ' South West Quarter
            CalculateSW_Waterfall layernum, X, Y
            ' South East Quarter
            CalculateSE_Waterfall layernum, X, Y

            ' Anything else
        Case Else
            ' Don't need to render anything... it's fake or not an autotile
    End Select

End Sub

' normal autotiling
Public Sub CalculateNW_normal(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North West
    If checkTileMatch(layernum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True

    ' North
    If checkTileMatch(layernum, X, Y, X, Y - 1) Then tmpTile(2) = True

    ' West
    If checkTileMatch(layernum, X, Y, X - 1, Y) Then tmpTile(3) = True

    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, X, Y, 1, "e"

        Case AUTO_OUTER
            placeAutotile layernum, X, Y, 1, "a"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, X, Y, 1, "i"

        Case AUTO_VERTICAL
            placeAutotile layernum, X, Y, 1, "m"

        Case AUTO_FILL
            placeAutotile layernum, X, Y, 1, "q"
    End Select

End Sub

Public Sub CalculateNE_normal(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North

    If checkTileMatch(layernum, X, Y, X, Y - 1) Then tmpTile(1) = True

    ' North East
    If checkTileMatch(layernum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True

    ' East
    If checkTileMatch(layernum, X, Y, X + 1, Y) Then tmpTile(3) = True

    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, X, Y, 2, "j"

        Case AUTO_OUTER
            placeAutotile layernum, X, Y, 2, "b"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, X, Y, 2, "f"

        Case AUTO_VERTICAL
            placeAutotile layernum, X, Y, 2, "r"

        Case AUTO_FILL
            placeAutotile layernum, X, Y, 2, "n"
    End Select

End Sub

Public Sub CalculateSW_normal(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)

    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' West
    If checkTileMatch(layernum, X, Y, X - 1, Y) Then tmpTile(1) = True

    ' South West
    If checkTileMatch(layernum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True

    ' South
    If checkTileMatch(layernum, X, Y, X, Y + 1) Then tmpTile(3) = True

    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL

    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, X, Y, 3, "o"

        Case AUTO_OUTER
            placeAutotile layernum, X, Y, 3, "c"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, X, Y, 3, "s"

        Case AUTO_VERTICAL
            placeAutotile layernum, X, Y, 3, "g"

        Case AUTO_FILL
            placeAutotile layernum, X, Y, 3, "k"
    End Select

End Sub

Public Sub CalculateSE_normal(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)

    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' South
    If checkTileMatch(layernum, X, Y, X, Y + 1) Then tmpTile(1) = True

    ' South East
    If checkTileMatch(layernum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True

    ' East
    If checkTileMatch(layernum, X, Y, X + 1, Y) Then tmpTile(3) = True

    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, X, Y, 4, "t"

        Case AUTO_OUTER
            placeAutotile layernum, X, Y, 4, "d"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, X, Y, 4, "p"

        Case AUTO_VERTICAL
            placeAutotile layernum, X, Y, 4, "l"

        Case AUTO_FILL
            placeAutotile layernum, X, Y, 4, "h"
    End Select

End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile As Boolean

    ' West
    If checkTileMatch(layernum, X, Y, X - 1, Y) Then tmpTile = True

    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layernum, X, Y, 1, "i"
    Else
        ' Edge
        placeAutotile layernum, X, Y, 1, "e"
    End If

End Sub

Public Sub CalculateNE_Waterfall(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile As Boolean

    ' East
    If checkTileMatch(layernum, X, Y, X + 1, Y) Then tmpTile = True

    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layernum, X, Y, 2, "f"
    Else
        ' Edge
        placeAutotile layernum, X, Y, 2, "j"
    End If

End Sub

Public Sub CalculateSW_Waterfall(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile As Boolean

    ' West
    If checkTileMatch(layernum, X, Y, X - 1, Y) Then tmpTile = True

    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layernum, X, Y, 3, "k"
    Else
        ' Edge
        placeAutotile layernum, X, Y, 3, "g"
    End If

End Sub

Public Sub CalculateSE_Waterfall(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile As Boolean

    ' East
    If checkTileMatch(layernum, X, Y, X + 1, Y) Then tmpTile = True

    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layernum, X, Y, 4, "h"
    Else
        ' Edge
        placeAutotile layernum, X, Y, 4, "l"
    End If

End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North West
    If checkTileMatch(layernum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True

    ' North
    If checkTileMatch(layernum, X, Y, X, Y - 1) Then tmpTile(2) = True

    ' West
    If checkTileMatch(layernum, X, Y, X - 1, Y) Then tmpTile(3) = True

    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, X, Y, 1, "e"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, X, Y, 1, "i"

        Case AUTO_VERTICAL
            placeAutotile layernum, X, Y, 1, "m"

        Case AUTO_FILL
            placeAutotile layernum, X, Y, 1, "q"
    End Select

End Sub

Public Sub CalculateNE_Cliff(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North
    If checkTileMatch(layernum, X, Y, X, Y - 1) Then tmpTile(1) = True

    ' North East
    If checkTileMatch(layernum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True

    ' East
    If checkTileMatch(layernum, X, Y, X + 1, Y) Then tmpTile(3) = True

    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, X, Y, 2, "j"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, X, Y, 2, "f"

        Case AUTO_VERTICAL
            placeAutotile layernum, X, Y, 2, "r"

        Case AUTO_FILL
            placeAutotile layernum, X, Y, 2, "n"
    End Select

End Sub

Public Sub CalculateSW_Cliff(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' West
    If checkTileMatch(layernum, X, Y, X - 1, Y) Then tmpTile(1) = True

    ' South West
    If checkTileMatch(layernum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True

    ' South
    If checkTileMatch(layernum, X, Y, X, Y + 1) Then tmpTile(3) = True

    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, X, Y, 3, "o"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, X, Y, 3, "s"

        Case AUTO_VERTICAL
            placeAutotile layernum, X, Y, 3, "g"

        Case AUTO_FILL
            placeAutotile layernum, X, Y, 3, "k"
    End Select

End Sub

Public Sub CalculateSE_Cliff(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' South
    If checkTileMatch(layernum, X, Y, X, Y + 1) Then tmpTile(1) = True

    ' South East
    If checkTileMatch(layernum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True

    ' East
    If checkTileMatch(layernum, X, Y, X + 1, Y) Then tmpTile(3) = True

    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, X, Y, 4, "t"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, X, Y, 4, "p"

        Case AUTO_VERTICAL
            placeAutotile layernum, X, Y, 4, "l"

        Case AUTO_FILL
            placeAutotile layernum, X, Y, 4, "h"
    End Select

End Sub

Public Function checkTileMatch(ByVal layernum As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    ' we'll exit out early if true
    checkTileMatch = True

    ' if it's off the map then set it as autotile and exit out early
    If X2 < 0 Or X2 > Map.MapData.MaxX Or Y2 < 0 Or Y2 > Map.MapData.MaxY Then
        checkTileMatch = True
        Exit Function
    End If

    ' fakes ALWAYS return true
    If Map.TileData.Tile(X2, Y2).Autotile(layernum) = AUTOTILE_FAKE Then
        checkTileMatch = True
        Exit Function
    End If

    ' check neighbour is an autotile
    If Map.TileData.Tile(X2, Y2).Autotile(layernum) = 0 Then
        checkTileMatch = False
        Exit Function
    End If

    ' check we're a matching
    If Map.TileData.Tile(X1, Y1).Layer(layernum).tileSet <> Map.TileData.Tile(X2, Y2).Layer(layernum).tileSet Then
        checkTileMatch = False
        Exit Function
    End If

    ' check tiles match
    If Map.TileData.Tile(X1, Y1).Layer(layernum).X <> Map.TileData.Tile(X2, Y2).Layer(layernum).X Then
        checkTileMatch = False
        Exit Function
    End If

    If Map.TileData.Tile(X1, Y1).Layer(layernum).Y <> Map.TileData.Tile(X2, Y2).Layer(layernum).Y Then
        checkTileMatch = False
        Exit Function
    End If

End Function

Public Sub OpenNpcChat(ByVal NpcNum As Long, ByVal mT As String, ByRef o() As String)
Dim i As Long, X As Long

    ' find out how many options we have
    convOptions = 0
    For i = 1 To 4
        If Len(o(i)) > 0 Then convOptions = convOptions + 1
    Next
    
    ' gui stuff
    With Windows(GetWindowIndex("winNpcChat"))
        ' set main text
        .Window.text = "Conversation with " & Trim$(Npc(NpcNum).Name)
        .Controls(GetControlIndex("winNpcChat", "lblChat")).text = mT
        ' make everything visible
        For i = 1 To 4
            .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).Top = optPos(i)
            .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).visible = True
        Next
        ' set sizes
        .Window.Height = optHeight
        .Controls(GetControlIndex("winNpcChat", "picParchment")).Height = .Window.Height - 30
        ' move options depending on count
        If convOptions < 4 Then
            For i = convOptions + 1 To 4
                .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).Top = optPos(i)
                .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).visible = False
            Next
            For i = 1 To convOptions
                .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).Top = optPos(i + (4 - convOptions))
                .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).visible = True
            Next
            .Window.Height = optHeight - ((4 - convOptions) * 18)
            .Controls(GetControlIndex("winNpcChat", "picParchment")).Height = .Window.Height - 32
        End If
        ' set labels
        X = convOptions
        For i = 1 To 4
            .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).text = X & ". " & o(i)
            X = X - 1
        Next
        For i = 0 To 5
            .Controls(GetControlIndex("winNpcChat", "picFace")).image(i) = TextureFace(Npc(NpcNum).sprite)
        Next
    End With
    
    ' we're in chat now boy
    inChat = True
    
    ' show the window
    ShowWindow GetWindowIndex("winNpcChat")
End Sub

Public Sub SetTutorialState(ByVal stateNum As Byte)
    Dim i As Long

    Select Case stateNum

        Case 1 ' introduction
            chatText = "Ah, so you have appeared at last my dear. Please, listen to what I have to say."
            chatOpt(1) = "*sigh* I suppose I should..."

            For i = 2 To 4
                chatOpt(i) = vbNullString
            Next

        Case 2 ' next
            chatText = "There are some important things you need to know. Here they are. To move, use W, A, S and D. To attack or to talk to someone, press CTRL. To initiate chat press ENTER."
            chatOpt(1) = "Go on..."

            For i = 2 To 4
                chatOpt(i) = vbNullString
            Next

        Case 3 ' chatting
            chatText = "When chatting you can talk in different channels. By default you're talking in the map channel. To talk globally append an apostrophe (') to the start of your message. To perform an emote append a hyphen (-) to the start of your message."
            chatOpt(1) = "Wait, what about combat?"

            For i = 2 To 4
                chatOpt(i) = vbNullString
            Next

        Case 4 ' combat
            chatText = "Combat can be done through melee and skills. You can melee an enemy by facing them and pressing CTRL. To use a skill you can double click it in your skill menu, double click it in the hotbar or use the number keys. (1, 2, 3, etc.)"
            chatOpt(1) = "Oh! What do stats do?"

            For i = 2 To 4
                chatOpt(i) = vbNullString
            Next

        Case 5 ' stats
            chatText = "Strength increases damage and allows you to equip better weaponry. Endurance increases your maximum health. Intelligence increases your maximum spirit. Agility allows you to reduce damage received and also increases critical hit chances. Willpower increase regeneration abilities."
            chatOpt(1) = "Thanks. See you later."

            For i = 2 To 4
                chatOpt(i) = vbNullString
            Next

        Case Else ' goodbye
            chatText = vbNullString

            For i = 1 To 4
                chatOpt(i) = vbNullString
            Next

            SendFinishTutorial
            inTutorial = False
            AddText "Well done, you finished the tutorial.", BrightGreen
            Exit Sub
    End Select

    ' set the state
    tutorialState = stateNum
End Sub

Public Sub ScrollChatBox(ByVal Direction As Byte)
    If Direction = 0 Then ' up
        If ChatScroll < ChatLines Then
            ChatScroll = ChatScroll + 1
        End If
    Else
        If ChatScroll > 0 Then
            ChatScroll = ChatScroll - 1
        End If
    End If
End Sub

Public Sub ClearMapCache()
    Dim i As Long, FileName As String

    For i = 1 To MAX_MAPS
        FileName = App.Path & "\data files\maps\map" & i & ".map"

        If FileExist(FileName) Then
            Kill FileName
        End If

    Next

    AddText "Map cache destroyed.", BrightGreen
End Sub

Public Sub AddChatBubble(ByVal target As Long, ByVal TargetType As Byte, ByVal Msg As String, ByVal Colour As Long)
    Dim i As Long, Index As Long
    ' set the global index
    chatBubbleIndex = chatBubbleIndex + 1
    
    ' reset to yourself for eventing
    If TargetType = 0 Then
        TargetType = TARGET_TYPE_PLAYER
        If target = 0 Then target = MyIndex
    End If

    If chatBubbleIndex < 1 Or chatBubbleIndex > MAX_BYTE Then chatBubbleIndex = 1
    ' default to new bubble
    Index = chatBubbleIndex

    ' loop through and see if that player/npc already has a chat bubble
    For i = 1 To MAX_BYTE
        If chatBubble(i).TargetType = TargetType Then
            If chatBubble(i).target = target Then
                ' reset master index
                If chatBubbleIndex > 1 Then chatBubbleIndex = chatBubbleIndex - 1
                ' we use this one now, yes?
                Index = i
                Exit For
            End If
        End If
    Next

    ' set the bubble up
    With chatBubble(Index)
        .target = target
        .TargetType = TargetType
        .Msg = Msg
        .Colour = Colour
        .timer = getTime
        .active = True
    End With
End Sub

Public Sub FindNearestTarget()
    Dim i As Long, X As Long, Y As Long, X2 As Long, Y2 As Long, xDif As Long, yDif As Long
    Dim bestX As Long, bestY As Long, bestIndex As Long
    X2 = GetPlayerX(MyIndex)
    Y2 = GetPlayerY(MyIndex)
    bestX = 255
    bestY = 255

    For i = 1 To MAX_MAP_NPCS

        If MapNpc(i).Num > 0 Then
            X = MapNpc(i).X
            Y = MapNpc(i).Y

            ' find the difference - x
            If X < X2 Then
                xDif = X2 - X
            ElseIf X > X2 Then
                xDif = X - X2
            Else
                xDif = 0
            End If

            ' find the difference - y
            If Y < Y2 Then
                yDif = Y2 - Y
            ElseIf Y > Y2 Then
                yDif = Y - Y2
            Else
                yDif = 0
            End If

            ' best so far?
            If (xDif + yDif) < (bestX + bestY) Then
                bestX = xDif
                bestY = yDif
                bestIndex = i
            End If
        End If

    Next

    ' target the best
    If bestIndex > 0 And bestIndex <> myTarget Then PlayerTarget bestIndex, TARGET_TYPE_NPC
End Sub

Public Sub FindTarget()
    Dim i As Long, X As Long, Y As Long

    ' check players
    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(MyIndex) = GetPlayerMap(i) Then
            X = (GetPlayerX(i) * 32) + Player(i).xOffset + 32
            Y = (GetPlayerY(i) * 32) + Player(i).yOffset + 32

            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    ' found our target!
                    PlayerTarget i, TARGET_TYPE_PLAYER
                    Exit Sub
                End If
            End If
        End If

    Next

    ' check npcs
    For i = 1 To MAX_MAP_NPCS

        If MapNpc(i).Num > 0 Then
            X = (MapNpc(i).X * 32) + MapNpc(i).xOffset + 32
            Y = (MapNpc(i).Y * 32) + MapNpc(i).yOffset + 32

            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    ' found our target!
                    PlayerTarget i, TARGET_TYPE_NPC
                    Exit Sub
                End If
            End If
        End If

    Next

End Sub

Public Sub SetBarWidth(ByRef MaxWidth As Long, ByRef Width As Long)
    Dim barDifference As Long

    If MaxWidth < Width Then
        ' find out the amount to increase per loop
        barDifference = ((Width - MaxWidth) / 100) * 10

        ' if it's less than 1 then default to 1
        If barDifference < 1 Then barDifference = 1
        ' set the width
        Width = Width - barDifference
    ElseIf MaxWidth > Width Then
        ' find out the amount to increase per loop
        barDifference = ((MaxWidth - Width) / 100) * 10

        ' if it's less than 1 then default to 1
        If barDifference < 1 Then barDifference = 1
        ' set the width
        Width = Width + barDifference
    End If

End Sub

Public Sub DialogueAlert(ByVal Index As Long)
    Dim header As String, body As String, body2 As String

    ' find the body/header
    Select Case Index

        Case MsgCONNECTION
            header = "Connection Problem"
            body = "You lost connection to the server."
            body2 = "Please try again later."

        Case MsgBANNED
            header = "Banned"
            body = "You have been banned from playing Crystalshire."
            body2 = "Please send all ban appeals to an administrator."

        Case MsgKICKED
            header = "Kicked"
            body = "You have been kicked from Crystalshire."
            body2 = "Please try and behave."

        Case MsgOUTDATED
            header = "Wrong Version"
            body = "Your game client is the wrong version."
            body2 = "Please re-load the game or wait for a patch."

        Case MsgUSERLENGTH
            header = "Invalid Length"
            body = "Your username or password is too short or too long."
            body2 = "Please enter a valid username and password."

        Case MsgILLEGALNAME
            header = "Illegal Characters"
            body = "Your username or password contains illegal characters."
            body2 = "Please enter a valid username and password."

        Case MsgREBOOTING
            header = "Connection Refused"
            body = "The server is currently rebooting."
            body2 = "Please try again soon."

        Case MsgNAMETAKEN
            header = "Invalid Name"
            body = "This name is already in use."
            body2 = "Please try another name."

        Case MsgNAMELENGTH
            header = "Invalid Name"
            body = "This name is too short or too long."
            body2 = "Please try another name."

        Case MsgNAMEILLEGAL
            header = "Invalid Name"
            body = "This name contains illegal characters."
            body2 = "Please try another name."

        Case MsgMYSQL
            header = "Connection Problem"
            body = "Cannot connect to database."
            body2 = "Please try again later."

        Case MsgWRONGPASS
            header = "Invalid Login"
            body = "Invalid username or password."
            body2 = "Please try again."

        Case MsgACTIVATED
            header = "Inactive Account"
            body = "Your account is not activated."
            body2 = "Please activate your account then try again."

        Case MsgMERGE
            header = "Successful Merge"
            body = "Character merged with new account."
            body2 = "Old account permanently destroyed."

        Case MsgMAXCHARS
            header = "Cannot Merge"
            body = "You cannot merge a full account."
            body2 = "Please clear a character slot."

        Case MsgMERGENAME
            header = "Cannot Merge"
            body = "An existing character has this name."
            body2 = "Please contact an administrator."
            
        Case MsgDELCHAR
            header = "Deleted Character"
            body = "Your character was successfully deleted."
            body2 = "Please log on to continue playing."
        Case MsgCreated
            header = "Account Created"
            body = "Your Account was successfully created."
            body2 = "Now, you can play!"
    End Select

    ' set the dialogue up!
    Dialogue header, body, body2, TypeALERT
End Sub

Public Function hasProficiency(ByVal Index As Long, ByVal proficiency As Long) As Boolean

    Select Case proficiency

        Case 0 ' None
            hasProficiency = True
            Exit Function

        Case 1 ' Heavy

            If GetPlayerClass(Index) = 1 Then
                hasProficiency = True
                Exit Function
            End If

        Case 2 ' Light

            If GetPlayerClass(Index) = 2 Or GetPlayerClass(Index) = 3 Then
                hasProficiency = True
                Exit Function
            End If

    End Select

    hasProficiency = False
End Function

Public Function Clamp(ByVal Value As Long, ByVal min As Long, ByVal max As Long) As Long
    Clamp = Value

    If Value < min Then Clamp = min
    If Value > max Then Clamp = max
End Function

Public Sub ShowClasses()
    HideWindows
    newCharClass = 1
    newCharSprite = 1
    newCharGender = SEX_MALE
    Windows(GetWindowIndex("winClasses")).Controls(GetControlIndex("winClasses", "lblClassName")).text = Trim$(Class(newCharClass).Name)
    Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "txtName")).text = vbNullString
    Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "chkMale")).Value = 1
    Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "chkFemale")).Value = 0
    ShowWindow GetWindowIndex("winClasses")
End Sub

Public Sub SetGoldLabel()
Dim i As Long, Amount As Long
    Amount = 0
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) = 1 Then
            Amount = GetPlayerInvItemValue(MyIndex, i)
        End If
    Next
    Windows(GetWindowIndex("winShop")).Controls(GetControlIndex("winShop", "lblGold")).text = Format$(Amount, "#,###,###,###") & "g"
    Windows(GetWindowIndex("winInventory")).Controls(GetControlIndex("winInventory", "lblGold")).text = Format$(Amount, "#,###,###,###") & "g"
End Sub

Public Sub ShowInvDesc(X As Long, Y As Long, invNum As Long)
    Dim soulBound As Boolean

    ' rte9
    If invNum <= 0 Or invNum > MAX_INV Then Exit Sub
    
    ' show
    If GetPlayerInvItemNum(MyIndex, invNum) Then
        If Item(GetPlayerInvItemNum(MyIndex, invNum)).BindType > 0 And PlayerInv(invNum).bound > 0 Then soulBound = True
        ShowItemDesc X, Y, GetPlayerInvItemNum(MyIndex, invNum), soulBound
    End If
End Sub

Public Sub ShowShopDesc(X As Long, Y As Long, ItemNum As Long)
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    ' show
    ShowItemDesc X, Y, ItemNum, False
End Sub

Public Sub ShowEqDesc(X As Long, Y As Long, eqNum As Long)
    Dim soulBound As Boolean

    ' rte9
    If eqNum <= 0 Or eqNum > Equipment.Equipment_Count - 1 Then Exit Sub
    
    ' show
    If Player(MyIndex).Equipment(eqNum) Then
        If Item(Player(MyIndex).Equipment(eqNum)).BindType > 0 Then soulBound = True
        ShowItemDesc X, Y, Player(MyIndex).Equipment(eqNum), soulBound
    End If
End Sub

Public Sub ShowOfferDesc(X As Long, Y As Long, OfferNum As Long)
    Dim Colour As Long, className As String, levelTxt As String, i As Long

    If inOffer(OfferNum) < 0 Then Exit Sub
    ' set globals
    descType = 3 ' offer
    descItem = OfferNum
    
    ' set position
    Windows(GetWindowIndex("winDescription")).Window.Left = X
    Windows(GetWindowIndex("winDescription")).Window.Top = Y
    
    ' show the window
    ShowWindow GetWindowIndex("winDescription"), , False
    
    ' exit out early if last is same
    If (descLastType = descType) And (descLastItem = descItem) Then Exit Sub
    
    ' set last to this
    descLastType = descType
    descLastItem = descItem
    
    ' clear
    ReDim descText(1 To 1) As TextColourRec
    
    ' show req. labels
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "lblClass")).visible = True
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "lblLevel")).visible = True
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "lblDescription")).visible = True
    
    ' set variables
    Select Case inOfferType(OfferNum)
        Case Offers.Offer_Type_Mission
            With Windows(GetWindowIndex("winDescription"))
                .Controls(GetControlIndex("winDescription", "lblName")).text = Trim$(Mission(inOffer(OfferNum)).Name)
                .Controls(GetControlIndex("winDescription", "lblName")).textColour = White
                
                Select Case Mission(inOffer(OfferNum)).Type
                    Case MissionType.Mission_TypeCollect
                        If Mission(inOffer(OfferNum)).CollectItem = 0 Then Exit Sub
                        .Controls(GetControlIndex("winDescription", "lblClass")).text = "Collect: " + Trim$(Item(Mission(inOffer(OfferNum)).CollectItem).Name)
                        .Controls(GetControlIndex("winDescription", "lblLevel")).text = "Amount: " + Trim$(Mission(inOffer(OfferNum)).CollectItemAmount)
                        .Controls(GetControlIndex("winDescription", "lblLevel")).textColour = Red
                        .Controls(GetControlIndex("winDescription", "lblDescription")).text = Trim$(Mission(inOffer(OfferNum)).Description)
                    Case MissionType.Mission_TypeKill
                        If Mission(inOffer(OfferNum)).KillNPC = 0 Then Exit Sub
                        .Controls(GetControlIndex("winDescription", "lblClass")).text = "Kill: " + Trim$(Item(Mission(inOffer(OfferNum)).KillNPC).Name)
                        .Controls(GetControlIndex("winDescription", "lblLevel")).text = "Amount: " + Trim$(Mission(inOffer(OfferNum)).KillNPCAmount)
                        .Controls(GetControlIndex("winDescription", "lblLevel")).textColour = Red
                        .Controls(GetControlIndex("winDescription", "lblDescription")).text = Trim$(Mission(inOffer(OfferNum)).Description)
                    Case MissionType.Mission_TypeTalk
                        If Mission(inOffer(OfferNum)).TalkNPC = 0 Then Exit Sub
                        .Controls(GetControlIndex("winDescription", "lblClass")).text = "Talk: " + Trim$(Npc(Mission(inOffer(OfferNum)).TalkNPC).Name)
                        .Controls(GetControlIndex("winDescription", "lblDescription")).text = Trim$(Mission(inOffer(OfferNum)).Description)
                End Select
                
                '.Controls(GetControlIndex("winDescription", "lblLevel")).textColour = Colour
            End With
    End Select
End Sub

Public Sub ShowPlayerSpellDesc(X As Long, Y As Long, slotNum As Long)
    
    ' rte9
    If slotNum <= 0 Or slotNum > MAX_PLAYER_SPELLS Then Exit Sub
    
    ' show
    If PlayerSpells(slotNum).Spell Then
        ShowSpellDesc X, Y, PlayerSpells(slotNum).Spell, slotNum
    End If
End Sub

Public Sub ShowSpellDesc(X As Long, Y As Long, spellnum As Long, spellSlot As Long)
Dim Colour As Long, theName As String, sUse As String, i As Long, barWidth As Long, tmpWidth As Long

    ' set globals
    descType = 2 ' spell
    descItem = spellnum
    
    ' set position
    Windows(GetWindowIndex("winDescription")).Window.Left = X
    Windows(GetWindowIndex("winDescription")).Window.Top = Y
    
    ' show the window
    ShowWindow GetWindowIndex("winDescription"), , False
    
    ' exit out early if last is same
    If (descLastType = descType) And (descLastItem = descItem) Then Exit Sub
    
    ' clear
    ReDim descText(1 To 1) As TextColourRec
    
    ' hide req. labels
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "lblLevel")).visible = False
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "picBar")).visible = True
    
    ' set variables
    With Windows(GetWindowIndex("winDescription"))
        ' set name
        .Controls(GetControlIndex("winDescription", "lblName")).text = Trim$(Spell(spellnum).Name)
        .Controls(GetControlIndex("winDescription", "lblName")).textColour = White
        
        ' find ranks
        If spellSlot > 0 Then
            ' draw the rank bar
            barWidth = 66
            If Spell(spellnum).NextRank > 0 Then
                tmpWidth = ((PlayerSpells(spellSlot).Uses / barWidth) / (Spell(spellnum).NextUses / barWidth)) * barWidth
            Else
                tmpWidth = 66
            End If
            .Controls(GetControlIndex("winDescription", "picBar")).Value = tmpWidth
            ' does it rank up?
            If Spell(spellnum).NextRank > 0 Then
                Colour = White
                sUse = "Uses: " & PlayerSpells(spellSlot).Uses & "/" & Spell(spellnum).NextUses
                If PlayerSpells(spellSlot).Uses = Spell(spellnum).NextUses Then
                    If Not GetPlayerLevel(MyIndex) >= Spell(Spell(spellnum).NextRank).LevelReq Then
                        Colour = BrightRed
                        sUse = "Lvl " & Spell(Spell(spellnum).NextRank).LevelReq & " req."
                    End If
                End If
            Else
                Colour = Grey
                sUse = "Max Rank"
            End If
            ' show controls
            .Controls(GetControlIndex("winDescription", "lblClass")).visible = True
            .Controls(GetControlIndex("winDescription", "picBar")).visible = True
             'set vals
            .Controls(GetControlIndex("winDescription", "lblClass")).text = sUse
            .Controls(GetControlIndex("winDescription", "lblClass")).textColour = Colour
        Else
            ' hide some controls
            .Controls(GetControlIndex("winDescription", "lblClass")).visible = False
            .Controls(GetControlIndex("winDescription", "picBar")).visible = False
        End If
    End With
    
    Select Case Spell(spellnum).Type
        Case SPELL_TYPE_DAMAGEHP
            AddDescInfo "Damage HP"
        Case SPELL_TYPE_DAMAGEMP
            AddDescInfo "Damage SP"
        Case SPELL_TYPE_HEALHP
            AddDescInfo "Heal HP"
        Case SPELL_TYPE_HEALMP
            AddDescInfo "Heal SP"
        Case SPELL_TYPE_WARP
            AddDescInfo "Warp"
    End Select
    
    ' more info
    Select Case Spell(spellnum).Type
        Case SPELL_TYPE_DAMAGEHP, SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP
            ' damage
            AddDescInfo "Vital: " & Spell(spellnum).Vital
            
            ' mp cost
            AddDescInfo "Cost: " & Spell(spellnum).MPCost & " SP"
            
            ' cast time
            AddDescInfo "Cast Time: " & Spell(spellnum).CastTime & "s"
            
            ' cd time
            AddDescInfo "Cooldown: " & Spell(spellnum).CDTime & "s"
            
            ' aoe
            If Spell(spellnum).RadiusX > 0 Then
                AddDescInfo "AoE: " & Spell(spellnum).RadiusX
            End If
            
            ' stun
            If Spell(spellnum).StunDuration > 0 Then
                AddDescInfo "Stun: " & Spell(spellnum).StunDuration & "s"
            End If
            
            ' dot
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Interval > 0 Then
                AddDescInfo "DoT: " & (Spell(spellnum).Duration / Spell(spellnum).Interval) & " tick"
            End If
    End Select
End Sub

Public Sub ResetControlsWinDesc()
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "lblClass")).text = ""
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "lblLevel")).text = ""
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "lblDescription")).text = ""
End Sub

Public Sub ShowItemDesc(X As Long, Y As Long, ItemNum As Long, soulBound As Boolean)
    Dim Colour As Long, theName As String, className As String, levelTxt As String, i As Long
    
    Call ResetControlsWinDesc
    ' set globals
    descType = 1 ' inventory
    descItem = ItemNum
    
    ' set position
    Windows(GetWindowIndex("winDescription")).Window.Left = X
    Windows(GetWindowIndex("winDescription")).Window.Top = Y
    
    ' show the window
    ShowWindow GetWindowIndex("winDescription"), , False
    
    ' exit out early if last is same
    If (descLastType = descType) And (descLastItem = descItem) Then Exit Sub
    
    ' set last to this
    descLastType = descType
    descLastItem = descItem
    
    ' show req. labels
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "lblClass")).visible = True
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "lblLevel")).visible = True
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "picBar")).visible = False
    
    ' set variables
    With Windows(GetWindowIndex("winDescription"))
        ' name
        If Not soulBound Then
            theName = Trim$(Item(ItemNum).Name)
        Else
            theName = "(SB) " & Trim$(Item(ItemNum).Name)
        End If
        .Controls(GetControlIndex("winDescription", "lblName")).text = theName
        Select Case Item(ItemNum).Rarity
            Case 0 ' white
                Colour = White
            Case 1 ' green
                Colour = Green
            Case 2 ' blue
                Colour = BrightBlue
            Case 3 ' maroon
                Colour = Red
            Case 4 ' purple
                Colour = Pink
            Case 5 ' orange
                Colour = Brown
        End Select
        .Controls(GetControlIndex("winDescription", "lblName")).textColour = Colour
        ' class req
        If Item(ItemNum).ClassReq > 0 Then
            className = Trim$(Class(Item(ItemNum).ClassReq).Name)
            ' do we match it?
            If GetPlayerClass(MyIndex) = Item(ItemNum).ClassReq Then
                Colour = Green
            Else
                Colour = BrightRed
            End If
        ElseIf Item(ItemNum).proficiency > 0 Then
            Select Case Item(ItemNum).proficiency
                Case 1 ' Sword/Armour
                    If Item(ItemNum).Type >= ITEM_TYPE_ARMOR And Item(ItemNum).Type <= ITEM_TYPE_FEET Then
                        className = "Heavy Armour"
                    ElseIf Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
                        className = "Heavy Weapon"
                    End If
                    If hasProficiency(MyIndex, Item(ItemNum).proficiency) Then
                        Colour = Green
                    Else
                        Colour = BrightRed
                    End If
                Case 2 ' Staff/Cloth
                    If Item(ItemNum).Type >= ITEM_TYPE_ARMOR And Item(ItemNum).Type <= ITEM_TYPE_FEET Then
                        className = "Cloth Armour"
                    ElseIf Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
                        className = "Light Weapon"
                    End If
                    If hasProficiency(MyIndex, Item(ItemNum).proficiency) Then
                        Colour = Green
                    Else
                        Colour = BrightRed
                    End If
            End Select
        Else
            className = "No class req."
            Colour = Green
        End If
        .Controls(GetControlIndex("winDescription", "lblClass")).text = className
        .Controls(GetControlIndex("winDescription", "lblClass")).textColour = Colour
        ' level
        If Item(ItemNum).LevelReq > 0 Then
            levelTxt = "Level " & Item(ItemNum).LevelReq
            ' do we match it?
            If GetPlayerLevel(MyIndex) >= Item(ItemNum).LevelReq Then
                Colour = Green
            Else
                Colour = BrightRed
            End If
        Else
            levelTxt = "No level req."
            Colour = Green
        End If
        .Controls(GetControlIndex("winDescription", "lblLevel")).text = levelTxt
        .Controls(GetControlIndex("winDescription", "lblLevel")).textColour = Colour
    End With
    
    ' clear
    ReDim descText(1 To 1) As TextColourRec
    
    ' go through the rest of the text
    Select Case Item(ItemNum).Type
        Case ITEM_TYPE_NONE
            AddDescInfo "No type"
        Case ITEM_TYPE_WEAPON
            AddDescInfo "Weapon"
        Case ITEM_TYPE_ARMOR
            AddDescInfo "Armour"
        Case ITEM_TYPE_HELMET
            AddDescInfo "Helmet"
        Case ITEM_TYPE_SHIELD
            AddDescInfo "Shield"
        Case ITEM_TYPE_PANTS
            AddDescInfo "Pants"
        Case ITEM_TYPE_FEET
            AddDescInfo "Feet"
        Case ITEM_TYPE_CONSUME
            AddDescInfo "Consume"
        Case ITEM_TYPE_KEY
            AddDescInfo "Key"
        Case ITEM_TYPE_CURRENCY
            AddDescInfo "Currency"
        Case ITEM_TYPE_SPELL
            AddDescInfo "Spell"
        Case ITEM_TYPE_FOOD
            AddDescInfo "Food"
    End Select
    
    ' more info
    Select Case Item(ItemNum).Type
        Case ITEM_TYPE_NONE, ITEM_TYPE_KEY, ITEM_TYPE_CURRENCY
            ' binding
            If Item(ItemNum).BindType = 1 Then
                AddDescInfo "Bind on Pickup"
            ElseIf Item(ItemNum).BindType = 2 Then
                AddDescInfo "Bind on Equip"
            End If
            ' price
            AddDescInfo "Value: " & Item(ItemNum).Price & "g"
        Case ITEM_TYPE_WEAPON, ITEM_TYPE_ARMOR, ITEM_TYPE_HELMET, ITEM_TYPE_SHIELD, ITEM_TYPE_PANTS, ITEM_TYPE_FEET
            ' damage/defence
            If Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
                AddDescInfo "Damage: " & Item(ItemNum).Data2
                ' speed
                AddDescInfo "Speed: " & (Item(ItemNum).Speed / 1000) & "s"
            Else
                If Item(ItemNum).Data2 > 0 Then
                    AddDescInfo "Defence: " & Item(ItemNum).Data2
                End If
            End If
            ' binding
            If Item(ItemNum).BindType = 1 Then
                AddDescInfo "Bind on Pickup"
            ElseIf Item(ItemNum).BindType = 2 Then
                AddDescInfo "Bind on Equip"
            End If
            ' price
            AddDescInfo "Value: " & Item(ItemNum).Price & "g"
            ' stat bonuses
            If Item(ItemNum).Add_Stat(Stats.Strength) > 0 Then
                AddDescInfo "+" & Item(ItemNum).Add_Stat(Stats.Strength) & " Str"
            End If
            If Item(ItemNum).Add_Stat(Stats.Endurance) > 0 Then
                AddDescInfo "+" & Item(ItemNum).Add_Stat(Stats.Endurance) & " End"
            End If
            If Item(ItemNum).Add_Stat(Stats.Intelligence) > 0 Then
                AddDescInfo "+" & Item(ItemNum).Add_Stat(Stats.Intelligence) & " Int"
            End If
            If Item(ItemNum).Add_Stat(Stats.Agility) > 0 Then
                AddDescInfo "+" & Item(ItemNum).Add_Stat(Stats.Agility) & " Agi"
            End If
            If Item(ItemNum).Add_Stat(Stats.Willpower) > 0 Then
                AddDescInfo "+" & Item(ItemNum).Add_Stat(Stats.Willpower) & " Will"
            End If
        Case ITEM_TYPE_CONSUME
            If Item(ItemNum).CastSpell > 0 Then
                AddDescInfo "Casts Spell"
            End If
            If Item(ItemNum).AddHP > 0 Then
                AddDescInfo "+" & Item(ItemNum).AddHP & " HP"
            End If
            If Item(ItemNum).AddMP > 0 Then
                AddDescInfo "+" & Item(ItemNum).AddMP & " SP"
            End If
            If Item(ItemNum).AddEXP > 0 Then
                AddDescInfo "+" & Item(ItemNum).AddEXP & " EXP"
            End If
            ' price
            AddDescInfo "Value: " & Item(ItemNum).Price & "g"
        Case ITEM_TYPE_SPELL
            ' price
            AddDescInfo "Value: " & Item(ItemNum).Price & "g"
        Case ITEM_TYPE_FOOD
            If Item(ItemNum).HPorSP = 2 Then
                AddDescInfo "Heal: " & (Item(ItemNum).FoodPerTick * Item(ItemNum).FoodTickCount) & " SP"
            Else
                AddDescInfo "Heal: " & (Item(ItemNum).FoodPerTick * Item(ItemNum).FoodTickCount) & " HP"
            End If
            ' time
            AddDescInfo "Time: " & (Item(ItemNum).FoodInterval * (Item(ItemNum).FoodTickCount / 1000)) & "s"
            ' price
            AddDescInfo "Value: " & Item(ItemNum).Price & "g"
    End Select
End Sub

Public Sub AddDescInfo(text As String, Optional Colour As Long = White)
Dim Count As Long
    Count = UBound(descText)
    ReDim Preserve descText(1 To Count + 1) As TextColourRec
    descText(Count + 1).text = text
    descText(Count + 1).Colour = Colour
End Sub

Public Sub SwitchHotbar(oldSlot As Long, newSlot As Long)
Dim oldSlot_type As Long, oldSlot_value As Long, newSlot_type As Long, newSlot_value As Long

    oldSlot_type = Hotbar(oldSlot).sType
    newSlot_type = Hotbar(newSlot).sType
    oldSlot_value = Hotbar(oldSlot).Slot
    newSlot_value = Hotbar(newSlot).Slot
    
    ' send the changes
    SendHotbarChange oldSlot_type, oldSlot_value, newSlot
    SendHotbarChange newSlot_type, newSlot_value, oldSlot
End Sub

Public Sub ShowChat()
    ShowWindow GetWindowIndex("winChat"), , False
    HideWindow GetWindowIndex("winChatSmall")
    ' Set the active control
    activeWindow = GetWindowIndex("winChat")
    SetActiveControl GetWindowIndex("winChat"), GetControlIndex("winChat", "txtChat")
    inSmallChat = False
    ChatScroll = 0
End Sub

Public Sub HideChat()
    ShowWindow GetWindowIndex("winChatSmall"), , False
    HideWindow GetWindowIndex("winChat")
    inSmallChat = True
    ChatScroll = 0
End Sub

Public Sub SetChatHeight(Height As Long)
    actChatHeight = Height
End Sub

Public Sub SetChatWidth(Width As Long)
    actChatWidth = Width
End Sub

Public Sub UpdateChat()
    SaveOptions
End Sub

Public Sub OpenShop(shopNum As Long)
    ' set globals
    InShop = shopNum
    shopSelectedSlot = 1
    shopSelectedItem = Shop(InShop).TradeItem(1).Item
    Windows(GetWindowIndex("winShop")).Controls(GetControlIndex("winShop", "chkSelling")).Value = 0
    Windows(GetWindowIndex("winShop")).Controls(GetControlIndex("winShop", "chkBuying")).Value = 1
    Windows(GetWindowIndex("winShop")).Controls(GetControlIndex("winShop", "btnSell")).visible = False
    Windows(GetWindowIndex("winShop")).Controls(GetControlIndex("winShop", "btnBuy")).visible = True
    shopIsSelling = False
    ' set the current item
    UpdateShop
    ' show the window
    ShowWindow GetWindowIndex("winShop")
End Sub

Public Sub CloseShop()
    SendCloseShop
    HideWindow GetWindowIndex("winShop")
    shopSelectedSlot = 0
    shopSelectedItem = 0
    shopIsSelling = False
    InShop = 0
End Sub

Sub UpdateShop()
Dim i As Long, CostValue As Long

    If InShop = 0 Then Exit Sub
    
    ' make sure we have an item selected
    If shopSelectedSlot = 0 Then shopSelectedSlot = 1
    
    With Windows(GetWindowIndex("winShop"))
        ' buying items
        If Not shopIsSelling Then
            shopSelectedItem = Shop(InShop).TradeItem(shopSelectedSlot).Item
            ' labels
            If shopSelectedItem > 0 Then
                .Controls(GetControlIndex("winShop", "lblName")).text = Trim$(Item(shopSelectedItem).Name)
                ' check if it's gold
                If Shop(InShop).TradeItem(shopSelectedSlot).CostItem = 1 Then
                    ' it's gold
                    .Controls(GetControlIndex("winShop", "lblCost")).text = Shop(InShop).TradeItem(shopSelectedSlot).CostValue & "g"
                Else
                    ' if it's one then just print the name
                    If Shop(InShop).TradeItem(shopSelectedSlot).CostValue = 1 Then
                        .Controls(GetControlIndex("winShop", "lblCost")).text = Trim$(Item(Shop(InShop).TradeItem(shopSelectedSlot).CostItem).Name)
                    Else
                        .Controls(GetControlIndex("winShop", "lblCost")).text = Shop(InShop).TradeItem(shopSelectedSlot).CostValue & " " & Trim$(Item(Shop(InShop).TradeItem(shopSelectedSlot).CostItem).Name)
                    End If
                End If
                ' draw the item
                For i = 0 To 5
                    .Controls(GetControlIndex("winShop", "picItem")).image(i) = TextureItem(Item(shopSelectedItem).pic)
                Next
            Else
                .Controls(GetControlIndex("winShop", "lblName")).text = "Empty Slot"
                .Controls(GetControlIndex("winShop", "lblCost")).text = vbNullString
                ' draw the item
                For i = 0 To 5
                    .Controls(GetControlIndex("winShop", "picItem")).image(i) = 0
                Next
            End If
        Else
            shopSelectedItem = GetPlayerInvItemNum(MyIndex, shopSelectedSlot)
            ' labels
            If shopSelectedItem > 0 Then
                .Controls(GetControlIndex("winShop", "lblName")).text = Trim$(Item(shopSelectedItem).Name)
                ' calc cost
                CostValue = (Item(shopSelectedItem).Price / 100) * Shop(InShop).BuyRate
                .Controls(GetControlIndex("winShop", "lblCost")).text = CostValue & "g"
                ' draw the item
                For i = 0 To 5
                    .Controls(GetControlIndex("winShop", "picItem")).image(i) = TextureItem(Item(shopSelectedItem).pic)
                Next
            Else
                .Controls(GetControlIndex("winShop", "lblName")).text = "Empty Slot"
                .Controls(GetControlIndex("winShop", "lblCost")).text = vbNullString
                ' draw the item
                For i = 0 To 5
                    .Controls(GetControlIndex("winShop", "picItem")).image(i) = 0
                Next
            End If
        End If
    End With
End Sub

Public Function IsShopSlot(StartX As Long, StartY As Long) As Long
Dim tempRec As RECT
Dim i As Long

    For i = 1 To MAX_TRADES
        With tempRec
            .Top = StartY + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
            .Bottom = .Top + PIC_Y
            .Left = StartX + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
            .Right = .Left + PIC_X
        End With

        If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
            If currMouseY >= tempRec.Top And currMouseY <= tempRec.Bottom Then
                IsShopSlot = i
                Exit Function
            End If
        End If
    Next
End Function

Sub ShowPlayerMenu(Index As Long, X As Long, Y As Long)
    PlayerMenuIndex = Index
    If PlayerMenuIndex = 0 Then Exit Sub
    Windows(GetWindowIndex("winPlayerMenu")).Window.Left = X - 5
    Windows(GetWindowIndex("winPlayerMenu")).Window.Top = Y - 5
    Windows(GetWindowIndex("winPlayerMenu")).Controls(GetControlIndex("winPlayerMenu", "btnName")).text = Trim$(GetPlayerName(PlayerMenuIndex))
    ShowWindow GetWindowIndex("winRightClickBG")
    ShowWindow GetWindowIndex("winPlayerMenu"), , False
End Sub

Public Function AryCount(ByRef Ary() As Byte) As Long
On Error Resume Next

    AryCount = UBound(Ary) + 1
End Function

Public Function ByteToInt(ByVal B1 As Long, ByVal B2 As Long) As Long
    ByteToInt = B1 * 256 + B2
End Function

Sub UpdateStats_UI()
    ' set the bar labels
    With Windows(GetWindowIndex("winBars"))
        .Controls(GetControlIndex("winBars", "lblHP")).text = GetPlayerVital(MyIndex, HP) & "/" & GetPlayerMaxVital(MyIndex, HP)
        .Controls(GetControlIndex("winBars", "lblMP")).text = GetPlayerVital(MyIndex, MP) & "/" & GetPlayerMaxVital(MyIndex, MP)
        .Controls(GetControlIndex("winBars", "lblEXP")).text = GetPlayerExp(MyIndex) & "/" & TNL
    End With
    ' update character screen
    With Windows(GetWindowIndex("winCharacter"))
        .Controls(GetControlIndex("winCharacter", "lblHealth")).text = "Health: " & GetPlayerVital(MyIndex, HP) & "/" & GetPlayerMaxVital(MyIndex, HP)
        .Controls(GetControlIndex("winCharacter", "lblSpirit")).text = "Spirit: " & GetPlayerVital(MyIndex, MP) & "/" & GetPlayerMaxVital(MyIndex, MP)
        .Controls(GetControlIndex("winCharacter", "lblExperience")).text = "Experience: " & Player(MyIndex).EXP & "/" & TNL
    End With
End Sub

Sub UpdatePartyInterface()
Dim i As Long, image(0 To 5) As Long, X As Long, pIndex As Long, Height As Long, cIn As Long

    ' unload it if we're not in a party
    If Party.Leader = 0 Then
        HideWindow GetWindowIndex("winParty")
        Exit Sub
    End If
    
    ' load the window
    ShowWindow GetWindowIndex("winParty")
    ' fill the controls
    With Windows(GetWindowIndex("winParty"))
        ' clear controls first
        For i = 1 To 3
            .Controls(GetControlIndex("winParty", "lblName" & i)).text = vbNullString
            .Controls(GetControlIndex("winParty", "picEmptyBar_HP" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picEmptyBar_SP" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picBar_HP" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picBar_SP" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picShadow" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picChar" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picChar" & i)).Value = 0
        Next
        ' labels
        cIn = 1
        For i = 1 To Party.MemberCount
            ' cache the index
            pIndex = Party.Member(i)
            If pIndex > 0 Then
                If pIndex <> MyIndex Then
                    If IsPlaying(pIndex) Then
                        ' name and level
                        .Controls(GetControlIndex("winParty", "lblName" & cIn)).visible = True
                        .Controls(GetControlIndex("winParty", "lblName" & cIn)).text = Trim$(GetPlayerName(pIndex))
                        ' picture
                        .Controls(GetControlIndex("winParty", "picShadow" & cIn)).visible = True
                        .Controls(GetControlIndex("winParty", "picChar" & cIn)).visible = True
                        ' store the player's index as a value for later use
                        .Controls(GetControlIndex("winParty", "picChar" & cIn)).Value = pIndex
                        For X = 0 To 5
                            .Controls(GetControlIndex("winParty", "picChar" & cIn)).image(X) = TextureChar(GetPlayerSprite(pIndex))
                        Next
                        ' bars
                        .Controls(GetControlIndex("winParty", "picEmptyBar_HP" & cIn)).visible = True
                        .Controls(GetControlIndex("winParty", "picEmptyBar_SP" & cIn)).visible = True
                        .Controls(GetControlIndex("winParty", "picBar_HP" & cIn)).visible = True
                        .Controls(GetControlIndex("winParty", "picBar_SP" & cIn)).visible = True
                        ' increment control usage
                        cIn = cIn + 1
                    End If
                End If
            End If
        Next
        ' update the bars
        UpdatePartyBars
        ' set the window size
        Select Case Party.MemberCount
            Case 2: Height = 78
            Case 3: Height = 118
            Case 4: Height = 158
        End Select
        .Window.Height = Height
    End With
End Sub

Sub UpdatePartyBars()
Dim i As Long, pIndex As Long, barWidth As Long, Width As Long

    ' unload it if we're not in a party
    If Party.Leader = 0 Then
        Exit Sub
    End If
    
    ' max bar width
    barWidth = 173
    
    ' make sure we're in a party
    With Windows(GetWindowIndex("winParty"))
        For i = 1 To 3
            ' get the pIndex from the control
            If .Controls(GetControlIndex("winParty", "picChar" & i)).visible = True Then
                pIndex = .Controls(GetControlIndex("winParty", "picChar" & i)).Value
                ' make sure they exist
                If pIndex > 0 Then
                    If IsPlaying(pIndex) Then
                        ' get their health
                        If GetPlayerVital(pIndex, HP) > 0 And GetPlayerMaxVital(pIndex, HP) > 0 Then
                            Width = ((GetPlayerVital(pIndex, Vitals.HP) / barWidth) / (GetPlayerMaxVital(pIndex, Vitals.HP) / barWidth)) * barWidth
                            .Controls(GetControlIndex("winParty", "picBar_HP" & i)).Width = Width
                        Else
                            .Controls(GetControlIndex("winParty", "picBar_HP" & i)).Width = 0
                        End If
                        ' get their spirit
                        If GetPlayerVital(pIndex, MP) > 0 And GetPlayerMaxVital(pIndex, MP) > 0 Then
                            Width = ((GetPlayerVital(pIndex, Vitals.MP) / barWidth) / (GetPlayerMaxVital(pIndex, Vitals.MP) / barWidth)) * barWidth
                            .Controls(GetControlIndex("winParty", "picBar_SP" & i)).Width = Width
                        Else
                            .Controls(GetControlIndex("winParty", "picBar_SP" & i)).Width = 0
                        End If
                    End If
                End If
            End If
        Next
    End With
End Sub

Sub ShowTrade()
    ' show the window
    ShowWindow GetWindowIndex("winTrade")
    ' set the controls up
    With Windows(GetWindowIndex("winTrade"))
        .Window.text = "Trading with " & Trim$(GetPlayerName(InTrade))
        .Controls(GetControlIndex("winTrade", "lblYourTrade")).text = Trim$(GetPlayerName(MyIndex)) & "'s Offer"
        .Controls(GetControlIndex("winTrade", "lblTheirTrade")).text = Trim$(GetPlayerName(InTrade)) & "'s Offer"
        .Controls(GetControlIndex("winTrade", "lblYourValue")).text = "0g"
        .Controls(GetControlIndex("winTrade", "lblTheirValue")).text = "0g"
        .Controls(GetControlIndex("winTrade", "lblStatus")).text = "Choose items to offer."
    End With
End Sub

Sub CheckResolution()
Dim resolution As Byte, Width As Long, Height As Long
    ' find the selected resolution
    resolution = Options.resolution
    ' reset
    If resolution = 0 Then
        resolution = 12
        ' loop through till we find one which fits
        Do Until ScreenFit(resolution) Or resolution > RES_COUNT
            ScreenFit resolution
            resolution = resolution + 1
            DoEvents
        Loop
        ' right resolution
        If resolution > RES_COUNT Then resolution = RES_COUNT
        Options.resolution = resolution
    End If
    
    ' size the window
    GetResolutionSize Options.resolution, Width, Height
    Resize Width, Height
    
    ' save it
    curResolution = Options.resolution
    
    SaveOptions
End Sub

Function ScreenFit(resolution As Byte) As Boolean
Dim sWidth As Long, sHeight As Long, Width As Long, Height As Long

    ' exit out early
    If resolution = 0 Then
        ScreenFit = False
        Exit Function
    End If

    ' get screen size
    sWidth = Screen.Width / Screen.TwipsPerPixelX
    sHeight = Screen.Height / Screen.TwipsPerPixelY
    
    GetResolutionSize resolution, Width, Height
    
    ' check if match
    If Width > sWidth Or Height > sHeight Then
        ScreenFit = False
    Else
        ScreenFit = True
    End If
End Function

Function GetResolutionSize(resolution As Byte, ByRef Width As Long, ByRef Height As Long)
    Select Case resolution
        Case 1
            Width = 1920
            Height = 1080
        Case 2
            Width = 1680
            Height = 1050
        Case 3
            Width = 1600
            Height = 900
        Case 4
            Width = 1440
            Height = 900
        Case 5
            Width = 1440
            Height = 1050
        Case 6
            Width = 1366
            Height = 768
        Case 7
            Width = 1360
            Height = 1024
        Case 8
            Width = 1360
            Height = 768
        Case 9
            Width = 1280
            Height = 1024
        Case 10
            Width = 1280
            Height = 800
        Case 11
            Width = 1280
            Height = 768
        Case 12
            Width = 1280
            Height = 720
        Case 13
            Width = 1024
            Height = 768
        Case 14
            Width = 1024
            Height = 576
        Case 15
            Width = 800
            Height = 600
        Case 16
            Width = 800
            Height = 450
    End Select
End Function

Sub Resize(ByVal Width As Long, ByVal Height As Long)
    frmMain.Width = (frmMain.Width \ 15 - frmMain.ScaleWidth + Width) * 15
    frmMain.Height = (frmMain.Height \ 15 - frmMain.ScaleHeight + Height) * 15
    frmMain.Left = (Screen.Width - frmMain.Width) \ 2
    frmMain.Top = (Screen.Height - frmMain.Height) \ 2
    
    '//Inicializar opções administrativas
    Call HandleDeveloperOptions
    
    DoEvents
End Sub

Sub ResizeGUI()
Dim Top As Long

    ' move hotbar
    Windows(GetWindowIndex("winHotbar")).Window.Left = ScreenWidth - 430
    ' move chat
    Windows(GetWindowIndex("winChat")).Window.Top = ScreenHeight - 178
    Windows(GetWindowIndex("winChatSmall")).Window.Top = ScreenHeight - 162
    ' move menu
    Windows(GetWindowIndex("winMenu")).Window.Left = ScreenWidth - 236
    Windows(GetWindowIndex("winMenu")).Window.Top = ScreenHeight - 37
    ' re-size right-click background
    Windows(GetWindowIndex("winRightClickBG")).Window.Width = ScreenWidth
    Windows(GetWindowIndex("winRightClickBG")).Window.Height = ScreenHeight
    ' re-size black background
    Windows(GetWindowIndex("winBlank")).Window.Width = ScreenWidth
    Windows(GetWindowIndex("winBlank")).Window.Height = ScreenHeight
    ' re-size combo background
    Windows(GetWindowIndex("winComboMenuBG")).Window.Width = ScreenWidth
    Windows(GetWindowIndex("winComboMenuBG")).Window.Height = ScreenHeight
    ' centralise windows
    CentraliseWindow GetWindowIndex("winLogin")
    CentraliseWindow GetWindowIndex("winCharacters")
    CentraliseWindow GetWindowIndex("winLoading")
    CentraliseWindow GetWindowIndex("winDialogue")
    CentraliseWindow GetWindowIndex("winClasses")
    CentraliseWindow GetWindowIndex("winNewChar")
    CentraliseWindow GetWindowIndex("winEscMenu")
    CentraliseWindow GetWindowIndex("winInventory")
    CentraliseWindow GetWindowIndex("winCharacter")
    CentraliseWindow GetWindowIndex("winSkills")
    CentraliseWindow GetWindowIndex("winOptions")
    CentraliseWindow GetWindowIndex("winShop")
    CentraliseWindow GetWindowIndex("winNpcChat")
    CentraliseWindow GetWindowIndex("winTrade")
    CentraliseWindow GetWindowIndex("winGuild")
End Sub

Sub SetResolution()
Dim Width As Long, Height As Long
    curResolution = Options.resolution
    GetResolutionSize curResolution, Width, Height
    Resize Width, Height
    ScreenWidth = Width
    ScreenHeight = Height
    TileWidth = Width / 32
    TileHeight = Height / 32
    ScreenX = (TileWidth + 1) * PIC_X
    ScreenY = (TileHeight + 1) * PIC_Y
    ResetGFX
    ResizeGUI
End Sub

Sub ShowComboMenu(curWindow As Long, curControl As Long)
Dim Top As Long
    With Windows(curWindow).Controls(curControl)
        ' linked to
        Windows(GetWindowIndex("winComboMenu")).Window.linkedToWin = curWindow
        Windows(GetWindowIndex("winComboMenu")).Window.linkedToCon = curControl
        ' set the size
        Windows(GetWindowIndex("winComboMenu")).Window.Height = 2 + (UBound(.list) * 16)
        Windows(GetWindowIndex("winComboMenu")).Window.Left = Windows(curWindow).Window.Left + .Left + 2
        Top = Windows(curWindow).Window.Top + .Top + .Height
        If Top + Windows(GetWindowIndex("winComboMenu")).Window.Height > ScreenHeight Then Top = ScreenHeight - Windows(GetWindowIndex("winComboMenu")).Window.Height
        Windows(GetWindowIndex("winComboMenu")).Window.Top = Top
        Windows(GetWindowIndex("winComboMenu")).Window.Width = .Width - 4
        ' set the values
        Windows(GetWindowIndex("winComboMenu")).Window.list() = .list()
        Windows(GetWindowIndex("winComboMenu")).Window.Value = .Value
        Windows(GetWindowIndex("winComboMenu")).Window.group = 0
        ' load the menu
        ShowWindow GetWindowIndex("winComboMenuBG"), True, False
        ShowWindow GetWindowIndex("winComboMenu"), True, False
    End With
End Sub

Sub ComboMenu_MouseMove(curWindow As Long)
Dim Y As Long, i As Long
    With Windows(curWindow).Window
        Y = currMouseY - .Top
        ' find the option we're hovering over
        If UBound(.list) > 0 Then
            For i = 1 To UBound(.list)
                If Y >= (16 * (i - 1)) And Y <= (16 * (i)) Then
                    .group = i
                End If
            Next
        End If
    End With
End Sub

Sub ComboMenu_MouseDown(curWindow As Long)
Dim Y As Long, i As Long
    With Windows(curWindow).Window
        Y = currMouseY - .Top
        ' find the option we're hovering over
        If UBound(.list) > 0 Then
            For i = 1 To UBound(.list)
                If Y >= (16 * (i - 1)) And Y <= (16 * (i)) Then
                    Windows(.linkedToWin).Controls(.linkedToCon).Value = i
                    CloseComboMenu
                End If
            Next
        End If
    End With
End Sub

Sub SetOptionsScreen()
    ' clear the combolists
    Erase Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRes")).list
    ReDim Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRes")).list(0)
    Erase Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRender")).list
    ReDim Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRender")).list(0)
    
    ' Resolutions
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1920x1080"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1680x1050"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1600x900"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1440x900"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1440x1050"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1366x768"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1360x1024"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1360x768"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1280x1024"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1280x800"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1280x768"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1280x720"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1024x768"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1024x576"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "800x600"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "800x450"
    
    ' Render Options
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRender"), "Automatic"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRender"), "Hardware"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRender"), "Mixed"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRender"), "Software"
    
    ' fill the options screen
    With Windows(GetWindowIndex("winOptions"))
        .Controls(GetControlIndex("winOptions", "chkMusic")).Value = Options.Music
        .Controls(GetControlIndex("winOptions", "chkSound")).Value = Options.sound
        If Options.NoAuto = 1 Then
            .Controls(GetControlIndex("winOptions", "chkAutotiles")).Value = 0
        Else
            .Controls(GetControlIndex("winOptions", "chkAutotiles")).Value = 1
        End If
        .Controls(GetControlIndex("winOptions", "chkFullscreen")).Value = Options.Fullscreen
        .Controls(GetControlIndex("winOptions", "cmbRes")).Value = Options.resolution
        .Controls(GetControlIndex("winOptions", "cmbRender")).Value = Options.Render + 1
    End With
End Sub

Function HasItem(ByVal ItemNum As Long) As Long
    Dim i As Long

    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(MyIndex, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(MyIndex, i)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next
End Function

Function ActiveEventPage(ByVal EventNum As Long) As Long
Dim X As Long, process As Boolean
    For X = Map.TileData.Events(EventNum).pageCount To 1 Step -1
        ' check if we match
        With Map.TileData.Events(EventNum).EventPage(X)
            process = True
            ' player var check
            If .chkPlayerVar Then
                If .PlayerVarNum > 0 Then
                    If Player(MyIndex).Variable(.PlayerVarNum) < .PlayerVariable Then
                        process = False
                    End If
                End If
            End If
            ' has item check
            If .chkHasItem Then
                If .HasItemNum > 0 Then
                    If HasItem(.HasItemNum) = 0 Then
                        process = False
                    End If
                End If
            End If
            ' this page
            If process = True Then
                ActiveEventPage = X
                Exit Function
            End If
        End With
    Next
End Function

Sub PlayerSwitchInvSlots(ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long, OldValue As Long, oldBound As Byte
Dim NewNum As Long, NewValue As Long, newBound As Byte

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(MyIndex, oldSlot)
    OldValue = GetPlayerInvItemValue(MyIndex, oldSlot)
    oldBound = PlayerInv(oldSlot).bound
    NewNum = GetPlayerInvItemNum(MyIndex, newSlot)
    NewValue = GetPlayerInvItemValue(MyIndex, newSlot)
    newBound = PlayerInv(newSlot).bound
    
    SetPlayerInvItemNum MyIndex, newSlot, OldNum
    SetPlayerInvItemValue MyIndex, newSlot, OldValue
    PlayerInv(newSlot).bound = oldBound
    
    SetPlayerInvItemNum MyIndex, oldSlot, NewNum
    SetPlayerInvItemValue MyIndex, oldSlot, NewValue
    PlayerInv(oldSlot).bound = newBound
End Sub

Sub PlayerSwitchSpellSlots(ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long, NewNum As Long, OldUses As Long, NewUses As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = PlayerSpells(oldSlot).Spell
    NewNum = PlayerSpells(newSlot).Spell
    OldUses = PlayerSpells(oldSlot).Uses
    NewUses = PlayerSpells(newSlot).Uses
    
    PlayerSpells(oldSlot).Spell = NewNum
    PlayerSpells(oldSlot).Uses = NewUses
    PlayerSpells(newSlot).Spell = OldNum
    PlayerSpells(newSlot).Uses = OldUses
End Sub

Sub CheckAppearTiles()
Dim X As Long, Y As Long, i As Long
    If GettingMap Then Exit Sub
    
    ' clear
    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            If Map.TileData.Tile(X, Y).Type = TILE_TYPE_APPEAR Then
                TempTile(X, Y).DoorOpen = 0
            End If
        Next
    Next
    
    ' set
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                X = GetPlayerX(i)
                Y = GetPlayerY(i)
                CheckAppearTile X, Y
                If Y - 1 >= 0 Then CheckAppearTile X, Y - 1
                If Y + 1 <= Map.MapData.MaxY Then CheckAppearTile X, Y + 1
                If X - 1 >= 0 Then CheckAppearTile X - 1, Y
                If X + 1 <= Map.MapData.MaxX Then CheckAppearTile X + 1, Y
            End If
        End If
    Next
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).Num > 0 Then
            If MapNpc(i).Vital(Vitals.HP) > 0 Then
                X = MapNpc(i).X
                Y = MapNpc(i).Y
                CheckAppearTile X, Y
                If Y - 1 >= 0 Then CheckAppearTile X, Y - 1
                If Y + 1 <= Map.MapData.MaxY Then CheckAppearTile X, Y + 1
                If X - 1 >= 0 Then CheckAppearTile X - 1, Y
                If X + 1 <= Map.MapData.MaxX Then CheckAppearTile X + 1, Y
            End If
        End If
    Next
    
    ' fade out old
    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            If TempTile(X, Y).DoorOpen = 0 Then
                ' exit if our mother is a bottom
                If Y > 0 Then
                    If Map.TileData.Tile(X, Y - 1).Data2 Then
                        If TempTile(X, Y - 1).DoorOpen = 1 Then GoTo continueLoop
                    End If
                End If
                ' not open - fade them out
                For i = 1 To MapLayer.Layer_Count - 1
                    If TempTile(X, Y).fadeAlpha(i) > 0 Then
                        TempTile(X, Y).isFading(i) = True
                        TempTile(X, Y).fadeAlpha(i) = TempTile(X, Y).fadeAlpha(i) - 1
                        TempTile(X, Y).FadeDir(i) = DIR_DOWN
                    End If
                Next
            End If
continueLoop:
        Next
    Next
End Sub

Sub CheckAppearTile(ByVal X As Long, ByVal Y As Long)
    If Y < 0 Or X < 0 Or Y > Map.MapData.MaxY Or X > Map.MapData.MaxX Then Exit Sub
    
    If Map.TileData.Tile(X, Y).Type = TILE_TYPE_APPEAR Then
        TempTile(X, Y).DoorOpen = 1
        
        If TempTile(X, Y).fadeAlpha(MapLayer.Mask) = 255 Then Exit Sub
        If TempTile(X, Y).isFading(MapLayer.Mask) Then
            If TempTile(X, Y).FadeDir(MapLayer.Mask) = DIR_DOWN Then
                TempTile(X, Y).FadeDir(MapLayer.Mask) = DIR_UP
                ' check if bottom
                If Y < Map.MapData.MaxY Then
                    If Map.TileData.Tile(X, Y).Data2 Then
                        TempTile(X, Y + 1).FadeDir(MapLayer.Ground) = DIR_UP
                    End If
                End If
                ' / bottom
            End If
            Exit Sub
        End If
        
        TempTile(X, Y).FadeDir(MapLayer.Mask) = DIR_UP
        TempTile(X, Y).isFading(MapLayer.Mask) = True
        TempTile(X, Y).fadeAlpha(MapLayer.Mask) = TempTile(X, Y).fadeAlpha(MapLayer.Mask) + 1
        
        ' check if bottom
        If Y < Map.MapData.MaxY Then
            If Map.TileData.Tile(X, Y).Data2 Then
                TempTile(X, Y + 1).FadeDir(MapLayer.Ground) = DIR_UP
                TempTile(X, Y + 1).isFading(MapLayer.Ground) = True
                TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) = TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) + 1
            End If
        End If
        ' / bottom
    End If
End Sub

Public Sub AppearTileFadeLogic()
Dim X As Long, Y As Long
    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            If Map.TileData.Tile(X, Y).Type = TILE_TYPE_APPEAR Then
                ' check if it's fading
                If TempTile(X, Y).isFading(MapLayer.Mask) Then
                    ' fading in
                    If TempTile(X, Y).FadeDir(MapLayer.Mask) = DIR_UP Then
                        If TempTile(X, Y).fadeAlpha(MapLayer.Mask) < 255 Then
                            TempTile(X, Y).fadeAlpha(MapLayer.Mask) = TempTile(X, Y).fadeAlpha(MapLayer.Mask) + 20
                            ' check if bottom
                            If Y < Map.MapData.MaxY Then
                                If Map.TileData.Tile(X, Y).Data2 Then
                                    TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) = TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) + 20
                                End If
                            End If
                            ' / bottom
                        End If
                        If TempTile(X, Y).fadeAlpha(MapLayer.Mask) >= 255 Then
                            TempTile(X, Y).fadeAlpha(MapLayer.Mask) = 255
                            TempTile(X, Y).isFading(MapLayer.Mask) = False
                            ' check if bottom
                            If Y < Map.MapData.MaxY Then
                                If Map.TileData.Tile(X, Y).Data2 Then
                                    TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) = 255
                                    TempTile(X, Y + 1).isFading(MapLayer.Ground) = False
                                End If
                            End If
                            ' / bottom
                        End If
                    ElseIf TempTile(X, Y).FadeDir(MapLayer.Mask) = DIR_DOWN Then
                        If TempTile(X, Y).fadeAlpha(MapLayer.Mask) > 0 Then
                            TempTile(X, Y).fadeAlpha(MapLayer.Mask) = TempTile(X, Y).fadeAlpha(MapLayer.Mask) - 20
                            ' check if bottom
                            If Y < Map.MapData.MaxY Then
                                If Map.TileData.Tile(X, Y).Data2 Then
                                    TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) = TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) - 20
                                End If
                            End If
                            ' / bottom
                        End If
                        If TempTile(X, Y).fadeAlpha(MapLayer.Mask) <= 0 Then
                            TempTile(X, Y).fadeAlpha(MapLayer.Mask) = 0
                            TempTile(X, Y).isFading(MapLayer.Mask) = False
                            ' check if bottom
                            If Y < Map.MapData.MaxY Then
                                If Map.TileData.Tile(X, Y).Data2 Then
                                    TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) = 0
                                    TempTile(X, Y + 1).isFading(MapLayer.Ground) = False
                                End If
                            End If
                            ' / bottom
                        End If
                    End If
                End If
            End If
        Next
    Next
End Sub

Public Sub ProcessWeather()
    Dim i As Long
    If CurrentWeather > 0 Then
        i = Rand(1, 101 - CurrentWeatherIntensity)
        If i = 1 Then
            'Add a new particle
            For i = 1 To MAX_WEATHER_PARTICLES
                If WeatherParticle(i).InUse = False Then
                    If Rand(1, 2) = 1 Then
                        WeatherParticle(i).InUse = True
                        WeatherParticle(i).Type = CurrentWeather
                        WeatherParticle(i).Velocity = Rand(8, 14)
                        WeatherParticle(i).X = (TileView.Left * 32) - 32
                        WeatherParticle(i).Y = (TileView.Top * 32) + Rand(-32, frmMain.ScaleHeight)
                    Else
                        WeatherParticle(i).InUse = True
                        WeatherParticle(i).Type = CurrentWeather
                        WeatherParticle(i).Velocity = Rand(10, 15)
                        WeatherParticle(i).X = (TileView.Left * 32) + Rand(-32, frmMain.ScaleWidth)
                        WeatherParticle(i).Y = (TileView.Top * 32) - 32
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    
    If CurrentWeather = WEATHER_TYPE_STORM Then
        i = Rand(1, 400 - CurrentWeatherIntensity)
        If i = 1 Then
            'Draw Thunder
            DrawThunder = Rand(15, 22)
            Play_Sound Sound_Thunder, -1, -1
        End If
    End If
    
    For i = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(i).InUse Then
            If WeatherParticle(i).X > TileView.Right * 32 Or WeatherParticle(i).Y > TileView.Bottom * 32 Then
                WeatherParticle(i).InUse = False
            Else
                WeatherParticle(i).X = WeatherParticle(i).X + WeatherParticle(i).Velocity
                WeatherParticle(i).Y = WeatherParticle(i).Y + WeatherParticle(i).Velocity
            End If
        End If
    Next
End Sub
