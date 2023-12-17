Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim i As Long, X As Long

    ServerOnline = True

    Do While ServerOnline
        tick = GetTickCount
        ElapsedTime = tick - FrameTime
        FrameTime = tick
        
        If tick > tmr25 Then
            ' loops
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(i).spellBuffer.Spell > 0 Then
                        If GetTickCount > TempPlayer(i).spellBuffer.Timer + (Spell(Player(i).Spell(TempPlayer(i).spellBuffer.Spell).Spell).CastTime * 1000) Then
                            CastSpell i, TempPlayer(i).spellBuffer.Spell, TempPlayer(i).spellBuffer.Target, TempPlayer(i).spellBuffer.tType
                            TempPlayer(i).spellBuffer.Spell = 0
                            TempPlayer(i).spellBuffer.Timer = 0
                            TempPlayer(i).spellBuffer.Target = 0
                            TempPlayer(i).spellBuffer.tType = 0
                        End If
                    End If
                    ' check if need to turn off stunned
                    If TempPlayer(i).StunDuration > 0 Then
                        If GetTickCount > TempPlayer(i).StunTimer + (TempPlayer(i).StunDuration * 1000) Then
                            TempPlayer(i).StunDuration = 0
                            TempPlayer(i).StunTimer = 0
                            SendStunned i
                        End If
                    End If
                    ' check regen timer
                    If TempPlayer(i).stopRegen Then
                        If TempPlayer(i).stopRegenTimer + 5000 < GetTickCount Then
                            TempPlayer(i).stopRegen = False
                            TempPlayer(i).stopRegenTimer = 0
                        End If
                    End If
                    ' HoT and DoT logic
                    For X = 1 To MAX_DOTS
                        HandleDoT_Player i, X
                        HandleHoT_Player i, X
                    Next
                    ' food processing
                    UpdatePlayerFood i
                End If
            Next
            ' update entity logic
            UpdateMapEntities
            ' update label
            frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
            tmr25 = GetTickCount + 25
        End If

        ' Check for disconnections every half second
        If tick > tmr500 Then
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).State > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            UpdateMapLogic
            tmr500 = GetTickCount + 500
        End If
        
        ' draw projectiles on map
        For i = 1 To MapProjectile_HighIndex
            If MapProjectile(i).Owner > 0 Then
                CheckProjectile i
            End If
        Next

        If tick > tmr1000 Then
            ' check if shutting down
            If isShuttingDown Then
                Call HandleShutdown
            End If

            ' reset timer
            tmr1000 = GetTickCount + 1000
        End If

        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If tick > LastUpdatePlayerVitals Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = GetTickCount + 5000
        End If

        ' Checks to save players every 5 minutes - Can be tweaked
        If tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            LastUpdateSavePlayers = GetTickCount + 300000
        End If

        If Not CPSUnlock Then Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < tick Then
            GameCPS = CPS
            TickCPS = tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
    Loop
End Sub

Sub UpdateMapEntities()
Dim mapnum As Long, i As Long, tick As Long, x1 As Long, y1 As Long, X As Long, Y As Long, Resource_index As Long

    tick = GetTickCount

    For mapnum = 1 To MAX_MAPS
        ' items appearing to everyone
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(mapnum, i).Num > 0 Then
                If MapItem(mapnum, i).playerName <> vbNullString Then
                    ' make item public?
                    If Not MapItem(mapnum, i).Bound Then
                        If MapItem(mapnum, i).playerTimer < tick Then
                            ' make it public
                            MapItem(mapnum, i).playerName = vbNullString
                            MapItem(mapnum, i).playerTimer = 0
                            ' send updates to everyone
                            SendMapItemsToAll mapnum
                        End If
                    End If
                    ' despawn item?
                    If MapItem(mapnum, i).canDespawn Then
                        If MapItem(mapnum, i).despawnTimer < tick Then
                            ' despawn it
                            ClearMapItem i, mapnum
                            ' send updates to everyone
                            SendMapItemsToAll mapnum
                        End If
                    End If
                End If
            End If
        Next
        
        '  Close the doors
        If tick > TempTile(mapnum).DoorTimer + 5000 Then
            For x1 = 0 To Map(mapnum).MapData.MaxX
                For y1 = 0 To Map(mapnum).MapData.MaxY
                    If Map(mapnum).TileData.Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(mapnum).DoorOpen(x1, y1) = YES Then
                        TempTile(mapnum).DoorOpen(x1, y1) = NO
                        SendMapKeyToMap mapnum, x1, y1, 0
                    End If
                Next
            Next
        End If
        
        ' check for DoTs + hots
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(mapnum).Npc(i).Num > 0 Then
                For X = 1 To MAX_DOTS
                    HandleDoT_Npc mapnum, i, X
                    HandleHoT_Npc mapnum, i, X
                Next
            End If
        Next

        ' Respawning Resources
        If ResourceCache(mapnum).Resource_Count > 0 Then
            For i = 0 To ResourceCache(mapnum).Resource_Count
                Resource_index = Map(mapnum).TileData.Tile(ResourceCache(mapnum).ResourceData(i).X, ResourceCache(mapnum).ResourceData(i).Y).Data1

                If Resource_index > 0 Then
                    If ResourceCache(mapnum).ResourceData(i).ResourceState = 1 Or ResourceCache(mapnum).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
                        If ResourceCache(mapnum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < tick Then
                            ResourceCache(mapnum).ResourceData(i).ResourceTimer = tick
                            ResourceCache(mapnum).ResourceData(i).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            ResourceCache(mapnum).ResourceData(i).cur_health = Resource(Resource_index).health
                            SendResourceCacheToMap mapnum, i
                        End If
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Sub UpdateMapLogic()
    Dim i As Long, X As Long, mapnum As Long, n As Long, x1 As Long, y1 As Long
    Dim TickCount As Long, Damage As Long, DistanceX As Long, DistanceY As Long, npcNum As Long
    Dim Target As Long, targetType As Byte, DidWalk As Boolean, buffer As clsBuffer, Resource_index As Long
    Dim targetX As Long, targetY As Long, target_verify As Boolean

    For mapnum = 1 To MAX_MAPS
        If PlayersOnMap(mapnum) = YES Then
            TickCount = GetTickCount
            
            For X = 1 To MAX_MAP_NPCS
                npcNum = MapNpc(mapnum).Npc(X).Num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapnum).MapData.Npc(X) > 0 And MapNpc(mapnum).Npc(X).Num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(npcNum).Behaviour = NPC_BEHAVIOUR_GUARD Then
                    
                        ' make sure it's not stunned
                        If Not MapNpc(mapnum).Npc(X).StunDuration > 0 Then
    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = mapnum And MapNpc(mapnum).Npc(X).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                        ' make sure it's within the level range
                                        If (GetPlayerLevel(i) <= Npc(npcNum).Level - 2) Or (Map(mapnum).MapData.Moral = MAP_MORAL_BOSS) Then
                                            n = Npc(npcNum).Range
                                            DistanceX = MapNpc(mapnum).Npc(X).X - GetPlayerX(i)
                                            DistanceY = MapNpc(mapnum).Npc(X).Y - GetPlayerY(i)
        
                                            ' Make sure we get a positive value
                                            If DistanceX < 0 Then DistanceX = DistanceX * -1
                                            If DistanceY < 0 Then DistanceY = DistanceY * -1
        
                                            ' Are they in range?  if so GET'M!
                                            If DistanceX <= n And DistanceY <= n Then
                                                If Npc(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                    If Len(Trim$(Npc(npcNum).AttackSay)) > 0 Then
                                                        Call PlayerMsg(i, Trim$(Npc(npcNum).Name) & " says: " & Trim$(Npc(npcNum).AttackSay), SayColor)
                                                    End If
                                                    MapNpc(mapnum).Npc(X).targetType = 1 ' player
                                                    MapNpc(mapnum).Npc(X).Target = i
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
                
                target_verify = False

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapnum).MapData.Npc(X) > 0 And MapNpc(mapnum).Npc(X).Num > 0 Then
                    If MapNpc(mapnum).Npc(X).StunDuration > 0 Then
                        ' check if we can unstun them
                        If GetTickCount > MapNpc(mapnum).Npc(X).StunTimer + (MapNpc(mapnum).Npc(X).StunDuration * 1000) Then
                            MapNpc(mapnum).Npc(X).StunDuration = 0
                            MapNpc(mapnum).Npc(X).StunTimer = 0
                        End If
                    Else
                        ' check if in conversation
                        If MapNpc(mapnum).Npc(X).c_inChatWith > 0 Then
                            ' check if we can stop having conversation
                            If Not TempPlayer(MapNpc(mapnum).Npc(X).c_inChatWith).inChatWith = npcNum Then
                                MapNpc(mapnum).Npc(X).c_inChatWith = 0
                                MapNpc(mapnum).Npc(X).Dir = MapNpc(mapnum).Npc(X).c_lastDir
                                NpcDir mapnum, X, MapNpc(mapnum).Npc(X).Dir
                            End If
                        Else
                            Target = MapNpc(mapnum).Npc(X).Target
                            targetType = MapNpc(mapnum).Npc(X).targetType
        
                            ' Check to see if its time for the npc to walk
                            If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                            
                                If targetType = 1 Then ' player
        
                                    ' Check to see if we are following a player or not
                                    If Target > 0 Then
            
                                        ' Check if the player is even playing, if so follow'm
                                        If IsPlaying(Target) And GetPlayerMap(Target) = mapnum Then
                                            DidWalk = False
                                            target_verify = True
                                            targetY = GetPlayerY(Target)
                                            targetX = GetPlayerX(Target)
                                        Else
                                            MapNpc(mapnum).Npc(X).targetType = 0 ' clear
                                            MapNpc(mapnum).Npc(X).Target = 0
                                        End If
                                    End If
                                
                                ElseIf targetType = 2 Then 'npc
                                    
                                    If Target > 0 Then
                                        
                                        If MapNpc(mapnum).Npc(Target).Num > 0 Then
                                            DidWalk = False
                                            target_verify = True
                                            targetY = MapNpc(mapnum).Npc(Target).Y
                                            targetX = MapNpc(mapnum).Npc(Target).X
                                        Else
                                            MapNpc(mapnum).Npc(X).targetType = 0 ' clear
                                            MapNpc(mapnum).Npc(X).Target = 0
                                        End If
                                    End If
                                End If
                                
                                If target_verify Then
                                    
                                    i = Int(Rnd * 5)
        
                                    ' Lets move the npc
                                    Select Case i
                                        Case 0
                                        
                                            ' UpLeft
                                            If MapNpc(mapnum).Npc(X).X > targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y > targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_UP_LEFT) Then
                                                        Call NpcMove(mapnum, X, DIR_UP_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
                                            
                                            ' UpRight
                                            If MapNpc(mapnum).Npc(X).X < targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y > targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_UP_RIGHT) Then
                                                        Call NpcMove(mapnum, X, DIR_UP_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' DownLeft
                                            If MapNpc(mapnum).Npc(X).X > targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y < targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_DOWN_LEFT) Then
                                                        Call NpcMove(mapnum, X, DIR_DOWN_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' DownRight
                                            If MapNpc(mapnum).Npc(X).X < targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y < targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_DOWN_RIGHT) Then
                                                        Call NpcMove(mapnum, X, DIR_DOWN_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Up
                                            If MapNpc(mapnum).Npc(X).Y > targetY And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_UP) Then
                                                    Call NpcMove(mapnum, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Down
                                            If MapNpc(mapnum).Npc(X).Y < targetY And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_DOWN) Then
                                                    Call NpcMove(mapnum, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Left
                                            If MapNpc(mapnum).Npc(X).X > targetX And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_LEFT) Then
                                                    Call NpcMove(mapnum, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                            ' Right
                                            If MapNpc(mapnum).Npc(X).X < targetX And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_RIGHT) Then
                                                    Call NpcMove(mapnum, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                        Case 1
        
                                            ' Right
                                            If MapNpc(mapnum).Npc(X).X < targetX And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_RIGHT) Then
                                                    Call NpcMove(mapnum, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            
                                            ' UpLeft
                                            If MapNpc(mapnum).Npc(X).X > targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y > targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_UP_LEFT) Then
                                                        Call NpcMove(mapnum, X, DIR_UP_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Left
                                            If MapNpc(mapnum).Npc(X).X > targetX And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_LEFT) Then
                                                    Call NpcMove(mapnum, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            
                                            ' UpRight
                                            If MapNpc(mapnum).Npc(X).X < targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y > targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_UP_RIGHT) Then
                                                        Call NpcMove(mapnum, X, DIR_UP_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Down
                                            If MapNpc(mapnum).Npc(X).Y < targetY And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_DOWN) Then
                                                    Call NpcMove(mapnum, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            
                                            ' DownLeft
                                            If MapNpc(mapnum).Npc(X).X > targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y < targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_DOWN_LEFT) Then
                                                        Call NpcMove(mapnum, X, DIR_DOWN_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Up
                                            If MapNpc(mapnum).Npc(X).Y > targetY And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_UP) Then
                                                    Call NpcMove(mapnum, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            
                                            ' DownRight
                                            If MapNpc(mapnum).Npc(X).X < targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y < targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_DOWN_RIGHT) Then
                                                        Call NpcMove(mapnum, X, DIR_DOWN_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                        Case 2
        
                                            ' Down
                                            If MapNpc(mapnum).Npc(X).Y < targetY And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_DOWN) Then
                                                    Call NpcMove(mapnum, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            
                                            ' UpRight
                                            If MapNpc(mapnum).Npc(X).X < targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y > targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_UP_RIGHT) Then
                                                        Call NpcMove(mapnum, X, DIR_UP_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Up
                                            If MapNpc(mapnum).Npc(X).Y > targetY And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_UP) Then
                                                    Call NpcMove(mapnum, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            
                                            ' DownLeft
                                            If MapNpc(mapnum).Npc(X).X > targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y < targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_DOWN_LEFT) Then
                                                        Call NpcMove(mapnum, X, DIR_DOWN_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Right
                                            If MapNpc(mapnum).Npc(X).X < targetX And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_RIGHT) Then
                                                    Call NpcMove(mapnum, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            
                                            ' UpLeft
                                            If MapNpc(mapnum).Npc(X).X > targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y > targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_UP_LEFT) Then
                                                        Call NpcMove(mapnum, X, DIR_UP_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Left
                                            If MapNpc(mapnum).Npc(X).X > targetX And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_LEFT) Then
                                                    Call NpcMove(mapnum, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            
                                            ' DownRight
                                            If MapNpc(mapnum).Npc(X).X < targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y < targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_DOWN_RIGHT) Then
                                                        Call NpcMove(mapnum, X, DIR_DOWN_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                        Case 3
                                        
                                            ' UpRight
                                            If MapNpc(mapnum).Npc(X).X < targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y > targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_UP_RIGHT) Then
                                                        Call NpcMove(mapnum, X, DIR_UP_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Left
                                            If MapNpc(mapnum).Npc(X).X > targetX And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_LEFT) Then
                                                    Call NpcMove(mapnum, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            
                                            ' UpLeft
                                            If MapNpc(mapnum).Npc(X).X > targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y > targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_UP_LEFT) Then
                                                        Call NpcMove(mapnum, X, DIR_UP_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Right
                                            If MapNpc(mapnum).Npc(X).X < targetX And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_RIGHT) Then
                                                    Call NpcMove(mapnum, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            
                                            ' DownLeft
                                            If MapNpc(mapnum).Npc(X).X > targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y < targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_DOWN_LEFT) Then
                                                        Call NpcMove(mapnum, X, DIR_DOWN_LEFT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Up
                                            If MapNpc(mapnum).Npc(X).Y > targetY And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_UP) Then
                                                    Call NpcMove(mapnum, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            
                                            ' DownRight
                                            If MapNpc(mapnum).Npc(X).X < targetX And Not DidWalk Then
                                                If MapNpc(mapnum).Npc(X).Y < targetY And Not DidWalk Then
                                                    If CanNpcMove(mapnum, X, DIR_DOWN_RIGHT) Then
                                                        Call NpcMove(mapnum, X, DIR_DOWN_RIGHT, MOVING_WALKING)
                                                        DidWalk = True
                                                    End If
                                                End If
                                            End If
        
                                            ' Down
                                            If MapNpc(mapnum).Npc(X).Y < targetY And Not DidWalk Then
                                                If CanNpcMove(mapnum, X, DIR_DOWN) Then
                                                    Call NpcMove(mapnum, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
        
                                    End Select
        
                                    ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                    If Not DidWalk Then
                                        If MapNpc(mapnum).Npc(X).X - 1 = targetX And MapNpc(mapnum).Npc(X).Y = targetY Then
                                            If MapNpc(mapnum).Npc(X).Dir <> DIR_LEFT Then
                                                Call NpcDir(mapnum, X, DIR_LEFT)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        If MapNpc(mapnum).Npc(X).X + 1 = targetX And MapNpc(mapnum).Npc(X).Y = targetY Then
                                            If MapNpc(mapnum).Npc(X).Dir <> DIR_RIGHT Then
                                                Call NpcDir(mapnum, X, DIR_RIGHT)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        If MapNpc(mapnum).Npc(X).X = targetX And MapNpc(mapnum).Npc(X).Y - 1 = targetY Then
                                            If MapNpc(mapnum).Npc(X).Dir <> DIR_UP Then
                                                Call NpcDir(mapnum, X, DIR_UP)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        If MapNpc(mapnum).Npc(X).X = targetX And MapNpc(mapnum).Npc(X).Y + 1 = targetY Then
                                            If MapNpc(mapnum).Npc(X).Dir <> DIR_DOWN Then
                                                Call NpcDir(mapnum, X, DIR_DOWN)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        ' We could not move so Target must be behind something, walk randomly.
                                        If Not DidWalk Then
                                            i = Int(Rnd * 2)
        
                                            If i = 1 Then
                                                i = Int(Rnd * 8)
        
                                                If CanNpcMove(mapnum, X, i) Then
                                                    Call NpcMove(mapnum, X, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If
        
                                Else
                                    i = Int(Rnd * 4)
        
                                    If i = 1 Then
                                        i = Int(Rnd * 8)
        
                                        If CanNpcMove(mapnum, X, i) Then
                                            Call NpcMove(mapnum, X, i, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapnum).MapData.Npc(X) > 0 And MapNpc(mapnum).Npc(X).Num > 0 Then
                    Target = MapNpc(mapnum).Npc(X).Target
                    targetType = MapNpc(mapnum).Npc(X).targetType

                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                        If targetType = 1 Then ' player
                            ' Is the target playing and on the same map?
                            If IsPlaying(Target) And GetPlayerMap(Target) = mapnum Then
                                ' melee combat
                                TryNpcAttackPlayer X, Target
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(mapnum).Npc(X).Target = 0
                                MapNpc(mapnum).Npc(X).targetType = 0 ' clear
                            End If
                        End If
                    End If
                    
                    ' check for spells
                    If MapNpc(mapnum).Npc(X).spellBuffer.Spell = 0 Then
                        ' loop through and try and cast our spells
                        For i = 1 To MAX_NPC_SPELLS
                            If Npc(npcNum).Spell(i) > 0 Then
                                NpcBufferSpell mapnum, X, i
                            End If
                        Next
                    Else
                        ' check the timer
                        If MapNpc(mapnum).Npc(X).spellBuffer.Timer + (Spell(Npc(npcNum).Spell(MapNpc(mapnum).Npc(X).spellBuffer.Spell)).CastTime * 1000) < GetTickCount Then
                            ' cast the spell
                            NpcCastSpell mapnum, X, MapNpc(mapnum).Npc(X).spellBuffer.Spell, MapNpc(mapnum).Npc(X).spellBuffer.Target, MapNpc(mapnum).Npc(X).spellBuffer.tType
                            ' clear the buffer
                            MapNpc(mapnum).Npc(X).spellBuffer.Spell = 0
                            MapNpc(mapnum).Npc(X).spellBuffer.Target = 0
                            MapNpc(mapnum).Npc(X).spellBuffer.Timer = 0
                            MapNpc(mapnum).Npc(X).spellBuffer.tType = 0
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If Not MapNpc(mapnum).Npc(X).stopRegen Then
                    If MapNpc(mapnum).Npc(X).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                        If MapNpc(mapnum).Npc(X).Vital(Vitals.HP) > 0 Then
                            MapNpc(mapnum).Npc(X).Vital(Vitals.HP) = MapNpc(mapnum).Npc(X).Vital(Vitals.HP) + GetNpcVitalRegen(npcNum, Vitals.HP)
    
                            ' Check if they have more then they should and if so just set it to max
                            If MapNpc(mapnum).Npc(X).Vital(Vitals.HP) > GetNpcMaxVital(npcNum, Vitals.HP) Then
                                MapNpc(mapnum).Npc(X).Vital(Vitals.HP) = GetNpcMaxVital(npcNum, Vitals.HP)
                            End If
                        End If
                    End If
                End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(mapnum).Npc(X).Num = 0 And Map(mapnum).MapData.Npc(X) > 0 Then
                    If TickCount > MapNpc(mapnum).Npc(X).SpawnWait + (Npc(Map(mapnum).MapData.Npc(X)).SpawnSecs * 1000) Then
                        ' if it's a boss chamber then don't let them respawn
                        If Map(mapnum).MapData.Moral = MAP_MORAL_BOSS Then
                            ' make sure the boss is alive
                            If Map(mapnum).MapData.BossNpc > 0 Then
                                If Map(mapnum).MapData.Npc(Map(mapnum).MapData.BossNpc) > 0 Then
                                    If X <> Map(mapnum).MapData.BossNpc Then
                                        If MapNpc(mapnum).Npc(Map(mapnum).MapData.BossNpc).Num > 0 Then
                                            Call SpawnNpc(X, mapnum)
                                        End If
                                    Else
                                        SpawnNpc X, mapnum
                                    End If
                                End If
                            End If
                        Else
                            Call SpawnNpc(X, mapnum)
                        End If
                    End If
                End If

            Next

        End If

        DoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If

End Sub

Private Sub UpdatePlayerFood(ByVal i As Long)
Dim vitalType As Long, colour As Long, X As Long

    For X = 1 To Vitals.Vital_Count - 1
        If TempPlayer(i).foodItem(X) > 0 Then
            ' make sure not in combat
            If Not TempPlayer(i).stopRegen Then
                ' timer ready?
                If TempPlayer(i).foodTimer(X) + Item(TempPlayer(i).foodItem(X)).FoodInterval < GetTickCount Then
                    ' get vital type
                    If Item(TempPlayer(i).foodItem(X)).HPorSP = 2 Then vitalType = Vitals.MP Else vitalType = Vitals.HP
                    ' make sure we haven't gone over the top
                    If GetPlayerVital(i, vitalType) >= GetPlayerMaxVital(i, vitalType) Then
                        ' bring it back down to normal
                        SetPlayerVital i, vitalType, GetPlayerMaxVital(i, vitalType)
                        SendVital i, vitalType
                        ' remove the food - no point healing when full
                        TempPlayer(i).foodItem(X) = 0
                        TempPlayer(i).foodTick(X) = 0
                        TempPlayer(i).foodTimer(X) = 0
                        Exit Sub
                    End If
                    ' give them the healing
                    SetPlayerVital i, vitalType, GetPlayerVital(i, vitalType) + Item(TempPlayer(i).foodItem(X)).FoodPerTick
                    ' let them know with messages
                    If vitalType = 2 Then colour = BrightBlue Else colour = Green
                    SendActionMsg GetPlayerMap(i), "+" & Item(TempPlayer(i).foodItem(X)).FoodPerTick, colour, ACTIONMSG_SCROLL, GetPlayerX(i) * 32, GetPlayerY(i) * 32
                    ' send vitals
                    SendVital i, vitalType
                    ' increment tick count
                    TempPlayer(i).foodTick(X) = TempPlayer(i).foodTick(X) + 1
                    ' make sure we're not over max ticks
                    If TempPlayer(i).foodTick(X) >= Item(TempPlayer(i).foodItem(X)).FoodTickCount Then
                        ' clear food
                        TempPlayer(i).foodItem(X) = 0
                        TempPlayer(i).foodTick(X) = 0
                        TempPlayer(i).foodTimer(X) = 0
                        Exit Sub
                    End If
                    ' reset the timer
                    TempPlayer(i).foodTimer(X) = GetTickCount
                End If
            Else
                ' remove the food effect
                TempPlayer(i).foodItem(X) = 0
                TempPlayer(i).foodTick(X) = 0
                TempPlayer(i).foodTimer(X) = 0
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub UpdatePlayerVitals()
Dim i As Long
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not TempPlayer(i).stopRegen Then
                If GetPlayerVital(i, Vitals.HP) <> GetPlayerMaxVital(i, Vitals.HP) Then
                    Call SetPlayerVital(i, Vitals.HP, GetPlayerVital(i, Vitals.HP) + GetPlayerVitalRegen(i, Vitals.HP))
                    Call SendVital(i, Vitals.HP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
    
                If GetPlayerVital(i, Vitals.MP) <> GetPlayerMaxVital(i, Vitals.MP) Then
                    Call SetPlayerVital(i, Vitals.MP, GetPlayerVital(i, Vitals.MP) + GetPlayerVitalRegen(i, Vitals.MP))
                    Call SendVital(i, Vitals.MP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
            End If
        End If
    Next
End Sub

Private Sub UpdateSavePlayers()
    Dim i As Long

    If TotalOnlinePlayers > 0 Then
        Call TextAdd("Saving all online players...")

        For i = 1 To Player_HighIndex

            If IsPlaying(i) Then
                Call SavePlayer(i)
            End If

            DoEvents
        Next

    End If

End Sub

Private Sub HandleShutdown()

    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        Call GlobalMsg("Server Shutdown.", BrightRed)
        Call DestroyServer
    End If

End Sub
