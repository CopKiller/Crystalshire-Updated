Attribute VB_Name = "modGameEditors"
Option Explicit

' Temp event storage
Public tmpEvent As EventRec
Public tmpItem As ItemRec
Public tmpSpell As SpellRec

Public curPageNum As Long
Public curCommand As Long
Public GraphicSelX As Long
Public GraphicSelY As Long

' ////////////////
' // Map Editor //
' ////////////////
Public Sub MapEditorInit()
    Dim i As Long
    ' set the width
    frmEditor_Map.Width = 9585
    ' we're in the map editor
    InMapEditor = True
    ' show the form
    frmEditor_Map.visible = True
    ' set the scrolly bars
    frmEditor_Map.scrlTileSet.max = Count_Tileset
    frmEditor_Map.fraTileSet.caption = "Tileset: " & 1
    frmEditor_Map.scrlTileSet.value = 1
    ' set the scrollbars
    frmEditor_Map.scrlPictureY.max = (frmEditor_Map.picBackSelect.Height \ PIC_Y) - (frmEditor_Map.picBack.Height \ PIC_Y)
    frmEditor_Map.scrlPictureX.max = (frmEditor_Map.picBackSelect.Width \ PIC_X) - (frmEditor_Map.picBack.Width \ PIC_X)
    shpSelectedWidth = 32
    shpSelectedHeight = 32
    MapEditorTileScroll
    ' set shops for the shop attribute
    frmEditor_Map.cmbShop.AddItem "None"

    For i = 1 To MAX_SHOPS
        frmEditor_Map.cmbShop.AddItem i & ": " & Shop(i).Name
    Next

    ' we're not in a shop
    frmEditor_Map.cmbShop.ListIndex = 0
End Sub

Public Sub MapEditorProperties()
    Dim X As Long, i As Long, tmpNum As Long

    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If

    ' add the array to the combo
    frmEditor_MapProperties.lstMusic.Clear
    frmEditor_MapProperties.lstMusic.AddItem "None."
    tmpNum = UBound(musicCache)

    For i = 1 To tmpNum
        frmEditor_MapProperties.lstMusic.AddItem musicCache(i)
    Next

    ' finished populating
    With frmEditor_MapProperties
        .scrlBoss.max = MAX_MAP_NPCS
        .txtName.text = Trim$(Map.MapData.Name)

        ' find the music we have set
        If .lstMusic.ListCount >= 0 Then
            .lstMusic.ListIndex = 0
            tmpNum = .lstMusic.ListCount

            For i = 0 To tmpNum - 1

                If .lstMusic.list(i) = Trim$(Map.MapData.Music) Then
                    .lstMusic.ListIndex = i
                End If

            Next

        End If

        ' rest of it
        .txtUp.text = CStr(Map.MapData.Up)
        .txtDown.text = CStr(Map.MapData.Down)
        .txtLeft.text = CStr(Map.MapData.Left)
        .txtRight.text = CStr(Map.MapData.Right)
        .cmbMoral.ListIndex = Map.MapData.Moral
        .txtBootMap.text = CStr(Map.MapData.BootMap)
        .txtBootX.text = CStr(Map.MapData.BootX)
        .txtBootY.text = CStr(Map.MapData.BootY)
        .CmbWeather.ListIndex = Map.MapData.Weather
        .scrlWeatherIntensity.value = Map.MapData.WeatherIntensity
        
        .ScrlFog.value = Map.MapData.Fog
        .ScrlFogSpeed.value = Map.MapData.FogSpeed
        .scrlFogOpacity.value = Map.MapData.FogOpacity
        
        .scrlRed.value = Map.MapData.Red
        .scrlGreen.value = Map.MapData.Green
        .scrlBlue.value = Map.MapData.Blue
        .scrlAlpha.value = Map.MapData.alpha
        .scrlBoss = Map.MapData.BossNpc
        ' show the map npcs
        .lstNpcs.Clear

        For X = 1 To MAX_MAP_NPCS

            If Map.MapData.Npc(X) > 0 Then
                .lstNpcs.AddItem X & ": " & Trim$(Npc(Map.MapData.Npc(X)).Name)
            Else
                .lstNpcs.AddItem X & ": No NPC"
            End If

        Next

        .lstNpcs.ListIndex = 0
        ' show the npc selection combo
        .cmbNpc.Clear
        .cmbNpc.AddItem "No NPC"

        For X = 1 To MAX_NPCS
            .cmbNpc.AddItem X & ": " & Trim$(Npc(X).Name)
        Next

        ' set the combo box properly
        Dim tmpString() As String
        Dim NpcNum As Long
        tmpString = Split(.lstNpcs.list(.lstNpcs.ListIndex))
        NpcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        .cmbNpc.ListIndex = Map.MapData.Npc(NpcNum)
        ' show the current map
        .lblMap.caption = "Current map: " & GetPlayerMap(MyIndex)
        .txtMaxX.text = Map.MapData.MaxX
        .txtMaxY.text = Map.MapData.MaxY
    End With

End Sub

Public Sub MapEditorSetTile(ByVal X As Long, ByVal Y As Long, ByVal CurLayer As Long, Optional ByVal multitile As Boolean = False, Optional ByVal theAutotile As Byte = 0)
    Dim x2 As Long, y2 As Long

    If theAutotile > 0 Then
        With Map.TileData.Tile(X, Y)
            ' set layer
            .Layer(CurLayer).X = EditorTileX
            .Layer(CurLayer).Y = EditorTileY
            .Layer(CurLayer).tileSet = frmEditor_Map.scrlTileSet.value
            .Autotile(CurLayer) = theAutotile
            cacheRenderState X, Y, CurLayer
        End With
        ' do a re-init so we can see our changes
        initAutotiles
        Exit Sub
    End If

    If Not multitile Then ' single
        With Map.TileData.Tile(X, Y)
            ' set layer
            .Layer(CurLayer).X = EditorTileX
            .Layer(CurLayer).Y = EditorTileY
            .Layer(CurLayer).tileSet = frmEditor_Map.scrlTileSet.value
            .Autotile(CurLayer) = 0
            cacheRenderState X, Y, CurLayer
        End With
    Else ' multitile
        y2 = 0 ' starting tile for y axis
        For Y = CurY To CurY + EditorTileHeight - 1
            x2 = 0 ' re-set x count every y loop
            For X = CurX To CurX + EditorTileWidth - 1
                If X >= 0 And X <= Map.MapData.MaxX Then
                    If Y >= 0 And Y <= Map.MapData.MaxY Then
                        With Map.TileData.Tile(X, Y)
                            .Layer(CurLayer).X = EditorTileX + x2
                            .Layer(CurLayer).Y = EditorTileY + y2
                            .Layer(CurLayer).tileSet = frmEditor_Map.scrlTileSet.value
                            .Autotile(CurLayer) = 0
                            cacheRenderState X, Y, CurLayer
                        End With
                    End If
                End If
                x2 = x2 + 1
            Next
            y2 = y2 + 1
        Next
    End If

End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal movedMouse As Boolean = True)
    Dim i As Long
    Dim CurLayer As Long

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1

        If frmEditor_Map.optLayer(i).value Then
            CurLayer = i
            Exit For
        End If

    Next

    If Not isInBounds Then Exit Sub
    If Button = vbLeftButton Then
        If frmEditor_Map.optLayers.value Then

            ' no autotiling
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
                MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.value
            Else ' multi tile!

                If frmEditor_Map.scrlAutotile.value = 0 Then
                    MapEditorSetTile CurX, CurY, CurLayer, True
                Else
                    MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.value
                End If
            End If

        ElseIf frmEditor_Map.optAttribs.value Then

            With Map.TileData.Tile(CurX, CurY)

                ' blocked tile
                If frmEditor_Map.optBlocked.value Then .Type = TILE_TYPE_BLOCKED

                ' warp tile
                If frmEditor_Map.optWarp.value Then
                    .Type = TILE_TYPE_WARP
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = EditorWarpFall
                    .Data5 = 0
                End If

                ' item spawn
                If frmEditor_Map.optItem.value Then
                    .Type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' npc avoid
                If frmEditor_Map.optNpcAvoid.value Then
                    .Type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' key
                If frmEditor_Map.optKey.value Then
                    .Type = TILE_TYPE_KEY
                    .Data1 = KeyEditorNum
                    .Data2 = KeyEditorTake
                    .Data3 = KeyEditorTime
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' key open
                If frmEditor_Map.optKeyOpen.value Then
                    .Type = TILE_TYPE_KEYOPEN
                    .Data1 = KeyOpenEditorX
                    .Data2 = KeyOpenEditorY
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' resource
                If frmEditor_Map.optResource.value Then
                    .Type = TILE_TYPE_RESOURCE
                    .Data1 = ResourceEditorNum
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' door
                If frmEditor_Map.optDoor.value Then
                    .Type = TILE_TYPE_DOOR
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' npc spawn
                If frmEditor_Map.optNpcSpawn.value Then
                    .Type = TILE_TYPE_NPCSPAWN
                    .Data1 = SpawnNpcNum
                    .Data2 = SpawnNpcDir
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' shop
                If frmEditor_Map.optShop.value Then
                    .Type = TILE_TYPE_SHOP
                    .Data1 = EditorShop
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' bank
                If frmEditor_Map.optBank.value Then
                    .Type = TILE_TYPE_BANK
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' heal
                If frmEditor_Map.optHeal.value Then
                    .Type = TILE_TYPE_HEAL
                    .Data1 = MapEditorHealType
                    .Data2 = MapEditorHealAmount
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' trap
                If frmEditor_Map.optTrap.value Then
                    .Type = TILE_TYPE_TRAP
                    .Data1 = MapEditorHealAmount
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' slide
                If frmEditor_Map.optSlide.value Then
                    .Type = TILE_TYPE_SLIDE
                    .Data1 = MapEditorSlideDir
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' chat
                If frmEditor_Map.optChat.value Then
                    .Type = TILE_TYPE_CHAT
                    .Data1 = MapEditorChatNpc
                    .Data2 = MapEditorChatDir
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If
                
                ' appear
                If frmEditor_Map.optAppear.value Then
                    .Type = TILE_TYPE_APPEAR
                    .Data1 = EditorAppearRange
                    .Data2 = EditorAppearBottom
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If
            End With

        ElseIf frmEditor_Map.optBlock.value Then

            If movedMouse Then Exit Sub
            ' find what tile it is
            X = X - ((X \ 32) * 32)
            Y = Y - ((Y \ 32) * 32)

            ' see if it hits an arrow
            For i = 1 To 4
                If X >= DirArrowX(i) And X <= DirArrowX(i) + 8 Then
                    If Y >= DirArrowY(i) And Y <= DirArrowY(i) + 8 Then
                        ' flip the value.
                        setDirBlock Map.TileData.Tile(CurX, CurY).DirBlock, CByte(i), Not isDirBlocked(Map.TileData.Tile(CurX, CurY).DirBlock, CByte(i))
                        Exit Sub
                    End If
                End If
            Next
        End If
    End If

    If Button = vbRightButton Then
        If frmEditor_Map.optLayers.value Then

            With Map.TileData.Tile(CurX, CurY)
                ' clear layer
                .Layer(CurLayer).X = 0
                .Layer(CurLayer).Y = 0
                .Layer(CurLayer).tileSet = 0

                If .Autotile(CurLayer) > 0 Then
                    .Autotile(CurLayer) = 0
                    ' do a re-init so we can see our changes
                    initAutotiles
                End If

                cacheRenderState X, Y, CurLayer
            End With

        ElseIf frmEditor_Map.optAttribs.value Then

            With Map.TileData.Tile(CurX, CurY)
                ' clear attribute
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
                .Data4 = 0
                .Data5 = 0
            End With

        End If
    End If

    CacheResources
End Sub

Public Sub MapEditorChooseTile(Button As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
        EditorTileX = X \ PIC_X
        EditorTileY = Y \ PIC_Y
        shpSelectedTop = EditorTileY * PIC_Y
        shpSelectedLeft = EditorTileX * PIC_X
        shpSelectedWidth = PIC_X
        shpSelectedHeight = PIC_Y
    End If

End Sub

Public Sub MapEditorDrag(Button As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        ' convert the pixel number to tile number
        X = (X \ PIC_X) + 1
        Y = (Y \ PIC_Y) + 1

        ' check it's not out of bounds
        If X < 0 Then X = 0
        If X > frmEditor_Map.picBackSelect.Width / PIC_X Then X = frmEditor_Map.picBackSelect.Width / PIC_X
        If Y < 0 Then Y = 0
        If Y > frmEditor_Map.picBackSelect.Height / PIC_Y Then Y = frmEditor_Map.picBackSelect.Height / PIC_Y

        ' find out what to set the width + height of map editor to
        If X > EditorTileX Then ' drag right
            EditorTileWidth = X - EditorTileX
        Else ' drag left
            ' TO DO
        End If

        If Y > EditorTileY Then ' drag down
            EditorTileHeight = Y - EditorTileY
        Else ' drag up
            ' TO DO
        End If

        shpSelectedWidth = EditorTileWidth * PIC_X
        shpSelectedHeight = EditorTileHeight * PIC_Y
    End If

End Sub

Public Sub NudgeMap(ByVal theDir As Byte)
Dim X As Long, Y As Long, i As Long
    
    ' if left or right
    If theDir = DIR_UP Or theDir = DIR_LEFT Then
        For Y = 0 To Map.MapData.MaxY
            For X = 0 To Map.MapData.MaxX
                Select Case theDir
                    Case DIR_UP
                        ' move up all one
                        If Y > 0 Then CopyTile Map.TileData.Tile(X, Y), Map.TileData.Tile(X, Y - 1)
                    Case DIR_LEFT
                        ' move left all one
                        If X > 0 Then CopyTile Map.TileData.Tile(X, Y), Map.TileData.Tile(X - 1, Y)
                End Select
            Next
        Next
    Else
        For Y = Map.MapData.MaxY To 0 Step -1
            For X = Map.MapData.MaxX To 0 Step -1
                Select Case theDir
                    Case DIR_DOWN
                        ' move down all one
                        If Y < Map.MapData.MaxY Then CopyTile Map.TileData.Tile(X, Y), Map.TileData.Tile(X, Y + 1)
                    Case DIR_RIGHT
                        ' move right all one
                        If X < Map.MapData.MaxX Then CopyTile Map.TileData.Tile(X, Y), Map.TileData.Tile(X + 1, Y)
                End Select
            Next
        Next
    End If
    
    ' do events
    If Map.TileData.EventCount > 0 Then
        For i = 1 To Map.TileData.EventCount
            Select Case theDir
                Case DIR_UP
                    Map.TileData.Events(i).Y = Map.TileData.Events(i).Y - 1
                Case DIR_LEFT
                    Map.TileData.Events(i).X = Map.TileData.Events(i).X - 1
                Case DIR_RIGHT
                    Map.TileData.Events(i).X = Map.TileData.Events(i).X + 1
                Case DIR_DOWN
                    Map.TileData.Events(i).Y = Map.TileData.Events(i).Y + 1
            End Select
        Next
    End If
    
    initAutotiles
End Sub

Public Sub CopyTile(ByRef origTile As TileRec, ByRef newTile As TileRec)
Dim tilesize As Long
    tilesize = LenB(origTile)
    CopyMemory ByVal VarPtr(newTile), ByVal VarPtr(origTile), tilesize
    ZeroMemory ByVal VarPtr(origTile), tilesize
End Sub

Public Sub MapEditorTileScroll()

    ' horizontal scrolling
    If frmEditor_Map.picBackSelect.Width < frmEditor_Map.picBack.Width Then
        frmEditor_Map.scrlPictureX.enabled = False
    Else
        frmEditor_Map.scrlPictureX.enabled = True
        frmEditor_Map.picBackSelect.Left = (frmEditor_Map.scrlPictureX.value * PIC_X) * -1
    End If

    ' vertical scrolling
    If frmEditor_Map.picBackSelect.Height < frmEditor_Map.picBack.Height Then
        frmEditor_Map.scrlPictureY.enabled = False
    Else
        frmEditor_Map.scrlPictureY.enabled = True
        frmEditor_Map.picBackSelect.Top = (frmEditor_Map.scrlPictureY.value * PIC_Y) * -1
    End If

End Sub

Public Sub MapEditorSend()
    Call SendMap
    InMapEditor = False
    'Unload frmEditor_Map
    frmEditor_Map.Hide
End Sub

Public Sub MapEditorCancel()
    InMapEditor = False
    LoadMap GetPlayerMap(MyIndex)
    initAutotiles
    'Unload frmEditor_Map
    frmEditor_Map.Hide
End Sub

Public Sub MapEditorClearLayer()
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim CurLayer As Long

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1

        If frmEditor_Map.optLayer(i).value Then
            CurLayer = i
            Exit For
        End If

    Next

    If CurLayer = 0 Then Exit Sub

    ' ask to clear layer
    If MsgBox("Are you sure you wish to clear this layer?", vbYesNo, GAME_NAME) = vbYes Then

        For X = 0 To Map.MapData.MaxX
            For Y = 0 To Map.MapData.MaxY
                Map.TileData.Tile(X, Y).Layer(CurLayer).X = 0
                Map.TileData.Tile(X, Y).Layer(CurLayer).Y = 0
                Map.TileData.Tile(X, Y).Layer(CurLayer).tileSet = 0
                cacheRenderState X, Y, CurLayer
            Next
        Next

        ' re-cache autos
        initAutotiles
    End If

End Sub

Public Sub MapEditorFillLayer()
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim CurLayer As Long

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1

        If frmEditor_Map.optLayer(i).value Then
            CurLayer = i
            Exit For
        End If

    Next

    ' Ground layer
    If MsgBox("Are you sure you wish to fill this layer?", vbYesNo, GAME_NAME) = vbYes Then

        For X = 0 To Map.MapData.MaxX
            For Y = 0 To Map.MapData.MaxY
                Map.TileData.Tile(X, Y).Layer(CurLayer).X = EditorTileX
                Map.TileData.Tile(X, Y).Layer(CurLayer).Y = EditorTileY
                Map.TileData.Tile(X, Y).Layer(CurLayer).tileSet = frmEditor_Map.scrlTileSet.value
                Map.TileData.Tile(X, Y).Autotile(CurLayer) = frmEditor_Map.scrlAutotile.value
                cacheRenderState X, Y, CurLayer
            Next
        Next

        ' now cache the positions
        initAutotiles
    End If

End Sub

Public Sub MapEditorClearAttribs()
    Dim X As Long
    Dim Y As Long

    If MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, GAME_NAME) = vbYes Then

        For X = 0 To Map.MapData.MaxX
            For Y = 0 To Map.MapData.MaxY
                Map.TileData.Tile(X, Y).Type = 0
            Next
        Next

    End If

End Sub

Public Sub MapEditorLeaveMap()

    If InMapEditor Then
        If MsgBox("Save changes to current map?", vbYesNo) = vbYes Then
            Call MapEditorSend
        Else
            Call MapEditorCancel
        End If
    End If

End Sub

' /////////////////
' // Item Editor //
' /////////////////
Public Sub ItemEditorInit()
    Dim i As Long, SoundSet As Boolean, tmpNum As Long

    If frmEditor_Item.visible = False Then Exit Sub
    EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1

    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If

    ' add the array to the combo
    frmEditor_Item.cmbSound.Clear
    frmEditor_Item.cmbSound.AddItem "None."
    tmpNum = UBound(soundCache)

    For i = 1 To tmpNum
        frmEditor_Item.cmbSound.AddItem soundCache(i)
    Next

    ' finished populating
    With Item(EditorIndex)
        frmEditor_Item.txtName.text = Trim$(.Name)

        If .Pic > frmEditor_Item.scrlPic.max Then .Pic = 0
        frmEditor_Item.scrlPic.value = .Pic
        frmEditor_Item.cmbType.ListIndex = .Type
        frmEditor_Item.scrlAnim.value = .Animation
        frmEditor_Item.txtDesc.text = Trim$(.Desc)

        ' find the sound we have set
        If frmEditor_Item.cmbSound.ListCount >= 0 Then
            tmpNum = frmEditor_Item.cmbSound.ListCount

            For i = 0 To tmpNum

                If frmEditor_Item.cmbSound.list(i) = Trim$(.sound) Then
                    frmEditor_Item.cmbSound.ListIndex = i
                    SoundSet = True
                End If

            Next

            If Not SoundSet Or frmEditor_Item.cmbSound.ListIndex = -1 Then frmEditor_Item.cmbSound.ListIndex = 0
        End If

        ' Type specific settings
        If (frmEditor_Item.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
            frmEditor_Item.fraEquipment.visible = True
            frmEditor_Item.scrlDamage.value = .Data2
            frmEditor_Item.cmbTool.ListIndex = .Data3

            If .speed < 100 Then .speed = 100
            frmEditor_Item.scrlSpeed.value = .speed

            ' loop for stats
            For i = 1 To Stats.Stat_Count - 1
                frmEditor_Item.scrlStatBonus(i).value = .Add_Stat(i)
            Next

            If Not .Paperdoll > Count_Paperdoll Then frmEditor_Item.scrlPaperdoll = .Paperdoll
            frmEditor_Item.scrlProf.value = .proficiency
        Else
            frmEditor_Item.fraEquipment.visible = False
        End If

        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_CONSUME Then
            frmEditor_Item.fraVitals.visible = True
            frmEditor_Item.scrlAddHp.value = .AddHP
            frmEditor_Item.scrlAddMP.value = .AddMP
            frmEditor_Item.scrlAddExp.value = .AddEXP
            frmEditor_Item.scrlCastSpell.value = .CastSpell
            frmEditor_Item.chkInstant.value = .instaCast
        Else
            frmEditor_Item.fraVitals.visible = False
        End If

        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            frmEditor_Item.fraSpell.visible = True
            frmEditor_Item.scrlSpell.value = .Data1
        Else
            frmEditor_Item.fraSpell.visible = False
        End If

        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_FOOD Then
            If .HPorSP = 2 Then
                frmEditor_Item.optSP.value = True
            Else
                frmEditor_Item.optHP.value = True
            End If

            frmEditor_Item.scrlFoodHeal = .FoodPerTick
            frmEditor_Item.scrlFoodTick = .FoodTickCount
            frmEditor_Item.scrlFoodInterval = .FoodInterval
            frmEditor_Item.fraFood.visible = True
        Else
            frmEditor_Item.fraFood.visible = False
        End If

        ' Basic requirements
        frmEditor_Item.scrlAccessReq.value = .AccessReq
        frmEditor_Item.scrlLevelReq.value = .LevelReq

        ' loop for stats
        For i = 1 To Stats.Stat_Count - 1
            frmEditor_Item.scrlStatReq(i).value = .Stat_Req(i)
        Next

        ' Build cmbClassReq
        frmEditor_Item.cmbClassReq.Clear
        frmEditor_Item.cmbClassReq.AddItem "None"

        For i = 1 To Max_Classes
            frmEditor_Item.cmbClassReq.AddItem Class(i).Name
        Next

        frmEditor_Item.cmbClassReq.ListIndex = .ClassReq
        ' Info
        frmEditor_Item.scrlPrice.value = .Price
        frmEditor_Item.cmbBind.ListIndex = .BindType
        frmEditor_Item.scrlRarity.value = .Rarity
        EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    End With

    Item_Changed(EditorIndex) = True
End Sub

Public Sub ItemEditorOk()
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Item_Changed(i) Then
            Call SendSaveItem(i)
        End If

    Next

    Unload frmEditor_Item
    Editor = 0
    ClearChanged_Item
End Sub

Sub ItemEditorCopy()
    CopyMemory ByVal VarPtr(tmpItem), ByVal VarPtr(Item(EditorIndex)), LenB(Item(EditorIndex))
End Sub

Sub ItemEditorPaste()
    CopyMemory ByVal VarPtr(Item(EditorIndex)), ByVal VarPtr(tmpItem), LenB(tmpItem)
    ItemEditorInit
    frmEditor_Item.txtName_Validate False
End Sub

Public Sub ItemEditorCancel()
    Editor = 0
    Unload frmEditor_Item
    ClearChanged_Item
    ClearItems
    SendRequestItems
End Sub

Public Sub ClearChanged_Item()
    ZeroMemory Item_Changed(1), MAX_ITEMS * 2 ' 2 = boolean length
End Sub

' /////////////////
' // Conv Editor //
' /////////////////
Public Sub ConvEditorInit()
    Dim i As Long, N As Long

    If frmEditor_Conv.visible = False Then Exit Sub
    EditorIndex = frmEditor_Conv.lstIndex.ListIndex + 1

    With frmEditor_Conv
        .txtName.text = Trim$(Conv(EditorIndex).Name)

        If Conv(EditorIndex).chatCount = 0 Then
            Conv(EditorIndex).chatCount = 1
            ReDim Conv(EditorIndex).Conv(1 To Conv(EditorIndex).chatCount)
        End If

        For N = 1 To 4
            .cmbReply(N).Clear
            .cmbReply(N).AddItem "None"

            For i = 1 To Conv(EditorIndex).chatCount
                .cmbReply(N).AddItem i
            Next
        Next

        .scrlChatCount = Conv(EditorIndex).chatCount
        .scrlConv.max = Conv(EditorIndex).chatCount
        .scrlConv.value = 1
        .txtConv = Conv(EditorIndex).Conv(.scrlConv.value).Conv

        For i = 1 To 4
            .txtReply(i).text = Conv(EditorIndex).Conv(.scrlConv.value).rText(i)
            .cmbReply(i).ListIndex = Conv(EditorIndex).Conv(.scrlConv.value).rTarget(i)
        Next

        .cmbEvent.ListIndex = Conv(EditorIndex).Conv(.scrlConv.value).Event
        .scrlData1.value = Conv(EditorIndex).Conv(.scrlConv.value).Data1
        .scrlData2.value = Conv(EditorIndex).Conv(.scrlConv.value).Data2
        .scrlData3.value = Conv(EditorIndex).Conv(.scrlConv.value).Data3
    End With

    Conv_Changed(EditorIndex) = True
End Sub

Public Sub ConvEditorOk()
    Dim i As Long

    For i = 1 To MAX_CONVS

        If Conv_Changed(i) Then
            Call SendSaveConv(i)
        End If

    Next

    Unload frmEditor_Conv
    Editor = 0
    ClearChanged_Conv
End Sub

Public Sub ConvEditorCancel()
    Editor = 0
    Unload frmEditor_Conv
    ClearChanged_Conv
    ClearConvs
    SendRequestConvs
End Sub

Public Sub ClearChanged_Conv()
    ZeroMemory Conv_Changed(1), MAX_CONVS * 2 ' 2 = boolean length
End Sub

Public Sub ClearAttributeDialogue()
    frmEditor_Map.fraNpcSpawn.visible = False
    frmEditor_Map.fraResource.visible = False
    frmEditor_Map.fraMapItem.visible = False
    frmEditor_Map.fraMapKey.visible = False
    frmEditor_Map.fraKeyOpen.visible = False
    frmEditor_Map.fraMapWarp.visible = False
    frmEditor_Map.fraShop.visible = False
End Sub
