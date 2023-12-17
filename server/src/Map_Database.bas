Attribute VB_Name = "Map_Database"
' **********
' ** Maps **
' **********
Public Sub SaveMap(ByVal mapnum As Long)
    Dim filename As String, f As Long, x As Long, y As Long, i As Long
    
    ' save map data
    filename = App.Path & "\data\maps\map" & mapnum & ".ini"
    
    ' if it exists then kill the ini
    If FileExist(filename) Then Kill filename
    
    ' General
    With Map(mapnum).MapData
        PutVar filename, "General", "Name", .Name
        PutVar filename, "General", "Music", .Music
        PutVar filename, "General", "Moral", Val(.Moral)
        PutVar filename, "General", "Up", Val(.Up)
        PutVar filename, "General", "Down", Val(.Down)
        PutVar filename, "General", "Left", Val(.left)
        PutVar filename, "General", "Right", Val(.Right)
        PutVar filename, "General", "BootMap", Val(.BootMap)
        PutVar filename, "General", "BootX", Val(.BootX)
        PutVar filename, "General", "BootY", Val(.BootY)
        PutVar filename, "General", "MaxX", Val(.MaxX)
        PutVar filename, "General", "MaxY", Val(.MaxY)
        
        PutVar filename, "General", "Weather", Val(.Weather)
        PutVar filename, "General", "WeatherIntensity", Val(.WeatherIntensity)
        
        PutVar filename, "General", "Fog", Val(.Fog)
        PutVar filename, "General", "FogSpeed", Val(.FogSpeed)
        PutVar filename, "General", "FogOpacity", Val(.FogOpacity)
        
        PutVar filename, "General", "Red", Val(.Red)
        PutVar filename, "General", "Green", Val(.Green)
        PutVar filename, "General", "Blue", Val(.Blue)
        PutVar filename, "General", "Alpha", Val(.Alpha)
        
        PutVar filename, "General", "BossNpc", Val(.BossNpc)
        For i = 1 To MAX_MAP_NPCS
            PutVar filename, "General", "Npc" & i, Val(.Npc(i))
        Next
    End With
    
    ' dump tile data
    filename = App.Path & "\data\maps\map" & mapnum & ".dat"
    f = FreeFile
    
    With Map(mapnum)
        Open filename For Binary As #f
            For x = 0 To .MapData.MaxX
                For y = 0 To .MapData.MaxY
                    Put #f, , .TileData.Tile(x, y).Type
                    Put #f, , .TileData.Tile(x, y).Data1
                    Put #f, , .TileData.Tile(x, y).Data2
                    Put #f, , .TileData.Tile(x, y).Data3
                    Put #f, , .TileData.Tile(x, y).Data4
                    Put #f, , .TileData.Tile(x, y).Data5
                    Put #f, , .TileData.Tile(x, y).Autotile
                    Put #f, , .TileData.Tile(x, y).DirBlock
                    For i = 1 To MapLayer.Layer_Count - 1
                        Put #f, , .TileData.Tile(x, y).Layer(i).Tileset
                        Put #f, , .TileData.Tile(x, y).Layer(i).x
                        Put #f, , .TileData.Tile(x, y).Layer(i).y
                    Next
                Next
            Next
        Close #f
    End With

    DoEvents
End Sub

Public Sub SaveMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next

End Sub

Public Sub CheckMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS

        If Not FileExist(App.Path & "\Data\maps\map" & i & ".dat") Or Not FileExist(App.Path & "\Data\maps\map" & i & ".ini") Then
            Call SaveMap(i)
        End If

    Next

End Sub

Public Sub LoadMap(mapnum As Long)
    Dim filename As String, i As Long, f As Long, x As Long, y As Long
    
    ' load map data
    filename = App.Path & "\data\maps\map" & mapnum & ".ini"
    
    ' General
    With Map(mapnum).MapData
        .Name = GetVar(filename, "General", "Name")
        .Music = GetVar(filename, "General", "Music")
        .Moral = Val(GetVar(filename, "General", "Moral"))
        .Up = Val(GetVar(filename, "General", "Up"))
        .Down = Val(GetVar(filename, "General", "Down"))
        .left = Val(GetVar(filename, "General", "Left"))
        .Right = Val(GetVar(filename, "General", "Right"))
        .BootMap = Val(GetVar(filename, "General", "BootMap"))
        .BootX = Val(GetVar(filename, "General", "BootX"))
        .BootY = Val(GetVar(filename, "General", "BootY"))
        .MaxX = Val(GetVar(filename, "General", "MaxX"))
        .MaxY = Val(GetVar(filename, "General", "MaxY"))
        
        .Weather = Val(GetVar(filename, "General", "Weather"))
        .WeatherIntensity = Val(GetVar(filename, "General", "WeatherIntensity"))
        
        .Fog = Val(GetVar(filename, "General", "Fog"))
        .FogSpeed = Val(GetVar(filename, "General", "FogSpeed"))
        .FogOpacity = Val(GetVar(filename, "General", "FogOpacity"))
        
        .Red = Val(GetVar(filename, "General", "Red"))
        .Green = Val(GetVar(filename, "General", "Green"))
        .Blue = Val(GetVar(filename, "General", "Blue"))
        .Alpha = Val(GetVar(filename, "General", "Alpha"))
        
        .BossNpc = Val(GetVar(filename, "General", "BossNpc"))
        For i = 1 To MAX_MAP_NPCS
            .Npc(i) = Val(GetVar(filename, "General", "Npc" & i))
        Next
    End With
        
    ' dump tile data
    filename = App.Path & "\data\maps\map" & mapnum & ".dat"
    f = FreeFile
    
    ' redim the map
    ReDim Map(mapnum).TileData.Tile(0 To Map(mapnum).MapData.MaxX, 0 To Map(mapnum).MapData.MaxY) As TileRec
    
    With Map(mapnum)
        Open filename For Binary As #f
            For x = 0 To .MapData.MaxX
                For y = 0 To .MapData.MaxY
                    Get #f, , .TileData.Tile(x, y).Type
                    Get #f, , .TileData.Tile(x, y).Data1
                    Get #f, , .TileData.Tile(x, y).Data2
                    Get #f, , .TileData.Tile(x, y).Data3
                    Get #f, , .TileData.Tile(x, y).Data4
                    Get #f, , .TileData.Tile(x, y).Data5
                    Get #f, , .TileData.Tile(x, y).Autotile
                    Get #f, , .TileData.Tile(x, y).DirBlock
                    For i = 1 To MapLayer.Layer_Count - 1
                        Get #f, , .TileData.Tile(x, y).Layer(i).Tileset
                        Get #f, , .TileData.Tile(x, y).Layer(i).x
                        Get #f, , .TileData.Tile(x, y).Layer(i).y
                    Next
                Next
            Next
        Close #f
    End With
End Sub

Public Sub LoadMaps()
    Dim filename As String, mapnum As Long

    Call CheckMaps

    For mapnum = 1 To MAX_MAPS
        LoadMap mapnum
        ClearTempTile mapnum
        CacheResources mapnum
        DoEvents
    Next
End Sub

Public Sub ClearMap(ByVal mapnum As Long)
    Map(mapnum) = EmptyMap
    Map(mapnum).MapData.Name = vbNullString
    Map(mapnum).MapData.MaxX = MAX_MAPX
    Map(mapnum).MapData.MaxY = MAX_MAPY
    ReDim Map(mapnum).TileData.Tile(0 To Map(mapnum).MapData.MaxX, 0 To Map(mapnum).MapData.MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(mapnum) = NO
    ' Reset the map cache array for this map.
    MapCache(mapnum).Data = vbNullString
End Sub

Public Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next
End Sub

Public Sub ClearMapItem(ByVal Index As Long, ByVal mapnum As Long)
    MapItem(mapnum, Index) = EmptyMapItem
    MapItem(mapnum, Index).playerName = vbNullString
End Sub

Public Sub ClearMapItems()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next

End Sub

Public Sub ClearMapNpc(ByVal Index As Long, ByVal mapnum As Long)
    MapNpc(mapnum) = EmptyMapNpc
End Sub

Public Sub ClearMapNpcs()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next

End Sub

Public Sub GetMapCRC32(mapnum As Long)
    Dim Data() As Byte, filename As String, f As Long
    ' map data
    filename = App.Path & "\data\maps\map" & mapnum & ".ini"
    If FileExist(filename) Then
        f = FreeFile
        Open filename For Binary As #f
        Data = Space$(LOF(f))
        Get #f, , Data
        Close #f
        MapCRC32(mapnum).MapDataCRC = CRC32(Data)
    Else
        MapCRC32(mapnum).MapDataCRC = 0
    End If
    ' clear
    Erase Data
    ' tile data
    filename = App.Path & "\data\maps\map" & mapnum & ".dat"
    If FileExist(filename) Then
        f = FreeFile
        Open filename For Binary As #f
        Data = Space$(LOF(f))
        Get #f, , Data
        Close #f
        MapCRC32(mapnum).MapTileCRC = CRC32(Data)
    Else
        MapCRC32(mapnum).MapTileCRC = 0
    End If
End Sub

Public Sub ClearProjectile(IndexProjectile As Long)
    If IndexProjectile < 0 Or IndexProjectile > MAX_PROJECTILE_MAP Then Exit Sub
    MapProjectile(IndexProjectile) = EmptyMapProjectile
End Sub

