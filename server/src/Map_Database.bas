Attribute VB_Name = "Map_Database"
' **********
' ** Maps **
' **********
Public Sub SaveMap(ByVal MapNum As Long)
    Dim filename As String, F As Long, x As Long, y As Long, i As Long
    
    ' save map data
    filename = App.Path & "\data\maps\map" & MapNum & ".ini"
    
    ' if it exists then kill the ini
    If FileExist(filename) Then Kill filename
    
    ' General
    With Map(MapNum).MapData
        PutVar filename, "General", "Name", .name
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
    filename = App.Path & "\data\maps\map" & MapNum & ".dat"
    F = FreeFile
    
    With Map(MapNum)
        Open filename For Binary As #F
            For x = 0 To .MapData.MaxX
                For y = 0 To .MapData.MaxY
                    Put #F, , .TileData.Tile(x, y).Type
                    Put #F, , .TileData.Tile(x, y).Data1
                    Put #F, , .TileData.Tile(x, y).Data2
                    Put #F, , .TileData.Tile(x, y).Data3
                    Put #F, , .TileData.Tile(x, y).Data4
                    Put #F, , .TileData.Tile(x, y).Data5
                    Put #F, , .TileData.Tile(x, y).Autotile
                    Put #F, , .TileData.Tile(x, y).DirBlock
                    For i = 1 To MapLayer.Layer_Count - 1
                        Put #F, , .TileData.Tile(x, y).Layer(i).Tileset
                        Put #F, , .TileData.Tile(x, y).Layer(i).x
                        Put #F, , .TileData.Tile(x, y).Layer(i).y
                    Next
                Next
            Next
        Close #F
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

Public Sub LoadMap(MapNum As Long)
    Dim filename As String, i As Long, F As Long, x As Long, y As Long
    
    ' load map data
    filename = App.Path & "\data\maps\map" & MapNum & ".ini"
    
    ' General
    With Map(MapNum).MapData
        .name = GetVar(filename, "General", "Name")
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
    filename = App.Path & "\data\maps\map" & MapNum & ".dat"
    F = FreeFile
    
    ' redim the map
    ReDim Map(MapNum).TileData.Tile(0 To Map(MapNum).MapData.MaxX, 0 To Map(MapNum).MapData.MaxY) As TileRec
    
    With Map(MapNum)
        Open filename For Binary As #F
            For x = 0 To .MapData.MaxX
                For y = 0 To .MapData.MaxY
                    Get #F, , .TileData.Tile(x, y).Type
                    Get #F, , .TileData.Tile(x, y).Data1
                    Get #F, , .TileData.Tile(x, y).Data2
                    Get #F, , .TileData.Tile(x, y).Data3
                    Get #F, , .TileData.Tile(x, y).Data4
                    Get #F, , .TileData.Tile(x, y).Data5
                    Get #F, , .TileData.Tile(x, y).Autotile
                    Get #F, , .TileData.Tile(x, y).DirBlock
                    For i = 1 To MapLayer.Layer_Count - 1
                        Get #F, , .TileData.Tile(x, y).Layer(i).Tileset
                        Get #F, , .TileData.Tile(x, y).Layer(i).x
                        Get #F, , .TileData.Tile(x, y).Layer(i).y
                    Next
                Next
            Next
        Close #F
    End With
End Sub

Public Sub LoadMaps()
    Dim filename As String, MapNum As Long

    Call CheckMaps

    For MapNum = 1 To MAX_MAPS
        LoadMap MapNum
        ClearTempTile MapNum
        CacheResources MapNum
        DoEvents
    Next
End Sub

Public Sub ClearMap(ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    Map(MapNum).MapData.name = vbNullString
    Map(MapNum).MapData.MaxX = MAX_MAPX
    Map(MapNum).MapData.MaxY = MAX_MAPY
    ReDim Map(MapNum).TileData.Tile(0 To Map(MapNum).MapData.MaxX, 0 To Map(MapNum).MapData.MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
    ' Reset the map cache array for this map.
    MapCache(MapNum).Data = vbNullString
End Sub

Public Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next
End Sub

Public Sub ClearMapItem(ByVal index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(MapNum, index)), LenB(MapItem(MapNum, index)))
    MapItem(MapNum, index).playerName = vbNullString
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

Public Sub ClearMapNpc(ByVal index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum).Npc(index)), LenB(MapNpc(MapNum).Npc(index)))
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

Public Sub GetMapCRC32(MapNum As Long)
    Dim Data() As Byte, filename As String, F As Long
    ' map data
    filename = App.Path & "\data\maps\map" & MapNum & ".ini"
    If FileExist(filename) Then
        F = FreeFile
        Open filename For Binary As #F
        Data = Space$(LOF(F))
        Get #F, , Data
        Close #F
        MapCRC32(MapNum).MapDataCRC = CRC32(Data)
    Else
        MapCRC32(MapNum).MapDataCRC = 0
    End If
    ' clear
    Erase Data
    ' tile data
    filename = App.Path & "\data\maps\map" & MapNum & ".dat"
    If FileExist(filename) Then
        F = FreeFile
        Open filename For Binary As #F
        Data = Space$(LOF(F))
        Get #F, , Data
        Close #F
        MapCRC32(MapNum).MapTileCRC = CRC32(Data)
    Else
        MapCRC32(MapNum).MapTileCRC = 0
    End If
End Sub
