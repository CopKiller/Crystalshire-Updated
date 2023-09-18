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
        PutVar filename, "General", "BossNpc", Val(.BossNpc)
        For i = 1 To MAX_MAP_NPCS
            PutVar filename, "General", "Npc" & i, Val(.Npc(i))
        Next
    End With

    ' Events
    PutVar filename, "Events", "EventCount", Val(Map(mapnum).TileData.EventCount)

    If Map(mapnum).TileData.EventCount > 0 Then
        For i = 1 To Map(mapnum).TileData.EventCount
            With Map(mapnum).TileData.Events(i)
                PutVar filename, "Event" & i, "Name", .Name
                PutVar filename, "Event" & i, "x", Val(.x)
                PutVar filename, "Event" & i, "y", Val(.y)
                PutVar filename, "Event" & i, "PageCount", Val(.PageCount)
            End With
            If Map(mapnum).TileData.Events(i).PageCount > 0 Then
                For x = 1 To Map(mapnum).TileData.Events(i).PageCount
                    With Map(mapnum).TileData.Events(i).EventPage(x)
                        PutVar filename, "Event" & i & "Page" & x, "chkPlayerVar", Val(.chkPlayerVar)
                        PutVar filename, "Event" & i & "Page" & x, "chkSelfSwitch", Val(.chkSelfSwitch)
                        PutVar filename, "Event" & i & "Page" & x, "chkHasItem", Val(.chkHasItem)
                        PutVar filename, "Event" & i & "Page" & x, "PlayerVarNum", Val(.PlayerVarNum)
                        PutVar filename, "Event" & i & "Page" & x, "SelfSwitchNum", Val(.SelfSwitchNum)
                        PutVar filename, "Event" & i & "Page" & x, "HasItemNum", Val(.HasItemNum)
                        PutVar filename, "Event" & i & "Page" & x, "PlayerVariable", Val(.PlayerVariable)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicType", Val(.GraphicType)
                        PutVar filename, "Event" & i & "Page" & x, "Graphic", Val(.Graphic)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicX", Val(.GraphicX)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicY", Val(.GraphicY)
                        PutVar filename, "Event" & i & "Page" & x, "MoveType", Val(.MoveType)
                        PutVar filename, "Event" & i & "Page" & x, "MoveSpeed", Val(.MoveSpeed)
                        PutVar filename, "Event" & i & "Page" & x, "MoveFreq", Val(.MoveFreq)
                        PutVar filename, "Event" & i & "Page" & x, "WalkAnim", Val(.WalkAnim)
                        PutVar filename, "Event" & i & "Page" & x, "StepAnim", Val(.StepAnim)
                        PutVar filename, "Event" & i & "Page" & x, "DirFix", Val(.DirFix)
                        PutVar filename, "Event" & i & "Page" & x, "WalkThrough", Val(.WalkThrough)
                        PutVar filename, "Event" & i & "Page" & x, "Priority", Val(.Priority)
                        PutVar filename, "Event" & i & "Page" & x, "Trigger", Val(.Trigger)
                        PutVar filename, "Event" & i & "Page" & x, "CommandCount", Val(.CommandCount)
                    End With
                    If Map(mapnum).TileData.Events(i).EventPage(x).CommandCount > 0 Then
                        For y = 1 To Map(mapnum).TileData.Events(i).EventPage(x).CommandCount
                            With Map(mapnum).TileData.Events(i).EventPage(x).Commands(y)
                                PutVar filename, "Event" & i & "Page" & x & "Command" & y, "Type", Val(.Type)
                                PutVar filename, "Event" & i & "Page" & x & "Command" & y, "Text", .Text
                                PutVar filename, "Event" & i & "Page" & x & "Command" & y, "Colour", Val(.colour)
                                PutVar filename, "Event" & i & "Page" & x & "Command" & y, "Channel", Val(.Channel)
                                PutVar filename, "Event" & i & "Page" & x & "Command" & y, "TargetType", Val(.targetType)
                                PutVar filename, "Event" & i & "Page" & x & "Command" & y, "Target", Val(.target)
                            End With
                        Next
                    End If
                Next
            End If
        Next
    End If

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
        .BossNpc = Val(GetVar(filename, "General", "BossNpc"))
        For i = 1 To MAX_MAP_NPCS
            .Npc(i) = Val(GetVar(filename, "General", "Npc" & i))
        Next
    End With

    ' Events
    Map(mapnum).TileData.EventCount = Val(GetVar(filename, "Events", "EventCount"))

    If Map(mapnum).TileData.EventCount > 0 Then
        ReDim Preserve Map(mapnum).TileData.Events(1 To Map(mapnum).TileData.EventCount)
        For i = 1 To Map(mapnum).TileData.EventCount
            With Map(mapnum).TileData.Events(i)
                .Name = GetVar(filename, "Event" & i, "Name")
                .x = Val(GetVar(filename, "Event" & i, "x"))
                .y = Val(GetVar(filename, "Event" & i, "y"))
                .PageCount = Val(GetVar(filename, "Event" & i, "PageCount"))
            End With
            If Map(mapnum).TileData.Events(i).PageCount > 0 Then
                ReDim Preserve Map(mapnum).TileData.Events(i).EventPage(1 To Map(mapnum).TileData.Events(i).PageCount)
                For x = 1 To Map(mapnum).TileData.Events(i).PageCount
                    With Map(mapnum).TileData.Events(i).EventPage(x)
                        .chkPlayerVar = Val(GetVar(filename, "Event" & i & "Page" & x, "chkPlayerVar"))
                        .chkSelfSwitch = Val(GetVar(filename, "Event" & i & "Page" & x, "chkSelfSwitch"))
                        .chkHasItem = Val(GetVar(filename, "Event" & i & "Page" & x, "chkHasItem"))
                        .PlayerVarNum = Val(GetVar(filename, "Event" & i & "Page" & x, "PlayerVarNum"))
                        .SelfSwitchNum = Val(GetVar(filename, "Event" & i & "Page" & x, "SelfSwitchNum"))
                        .HasItemNum = Val(GetVar(filename, "Event" & i & "Page" & x, "HasItemNum"))
                        .PlayerVariable = Val(GetVar(filename, "Event" & i & "Page" & x, "PlayerVariable"))
                        .GraphicType = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicType"))
                        .Graphic = Val(GetVar(filename, "Event" & i & "Page" & x, "Graphic"))
                        .GraphicX = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicX"))
                        .GraphicY = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicY"))
                        .MoveType = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveType"))
                        .MoveSpeed = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveSpeed"))
                        .MoveFreq = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveFreq"))
                        .WalkAnim = Val(GetVar(filename, "Event" & i & "Page" & x, "WalkAnim"))
                        .StepAnim = Val(GetVar(filename, "Event" & i & "Page" & x, "StepAnim"))
                        .DirFix = Val(GetVar(filename, "Event" & i & "Page" & x, "DirFix"))
                        .WalkThrough = Val(GetVar(filename, "Event" & i & "Page" & x, "WalkThrough"))
                        .Priority = Val(GetVar(filename, "Event" & i & "Page" & x, "Priority"))
                        .Trigger = Val(GetVar(filename, "Event" & i & "Page" & x, "Trigger"))
                        .CommandCount = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandCount"))
                    End With
                    If Map(mapnum).TileData.Events(i).EventPage(x).CommandCount > 0 Then
                        ReDim Preserve Map(mapnum).TileData.Events(i).EventPage(x).Commands(1 To Map(mapnum).TileData.Events(i).EventPage(x).CommandCount)
                        For y = 1 To Map(mapnum).TileData.Events(i).EventPage(x).CommandCount
                            With Map(mapnum).TileData.Events(i).EventPage(x).Commands(y)
                                .Type = Val(GetVar(filename, "Event" & i & "Page" & x & "Command" & y, "Type"))
                                .Text = GetVar(filename, "Event" & i & "Page" & x & "Command" & y, "Text")
                                .colour = Val(GetVar(filename, "Event" & i & "Page" & x & "Command" & y, "Colour"))
                                .Channel = Val(GetVar(filename, "Event" & i & "Page" & x & "Command" & y, "Channel"))
                                .targetType = Val(GetVar(filename, "Event" & i & "Page" & x & "Command" & y, "TargetType"))
                                .target = Val(GetVar(filename, "Event" & i & "Page" & x & "Command" & y, "Target"))
                            End With
                        Next
                    End If
                Next
            End If
        Next
    End If

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
    Call ZeroMemory(ByVal VarPtr(Map(mapnum)), LenB(Map(mapnum)))
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

Public Sub ClearMapItem(ByVal index As Long, ByVal mapnum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(mapnum, index)), LenB(MapItem(mapnum, index)))
    MapItem(mapnum, index).playerName = vbNullString
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

Public Sub ClearMapNpc(ByVal index As Long, ByVal mapnum As Long)
    Call ZeroMemory(ByVal VarPtr(MapNpc(mapnum).Npc(index)), LenB(MapNpc(mapnum).Npc(index)))
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
