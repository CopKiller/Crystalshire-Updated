Attribute VB_Name = "modDatabase"
Option Explicit
' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

Private crcTable(0 To 255) As Long

Public Sub InitCRC32()
Dim i As Long, n As Long, CRC As Long

    For i = 0 To 255
        CRC = i
        For n = 0 To 7
            If CRC And 1 Then
                CRC = (((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor &HEDB88320
            Else
                CRC = ((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF
            End If
        Next
        crcTable(i) = CRC
    Next
End Sub

Public Function CRC32(ByRef Data() As Byte) As Long
Dim lCurPos As Long
Dim lLen As Long

    lLen = AryCount(Data) - 1
    CRC32 = &HFFFFFFFF
    
    For lCurPos = 0 To lLen
        CRC32 = (((CRC32 And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (crcTable((CRC32 And 255) Xor Data(lCurPos)))
    Next
    
    CRC32 = CRC32 Xor &HFFFFFFFF
End Function

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)

    If LCase$(Dir$(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

Public Function FileExist(ByVal FileName As String) As Boolean

    If LenB(Dir$(FileName)) > 0 Then
        FileExist = True
    End If

End Function

' gets a string from a text file
Public Function GetVar(File As String, header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(header, Var, Value, File)
End Sub

Public Sub SaveOptions()
    Dim FileName As String, i As Long
    
    FileName = App.Path & "\Data Files\config_v2.ini"
    
    Call PutVar(FileName, "Options", "Username", Options.Username)
    Call PutVar(FileName, "Options", "Music", Str$(Options.Music))
    Call PutVar(FileName, "Options", "Sound", Str$(Options.sound))
    Call PutVar(FileName, "Options", "NoAuto", Str$(Options.NoAuto))
    Call PutVar(FileName, "Options", "Render", Str$(Options.Render))
    Call PutVar(FileName, "Options", "SaveUser", Str$(Options.SaveUser))
    Call PutVar(FileName, "Options", "Resolution", Str$(Options.resolution))
    Call PutVar(FileName, "Options", "Fullscreen", Str$(Options.Fullscreen))
    
    Call PutVar(FileName, "Options", "FPSLock", Trim$(Options.FPSLock))
    For i = 0 To ChatChannel.Channel_Count - 1
        Call PutVar(FileName, "Options", "Channel" & i, Str$(Options.channelState(i)))
    Next
End Sub

Public Sub LoadOptions()
    Dim FileName As String, i As Long
    
    On Error GoTo ErrorHandler
    
    FileName = App.Path & "\Data Files\config_v2.ini"

    If Not FileExist(FileName) Then
        GoTo ErrorHandler
    Else
        Options.Username = GetVar(FileName, "Options", "Username")
        Options.Music = GetVar(FileName, "Options", "Music")
        Options.sound = Val(GetVar(FileName, "Options", "Sound"))
        Options.NoAuto = Val(GetVar(FileName, "Options", "NoAuto"))
        Options.Render = Val(GetVar(FileName, "Options", "Render"))
        Options.SaveUser = Val(GetVar(FileName, "Options", "SaveUser"))
        Options.resolution = Val(GetVar(FileName, "Options", "Resolution"))
        Options.Fullscreen = Val(GetVar(FileName, "Options", "Fullscreen"))
        
        If Not GetVar(FileName, "Options", "FPSLock") = "" Then
            Options.FPSLock = CBool(GetVar(FileName, "Options", "FPSLock"))
        Else
            Options.FPSLock = False
            Call PutVar(FileName, "Options", "FPSLock", Trim$(Options.FPSLock))
        End If
        
        For i = 0 To ChatChannel.Channel_Count - 1
            Options.channelState(i) = Val(GetVar(FileName, "Options", "Channel" & i))
        Next
    End If
    
    Exit Sub
ErrorHandler:
    Options.Music = 1
    Options.sound = 1
    Options.NoAuto = 0
    Options.Username = vbNullString
    Options.Fullscreen = 0
    Options.Render = 0
    Options.SaveUser = 0
    For i = 0 To ChatChannel.Channel_Count - 1
        Options.channelState(i) = 1
    Next
    SaveOptions
    Exit Sub
End Sub

Public Sub SaveMap(ByVal mapNum As Long)
    Dim FileName As String, f As Long, X As Long, Y As Long, i As Long
    
    ' save map data
    FileName = App.Path & MAP_PATH & mapNum & "_.dat"
    
    ' if it exists then kill the ini
    If FileExist(FileName) Then Kill FileName
    
    ' General
    With Map.MapData
        PutVar FileName, "General", "Name", .Name
        PutVar FileName, "General", "Music", .Music
        PutVar FileName, "General", "Moral", Val(.Moral)
        PutVar FileName, "General", "Up", Val(.Up)
        PutVar FileName, "General", "Down", Val(.Down)
        PutVar FileName, "General", "Left", Val(.Left)
        PutVar FileName, "General", "Right", Val(.Right)
        PutVar FileName, "General", "BootMap", Val(.BootMap)
        PutVar FileName, "General", "BootX", Val(.BootX)
        PutVar FileName, "General", "BootY", Val(.BootY)
        PutVar FileName, "General", "MaxX", Val(.MaxX)
        PutVar FileName, "General", "MaxY", Val(.MaxY)
        
        PutVar FileName, "General", "Weather", Val(.Weather)
        PutVar FileName, "General", "WeatherIntensity", Val(.WeatherIntensity)
        
        PutVar FileName, "General", "Fog", Val(.Fog)
        PutVar FileName, "General", "FogSpeed", Val(.FogSpeed)
        PutVar FileName, "General", "FogOpacity", Val(.FogOpacity)
        
        PutVar FileName, "General", "Red", Val(.Red)
        PutVar FileName, "General", "Green", Val(.Green)
        PutVar FileName, "General", "Blue", Val(.Blue)
        PutVar FileName, "General", "Alpha", Val(.alpha)
        
        PutVar FileName, "General", "BossNpc", Val(.BossNpc)
        For i = 1 To MAX_MAP_NPCS
            PutVar FileName, "General", "Npc" & i, Val(.Npc(i))
        Next
    End With
    
    ' dump tile data
    FileName = App.Path & MAP_PATH & mapNum & ".dat"
    
    ' if it exists then kill the ini
    If FileExist(FileName) Then Kill FileName
    
    f = FreeFile
    With Map
        Open FileName For Binary As #f
            For X = 0 To .MapData.MaxX
                For Y = 0 To .MapData.MaxY
                    Put #f, , .TileData.Tile(X, Y).Type
                    Put #f, , .TileData.Tile(X, Y).Data1
                    Put #f, , .TileData.Tile(X, Y).Data2
                    Put #f, , .TileData.Tile(X, Y).Data3
                    Put #f, , .TileData.Tile(X, Y).Data4
                    Put #f, , .TileData.Tile(X, Y).Data5
                    Put #f, , .TileData.Tile(X, Y).Autotile
                    Put #f, , .TileData.Tile(X, Y).DirBlock
                    For i = 1 To MapLayer.Layer_Count - 1
                        Put #f, , .TileData.Tile(X, Y).Layer(i).tileSet
                        Put #f, , .TileData.Tile(X, Y).Layer(i).X
                        Put #f, , .TileData.Tile(X, Y).Layer(i).Y
                    Next
                Next
            Next
        Close #f
    End With
    
    Close #f
End Sub

Public Sub GetMapCRC32(mapNum As Long)
    Dim Data() As Byte, FileName As String, f As Long
    ' map data
    FileName = App.Path & MAP_PATH & mapNum & "_.dat"
    If FileExist(FileName) Then
        f = FreeFile
        Open FileName For Binary As #f
            Data = Space$(LOF(f))
            Get #f, , Data
        Close #f
        MapCRC32(mapNum).MapDataCRC = CRC32(Data)
    Else
        MapCRC32(mapNum).MapDataCRC = 0
    End If
    ' clear
    Erase Data
    ' tile data
    FileName = App.Path & MAP_PATH & mapNum & ".dat"
    If FileExist(FileName) Then
        f = FreeFile
        Open FileName For Binary As #f
            Data = Space$(LOF(f))
            Get #f, , Data
        Close #f
        MapCRC32(mapNum).MapTileCRC = CRC32(Data)
    Else
        MapCRC32(mapNum).MapTileCRC = 0
    End If
End Sub

Public Sub LoadMap(ByVal mapNum As Long)
    Dim FileName As String, i As Long, f As Long, X As Long, Y As Long
    
    ' load map data
    FileName = App.Path & MAP_PATH & mapNum & "_.dat"
    
    ' General
    With Map.MapData
        .Name = GetVar(FileName, "General", "Name")
        .Music = GetVar(FileName, "General", "Music")
        .Moral = Val(GetVar(FileName, "General", "Moral"))
        .Up = Val(GetVar(FileName, "General", "Up"))
        .Down = Val(GetVar(FileName, "General", "Down"))
        .Left = Val(GetVar(FileName, "General", "Left"))
        .Right = Val(GetVar(FileName, "General", "Right"))
        .BootMap = Val(GetVar(FileName, "General", "BootMap"))
        .BootX = Val(GetVar(FileName, "General", "BootX"))
        .BootY = Val(GetVar(FileName, "General", "BootY"))
        .MaxX = Val(GetVar(FileName, "General", "MaxX"))
        .MaxY = Val(GetVar(FileName, "General", "MaxY"))
        
        .Weather = Val(GetVar(FileName, "General", "Weather"))
        .WeatherIntensity = Val(GetVar(FileName, "General", "WeatherIntensity"))
        
        .Fog = Val(GetVar(FileName, "General", "Fog"))
        .FogSpeed = Val(GetVar(FileName, "General", "FogSpeed"))
        .FogOpacity = Val(GetVar(FileName, "General", "FogOpacity"))
        
        .Red = Val(GetVar(FileName, "General", "Red"))
        .Green = Val(GetVar(FileName, "General", "Green"))
        .Blue = Val(GetVar(FileName, "General", "Blue"))
        .alpha = Val(GetVar(FileName, "General", "Alpha"))
        .BossNpc = Val(GetVar(FileName, "General", "BossNpc"))
        For i = 1 To MAX_MAP_NPCS
            .Npc(i) = Val(GetVar(FileName, "General", "Npc" & i))
        Next
    End With
    
    ' dump tile data
    FileName = App.Path & MAP_PATH & mapNum & ".dat"
    f = FreeFile
    
    ReDim Map.TileData.Tile(0 To Map.MapData.MaxX, 0 To Map.MapData.MaxY) As TileRec
    
    With Map
        Open FileName For Binary As #f
            For X = 0 To .MapData.MaxX
                For Y = 0 To .MapData.MaxY
                    Get #f, , .TileData.Tile(X, Y).Type
                    Get #f, , .TileData.Tile(X, Y).Data1
                    Get #f, , .TileData.Tile(X, Y).Data2
                    Get #f, , .TileData.Tile(X, Y).Data3
                    Get #f, , .TileData.Tile(X, Y).Data4
                    Get #f, , .TileData.Tile(X, Y).Data5
                    Get #f, , .TileData.Tile(X, Y).Autotile
                    Get #f, , .TileData.Tile(X, Y).DirBlock
                    For i = 1 To MapLayer.Layer_Count - 1
                        Get #f, , .TileData.Tile(X, Y).Layer(i).tileSet
                        Get #f, , .TileData.Tile(X, Y).Layer(i).X
                        Get #f, , .TileData.Tile(X, Y).Layer(i).Y
                    Next
                Next
            Next
        Close #f
    End With
    
    ClearTempTile
End Sub

Public Sub ClearPlayer(ByVal Index As Long)
    Player(Index) = EmptyPlayer
    Player(Index).Name = vbNullString
End Sub

Public Sub ClearProjectile(ByVal ProjectileSlot As Long)
    MapProjectile(ProjectileSlot) = EmptyMapProjectile
End Sub

Public Sub ClearItem(ByVal Index As Long)
    Item(Index) = EmptyItem
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).sound = "None."
End Sub

Public Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

Public Sub ClearMapItem(ByVal Index As Long)
    MapItem(Index) = EmptyMapItem
End Sub

Public Sub ClearMap()
    Map = EmptyMap
    Map.MapData.Name = vbNullString
    Map.MapData.MaxX = MAX_MAPX
    Map.MapData.MaxY = MAX_MAPY
    ReDim Map.TileData.Tile(0 To Map.MapData.MaxX, 0 To Map.MapData.MaxY)
    initAutotiles
End Sub

Public Sub ClearMapItems()
    Dim i As Long

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next

End Sub

Public Sub ClearMapNpc(ByVal Index As Long)
    MapNpc(Index) = EmptyMapNpc
End Sub

Public Sub ClearMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next

End Sub

' **********************
' ** Player functions **
' **********************
Function GetPlayerName(ByVal Index As Long) As String

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Name = Name
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal sprite As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).sprite = sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Level = Level
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(Index).EXP
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).EXP = EXP
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).PK = PK
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerMaxVital = Player(Index).MaxVital(Vital)
End Function

Function GetPlayerStat(ByVal Index As Long, Stat As Stats) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(Index).Stat(Stat)
End Function

Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal Value As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    If Value <= 0 Then Value = 1
    If Value > MAX_BYTE Then Value = MAX_BYTE
    Player(Index).Stat(Stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Or Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal mapNum As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Map = mapNum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Dir = Dir
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    GetPlayerInvItemNum = PlayerInv(invSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemNum As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(invSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invSlot).Value = ItemValue
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot <= 0 Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Equipment(EquipmentSlot) = invNum
End Sub
