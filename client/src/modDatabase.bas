Attribute VB_Name = "modDatabase"
Option Explicit
' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

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

Public Function FileExist(ByVal filename As String) As Boolean

    If LenB(Dir$(filename)) > 0 Then
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
Public Sub PutVar(File As String, header As String, Var As String, value As String)
    Call WritePrivateProfileString$(header, Var, value, File)
End Sub

Public Sub SaveOptions()
    Dim filename As String, i As Long
    
    filename = App.path & "\Data Files\config_v2.ini"
    
    Call PutVar(filename, "Options", "Username", Options.Username)
    Call PutVar(filename, "Options", "Music", Str$(Options.Music))
    Call PutVar(filename, "Options", "Sound", Str$(Options.sound))
    Call PutVar(filename, "Options", "NoAuto", Str$(Options.NoAuto))
    Call PutVar(filename, "Options", "Render", Str$(Options.Render))
    Call PutVar(filename, "Options", "SaveUser", Str$(Options.SaveUser))
    Call PutVar(filename, "Options", "Resolution", Str$(Options.Resolution))
    Call PutVar(filename, "Options", "Fullscreen", Str$(Options.Fullscreen))
    For i = 0 To ChatChannel.Channel_Count - 1
        Call PutVar(filename, "Options", "Channel" & i, Str$(Options.channelState(i)))
    Next
End Sub

Public Sub LoadOptions()
    Dim filename As String, i As Long
    
    On Error GoTo errorhandler
    
    filename = App.path & "\Data Files\config_v2.ini"

    If Not FileExist(filename) Then
        GoTo errorhandler
    Else
        Options.Username = GetVar(filename, "Options", "Username")
        Options.Music = GetVar(filename, "Options", "Music")
        Options.sound = Val(GetVar(filename, "Options", "Sound"))
        Options.NoAuto = Val(GetVar(filename, "Options", "NoAuto"))
        Options.Render = Val(GetVar(filename, "Options", "Render"))
        Options.SaveUser = Val(GetVar(filename, "Options", "SaveUser"))
        Options.Resolution = Val(GetVar(filename, "Options", "Resolution"))
        Options.Fullscreen = Val(GetVar(filename, "Options", "Fullscreen"))
        For i = 0 To ChatChannel.Channel_Count - 1
            Options.channelState(i) = Val(GetVar(filename, "Options", "Channel" & i))
        Next
    End If
    
    Exit Sub
errorhandler:
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
    Dim filename As String, f As Long, x As Long, y As Long, i As Long
    
    ' save map data
    filename = App.path & MAP_PATH & mapNum & "_.dat"
    
    ' if it exists then kill the ini
    If FileExist(filename) Then Kill filename
    
    ' General
    With Map.MapData
        PutVar filename, "General", "Name", .name
        PutVar filename, "General", "Music", .Music
        PutVar filename, "General", "Moral", Val(.Moral)
        PutVar filename, "General", "Up", Val(.Up)
        PutVar filename, "General", "Down", Val(.Down)
        PutVar filename, "General", "Left", Val(.Left)
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
        PutVar filename, "General", "Alpha", Val(.alpha)
        
        PutVar filename, "General", "BossNpc", Val(.BossNpc)
        For i = 1 To MAX_MAP_NPCS
            PutVar filename, "General", "Npc" & i, Val(.Npc(i))
        Next
    End With
    
    ' dump tile data
    filename = App.path & MAP_PATH & mapNum & ".dat"
    
    ' if it exists then kill the ini
    If FileExist(filename) Then Kill filename
    
    f = FreeFile
    With Map
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
                        Put #f, , .TileData.Tile(x, y).Layer(i).tileSet
                        Put #f, , .TileData.Tile(x, y).Layer(i).x
                        Put #f, , .TileData.Tile(x, y).Layer(i).y
                    Next
                Next
            Next
        Close #f
    End With
    
    Close #f
End Sub

Sub GetMapCRC32(mapNum As Long)
Dim Data() As Byte, filename As String, f As Long
    ' map data
    filename = App.path & MAP_PATH & mapNum & "_.dat"
    If FileExist(filename) Then
        f = FreeFile
        Open filename For Binary As #f
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
    filename = App.path & MAP_PATH & mapNum & ".dat"
    If FileExist(filename) Then
        f = FreeFile
        Open filename For Binary As #f
            Data = Space$(LOF(f))
            Get #f, , Data
        Close #f
        MapCRC32(mapNum).MapTileCRC = CRC32(Data)
    Else
        MapCRC32(mapNum).MapTileCRC = 0
    End If
End Sub

Public Sub LoadMap(ByVal mapNum As Long)
    Dim filename As String, i As Long, f As Long, x As Long, y As Long
    
    ' load map data
    filename = App.path & MAP_PATH & mapNum & "_.dat"
    
    ' General
    With Map.MapData
        .name = GetVar(filename, "General", "Name")
        .Music = GetVar(filename, "General", "Music")
        .Moral = Val(GetVar(filename, "General", "Moral"))
        .Up = Val(GetVar(filename, "General", "Up"))
        .Down = Val(GetVar(filename, "General", "Down"))
        .Left = Val(GetVar(filename, "General", "Left"))
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
        .alpha = Val(GetVar(filename, "General", "Alpha"))
        .BossNpc = Val(GetVar(filename, "General", "BossNpc"))
        For i = 1 To MAX_MAP_NPCS
            .Npc(i) = Val(GetVar(filename, "General", "Npc" & i))
        Next
    End With
    
    ' dump tile data
    filename = App.path & MAP_PATH & mapNum & ".dat"
    f = FreeFile
    
    ReDim Map.TileData.Tile(0 To Map.MapData.MaxX, 0 To Map.MapData.MaxY) As TileRec
    
    With Map
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
                        Get #f, , .TileData.Tile(x, y).Layer(i).tileSet
                        Get #f, , .TileData.Tile(x, y).Layer(i).x
                        Get #f, , .TileData.Tile(x, y).Layer(i).y
                    Next
                Next
            Next
        Close #f
    End With
    
    ClearTempTile
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Player(Index) = EmptyPlayer
    Player(Index).name = vbNullString
End Sub

Sub ClearItem(ByVal Index As Long)
    Item(Index) = EmptyItem
    Item(Index).name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).sound = "None."
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

Sub ClearAnimInstance(ByVal Index As Long)
    AnimInstance(Index) = EmptyAnimInstance
End Sub

Sub ClearAnimation(ByVal Index As Long)
    Animation(Index) = EmptyAnimation
    Animation(Index).name = vbNullString
    Animation(Index).sound = "None."
End Sub

Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next

End Sub

Sub ClearNPC(ByVal Index As Long)
    Npc(Index) = EmptyNpc
    Npc(Index).name = vbNullString
    Npc(Index).sound = "None."
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next

End Sub

Sub ClearSpell(ByVal Index As Long)
    Spell(Index) = EmptySpell
    Spell(Index).name = vbNullString
    Spell(Index).Desc = vbNullString
    Spell(Index).sound = "None."
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

Sub ClearShop(ByVal Index As Long)
    Shop(Index) = EmptyShop
    Shop(Index).name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

Sub ClearResource(ByVal Index As Long)
    Resource(Index) = EmptyResource
    Resource(Index).name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).sound = "None."
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next

End Sub

Sub ClearMapItem(ByVal Index As Long)
    MapItem(Index) = EmptyMapItem
End Sub

Sub ClearMap()
    Map = EmptyMap
    Map.MapData.name = vbNullString
    Map.MapData.MaxX = MAX_MAPX
    Map.MapData.MaxY = MAX_MAPY
    ReDim Map.TileData.Tile(0 To Map.MapData.MaxX, 0 To Map.MapData.MaxY)
    initAutotiles
End Sub

Sub ClearMapItems()
    Dim i As Long

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next

End Sub

Sub ClearMapNpc(ByVal Index As Long)
    MapNpc(Index) = EmptyMapNpc
End Sub

Sub ClearMapNpcs()
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
    GetPlayerName = Trim$(Player(Index).name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal name As String)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).name = name
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

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal value As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Vital(Vital) = value

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

Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal value As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    If value <= 0 Then value = 1
    If value > MAX_BYTE Then value = MAX_BYTE
    Player(Index).Stat(Stat) = value
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
    GetPlayerX = Player(Index).x
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).x = x
End Sub

Function GetPlayerY(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).y = y
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
    GetPlayerInvItemNum = PlayerInv(invSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long, ByVal itemNum As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invSlot).num = itemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(invSlot).value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invSlot).value = ItemValue
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Equipment(EquipmentSlot) = invNum
End Sub

Sub ClearConv(ByVal Index As Long)
    Conv(Index) = EmptyConv
    Conv(Index).name = vbNullString
    ReDim Conv(Index).Conv(1)
End Sub

Sub ClearConvs()
    Dim i As Long

    For i = 1 To MAX_CONVS
        Call ClearConv(i)
    Next

End Sub
