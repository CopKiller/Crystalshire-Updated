Attribute VB_Name = "Map_UDT"
Public Map(1 To MAX_MAPS) As MapRec
Public MapCRC32(1 To MAX_MAPS) As MapCRCStruct
Public MapCache(1 To MAX_MAPS) As Cache
Public TempTile(1 To MAX_MAPS) As TempTileRec

Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS) As MapNpcDataRec

Public Type MapCRCStruct
    MapDataCRC As Long
    MapTileCRC As Long
End Type

Private Type MapDataRec
    Name As String
    Music As String
    Moral As Byte
    
    Up As Long
    Down As Long
    left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    
    BossNpc As Long
    
    Npc(1 To MAX_MAP_NPCS) As Long
End Type

Private Type TileDataRec
    x As Long
    y As Long
    Tileset As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte

    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    Data5 As Long
    DirBlock As Byte
End Type

Private Type MapTileRec
    EventCount As Long
    Tile() As TileRec
    Events() As EventRec
End Type

Private Type MapRec
    MapData As MapDataRec
    TileData As MapTileRec
End Type

Private Type MapItemRec
    Num As Long
    Value As Long
    x As Byte
    y As Byte
    ' ownership + despawn
    playerName As String
    playerTimer As Long
    canDespawn As Boolean
    despawnTimer As Long
    Bound As Boolean
End Type

Private Type MapNpcRec
    Num As Long
    target As Long
    targetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    x As Byte
    y As Byte
    dir As Byte
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    StunDuration As Long
    StunTimer As Long
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' chat
    c_lastDir As Byte
    c_inChatWith As Long
    ' spell casting
    spellBuffer As SpellBufferRec
    SpellCD(1 To MAX_NPC_SPELLS) As Long
End Type

Private Type MapNpcDataRec
    Npc(1 To MAX_MAP_NPCS) As MapNpcRec
End Type

Private Type TempMapDataRec
    Npc() As MapNpcRec
End Type

Private Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Long
    x As Long
    y As Long
    cur_health As Long
End Type

Private Type TempTileRec
    DoorOpen() As Byte
    DoorTimer As Long
End Type
