Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map As MapRec
Public MapCRC32(1 To MAX_MAPS) As MapCRCStruct
Public Particula(1 To MAX_WEATHER_PARTICLES) As ParticulaRec
Public Bank As BankRec
Public TempTile() As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public Blood(1 To MAX_BYTE) As BloodRec
Public Party As PartyRec
Public Autotile() As AutotileRec
Public MapSounds() As MapSoundRec
Public MapSoundCount As Long
Public Options As OptionsRec

' Cliente strutures
Public MapProjectile() As ProjectileRenderRec
Public EmptyMapProjectile As ProjectileRenderRec

Public EmptyMap As MapRec
Public EmptyPlayer As PlayerRec
Public EmptyItem As ItemRec
Public EmptyMapItem As MapItemRec
Public EmptyMapNpc As MapNpcRec
Public EmptyActionMsg As ActionMsgRec

'Client
Public WeatherParticle(1 To MAX_WEATHER_PARTICLES) As WeatherParticleRec

' Type recs
Public Type MapCRCStruct
    MapDataCRC As Long
    MapTileCRC As Long
End Type

Private Type OptionsRec
    Music As Byte
    sound As Byte
    NoAuto As Byte
    Render As Byte
    Username As String
    SaveUser As Long
    FPSLock As Boolean
    channelState(0 To Channel_Count - 1) As Byte
    PlayIntro As Byte
    resolution As Byte
    Fullscreen As Byte
End Type

Public Type XYRec
    X As Double
    Y As Double
End Type

Public Type ProjectileRenderRec
    Owner As Long
    OwnerType As Byte
    Graphic As Long
    
    speed As Long
    RotateSpeed As Byte
    Rotate As Single
    Duration As Long
    ProjectileOffset(1 To 4) As XYRec
    Direction As Byte
    X As Long
    Y As Long
    xOffset As Long
    yOffset As Long
    tx As Long
    ty As Long
    'Cliente Apenas
    IsAoE As Boolean
    IsDirectional As Boolean
    curAnim As Long
    EndTime As Long
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
    bound As Byte
End Type

Public Type PlayerSpellRec
    Spell As Long
    Uses As Long
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Private Type PlayerMission
    ID As Long
    Count As Long
End Type

Private Type PlayerRec
    ' General
    Name As String
    Class As Long
    sprite As Long
    Level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    MaxVital(1 To Vitals.Vital_Count - 1) As Long
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    'Projectiles
    Projectile(1 To MAX_PROJECTILE_PLAYER) As Long
    ' Position
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    ' Variables
    Variable(1 To MAX_BYTE) As Long
    ' Missions
    Mission(1 To MAX_PLAYER_MISSIONS) As PlayerMission
    CompletedMission(1 To MAX_MISSIONS) As Long
    ' Client use only
    xOffset As Integer
    yOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    Step As Byte
    Anim As Long
    AnimTimer As Long
End Type

Private Type EventCommandRec
    Type As Byte
    text As String
    Colour As Long
    Channel As Byte
    TargetType As Byte
    target As Long
    X As Long
    Y As Long
End Type

Public Type EventPageRec
    chkPlayerVar As Byte
    chkSelfSwitch As Byte
    chkHasItem As Byte
    
    PlayerVarNum As Long
    SelfSwitchNum As Long
    HasItemNum As Long
    
    PlayerVariable As Long
    
    GraphicType As Byte
    Graphic As Long
    GraphicX As Long
    GraphicY As Long
    
    MoveType As Byte
    MoveSpeed As Byte
    MoveFreq As Byte
    
    WalkAnim As Byte
    StepAnim As Byte
    DirFix As Byte
    WalkThrough As Byte
    
    Priority As Byte
    Trigger As Byte
    
    CommandCount As Long
    Commands() As EventCommandRec
End Type

Public Type EventRec
    Name As String
    X As Long
    Y As Long
    pageCount As Long
    EventPage() As EventPageRec
End Type

Private Type MapDataRec
    Name As String
    Music As String
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    
    Weather As Long
    WeatherIntensity As Long
    
    Fog As Long
    FogSpeed As Long
    FogOpacity As Long
    
    Red As Long
    Green As Long
    Blue As Long
    alpha As Long
    
    BossNpc As Long
    
    Npc(1 To MAX_MAP_NPCS) As Long
End Type

Private Type TileDataRec
    X As Long
    Y As Long
    tileSet As Long
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

Private Type ClassRec
    Name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
End Type

Public Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH
    pic As Long

    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Mastery As Byte
    Price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Rarity As Byte
    speed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Byte
    Animation As Long
    Paperdoll As Long
    ' consume
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    instaCast As Byte
    ' food
    HPorSP As Long
    FoodPerTick As Long
    FoodTickCount As Long
    FoodInterval As Long
    ' requirements
    proficiency As Long
End Type

Private Type MapItemRec
    playerName As String
    Num As Long
    Value As Long
    Frame As Byte
    X As Byte
    Y As Byte
    bound As Boolean
End Type

Private Type MapNpcRec
    Num As Long
    target As Long
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    ' Client use only
    xOffset As Long
    yOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
    Anim As Long
    AnimTimer As Long
End Type

Private Type TempTileRec
    ' doors... obviously
    DoorOpen As Byte
    DoorFrame As Byte
    DoorTimer As Long
    DoorAnimate As Byte ' 0 = nothing| 1 = opening | 2 = closing
    ' fading appear tiles
    isFading(1 To MapLayer.Layer_Count - 1) As Boolean
    fadeAlpha(1 To MapLayer.Layer_Count - 1) As Long
    FadeTimer(1 To MapLayer.Layer_Count - 1) As Long
    FadeDir(1 To MapLayer.Layer_Count - 1) As Byte
End Type

Public Type MapResourceRec
    X As Long
    Y As Long
    ResourceState As Byte
End Type

Private Type ActionMsgRec
    message As String
    Created As Long

    Type As Long
    Color As Long
    Scroll As Long
    X As Long
    Y As Long
    timer As Long
    alpha As Long
End Type

Private Type BloodRec
    sprite As Long
    timer As Long
    X As Long
    Y As Long
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Public Type PointRec
    X As Long
    Y As Long
End Type

Public Type QuarterTileRec
    QuarterTile(1 To 4) As PointRec
    RenderState As Byte
    srcX(1 To 4) As Long
    srcY(1 To 4) As Long
End Type

Public Type AutotileRec
    Layer(1 To MapLayer.Layer_Count - 1) As QuarterTileRec
End Type

Public Type ChatBubbleRec
    Msg As String
    Colour As Long
    target As Long
    TargetType As Byte
    timer As Long
    active As Boolean
End Type

Public Type TextColourRec
    text As String
    Colour As Long
End Type

Public Type GeomRec
    Top As Long
    Left As Long
    Height As Long
    Width As Long
End Type

Public Type WeatherParticleRec
    Type As Long
    X As Long
    Y As Long
    Velocity As Long
    InUse As Long
End Type

Public Type ParticulaRec
    Type As Long
    X As Long
    Y As Long
    Movimento As Long
    InUse As Long
    Dir As Byte
    Opacidade As Byte
    Tamanho As Byte
    TempoG As Long
    Rotação As Long
    RotaçãoOld As Long
    Cor As Byte
    DirUp As Byte
End Type

Public Type MapSoundRec
    X As Long
    Y As Long
    SoundHandle As Long
    InUse As Boolean
    Channel As Long
End Type

