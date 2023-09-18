Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public ResourceCache(1 To MAX_MAPS) As ResourceCacheRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public Options As OptionsRec

Private Type OptionsRec
    MOTD As String
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Private Type Cache
    Data() As Byte
End Type

Public Type SpellBufferRec
    Spell As Long
    Timer As Long
    target As Long
    tType As Byte
End Type

Private Type TempEventRec
    x As Long
    y As Long
    SelfSwitch As Byte
End Type

Private Type EventCommandRec
    Type As Byte
    Text As String
    colour As Long
    Channel As Byte
    targetType As Byte
    target As Long
End Type

Private Type EventPageRec
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

Private Type EventRec
    Name As String
    x As Long
    y As Long
    PageCount As Long
    EventPage() As EventPageRec
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    costitem As Long
    costvalue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Type As Byte
    mpCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    x As Long
    y As Long
    dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    
    ' ranking
    UniqueIndex As Long
    NextRank As Long
    NextUses As Long
End Type

Private Type ResourceCacheRec
    Resource_Count As Long
    ResourceData() As MapResourceRec
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    ResourceType As Byte
    ResourceImage As Long
    ExhaustedImage As Long
    ItemReward As Long
    ToolRequired As Long
    health As Long
    RespawnTime As Long
    WalkThrough As Boolean
    Animation As Long
End Type

