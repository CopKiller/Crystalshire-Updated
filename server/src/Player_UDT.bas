Attribute VB_Name = "Player_UDT"
Option Explicit

Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Class() As ClassRec

Public EmptyPlayer As PlayerRec
Public EmptyTempPlayer As TempPlayerRec
Public EmptyParty As PartyRec
Public EmptyClass As ClassRec

Public Type PlayerSpellRec
    Spell As Long
    Uses As Long
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
    Bound As Byte
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Public Type DoTRec
    Used As Boolean
    Spell As Long
    Timer As Long
    Caster As Long
    StartTime As Long
End Type

Private Type PlayerMission
    ID As Long
    Count As Long
End Type

Public Type PlayerRec
    ' General
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Class As Long
    Sprite As Long
    Level As Byte
    exp As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As PlayerSpellRec
    Bank(1 To MAX_BANK) As PlayerInvRec
    
    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
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
    
    ' Tutorial
    TutorialState As Byte
    
    ' Banned
    isBanned As Byte
    isMuted As Byte
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Private Type ProjectileRec

    Owner As Long
    TravelTime As Long
    direction As Long
    X As Double
    Y As Double
    StartX As Double
    StartY As Double
    pic As Long
    Range As Long
    Damage As Long
    Speed As Long
    ItemAmmo As Long

End Type

Public Type TempPlayerRec
    ' Non saved local vars
    buffer As clsBuffer
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    targetType As Byte
    Target As Long
    SpellCastType As Long
    Projectile(1 To MAX_PROJECTILE_PLAYER) As ProjectileRec
    GettingMap As Byte
    SpellCD(1 To MAX_PLAYER_SPELLS) As Long
    InShop As Long
    StunTimer As Long
    StunDuration As Long
    InBank As Boolean
    ' mission
    MissionRequest As Long
    ' trade
    TradeRequest As Long
    InTrade As Long
    TradeOffer(1 To MAX_INV) As PlayerInvRec
    AcceptTrade As Boolean
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' spell buffer
    spellBuffer As SpellBufferRec
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' party
    inParty As Long
    partyInvite As Long
    ' chat
    inChatWith As Long
    curChat As Long
    c_mapNum As Long
    c_mapNpcNum As Long
    ' food
    foodItem(1 To Vitals.Vital_Count - 1) As Long
    foodTick(1 To Vitals.Vital_Count - 1) As Long
    foodTimer(1 To Vitals.Vital_Count - 1) As Long
    
    ' character selection
    charNum As Long
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    
    startItemCount As Long
    StartItem() As Long
    StartValue() As Long
    
    startSpellCount As Long
    StartSpell() As Long
End Type
