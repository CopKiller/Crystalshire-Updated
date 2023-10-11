Attribute VB_Name = "Spell_UDT"
Option Explicit

Public Spell(1 To MAX_SPELLS) As SpellRec
Public EmptySpell As SpellRec

Public Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH

    Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    icon As Long
    Map As Long
    x As Long
    y As Long
    Dir As Byte
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

