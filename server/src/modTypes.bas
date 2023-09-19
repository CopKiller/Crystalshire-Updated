Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public Options As OptionsRec

Private Type OptionsRec
    MOTD As String
End Type

Public Type SpellBufferRec
    Spell As Long
    Timer As Long
    target As Long
    tType As Byte
End Type

