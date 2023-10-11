Attribute VB_Name = "NPC_UDT"
Option Explicit

Public Npc(1 To MAX_NPCS) As NpcRec
Public EmptyNpc As NpcRec

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Mission As Long
    Range As Byte
    Stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    exp As Long
    Animation As Long
    damage As Long
    Level As Long
    Conv As Long
    
    ' Npc drops
    DropChance(1 To MAX_NPC_DROPS) As Double
    DropItem(1 To MAX_NPC_DROPS) As Byte
    DropItemValue(1 To MAX_NPC_DROPS) As Integer
    
    ' Casting
    Spirit As Long
    Spell(1 To MAX_NPC_SPELLS) As Long
End Type
