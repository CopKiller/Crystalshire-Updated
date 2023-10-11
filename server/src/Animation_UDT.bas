Attribute VB_Name = "Animation_UDT"
Option Explicit

Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public EmptyAnimation As AnimationRec

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
End Type
