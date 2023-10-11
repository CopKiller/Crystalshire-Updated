Attribute VB_Name = "Animation_UDT"
Option Explicit

Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec

Public EmptyAnimation As AnimationRec
Public EmptyAnimInstance As AnimInstanceRec

Private Type AnimationRec
    Name As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
End Type

Private Type AnimInstanceRec
    Animation As Long
    x As Long
    y As Long
    ' used for locking to players/npcs
    lockindex As Long
    LockType As Byte
    isCasting As Byte
    ' timing
    timer(0 To 1) As Long
    ' rendering check
    Used(0 To 1) As Boolean
    ' counting the loop
    LoopIndex(0 To 1) As Long
    FrameIndex(0 To 1) As Long
End Type
