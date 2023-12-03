Attribute VB_Name = "Conv_UDT"
Option Explicit

Public Conv(1 To MAX_CONVS) As ConvWrapperRec
Public EmptyConv As ConvWrapperRec

Public Type ConvRec
    Conv As String
    rText(1 To 4) As String
    rTarget(1 To 4) As Long
    EventType As Long
    EventNum As Long
End Type

Private Type ConvWrapperRec
    Name As String * NAME_LENGTH
    chatCount As Long
    Conv() As ConvRec
End Type
