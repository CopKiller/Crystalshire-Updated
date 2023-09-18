Attribute VB_Name = "Conv_UDT"
Option Explicit

Public Conv(1 To MAX_CONVS) As ConvWrapperRec

Private Type ConvRec
    Conv As String
    rText(1 To 4) As String
    rTarget(1 To 4) As Long
    Event As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
End Type

Private Type ConvWrapperRec
    Name As String * NAME_LENGTH
    chatCount As Long
    Conv() As ConvRec
End Type
