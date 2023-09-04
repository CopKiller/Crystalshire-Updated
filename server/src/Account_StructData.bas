Attribute VB_Name = "Account_UDT"
Option Explicit

Public Account(1 To MAX_PLAYERS) As AccountRec

Private Type AccountRec
    ' Account
    Login As String * ACCOUNT_LENGTH
    Password As String * ACCOUNT_LENGTH
    Mail As String * EMAIL_LENGTH
End Type
