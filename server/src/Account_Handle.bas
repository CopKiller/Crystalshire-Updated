Attribute VB_Name = "Account_Handle"
Option Explicit

Public Sub HandleNewAccount(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Name As String, Pass As String, Code As String
  '  MsgBox "ok"
If Not IsPlaying(index) Then
    Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Pass = Buffer.ReadString
            Code = Buffer.ReadString

    
    If Len(Trim$(Name)) < 3 Or Len(Trim$(Pass)) < 3 Or Len(Trim$(Code)) < 3 Then
        Call AlertMsg(index, DIALOGUE_MSG_NAMELENGTH, MENU_REGISTER)
        Exit Sub
    End If
    
    If AccountExist(Name) Then
        Call AlertMsg(index, DIALOGUE_MSG_NAMETAKEN, MENU_REGISTER)
        Exit Sub
    Else
        Call AddAccount(index, Name, Pass, Code)
        Call AlertMsg(index, DIALOGUE_ACCOUNT_CREATED, MENU_LOGIN)
    End If
    
        Set Buffer = Nothing
    Exit Sub
End If
End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Public Sub HandleDelAccount(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' No deleting accounts lOL
End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Public Sub HandleLogin(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, Name As String, i As Long, n As Long, Password As String, charNum As Long

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Check versions
            If Buffer.ReadLong <> CLIENT_MAJOR Or Buffer.ReadLong <> CLIENT_MINOR Or Buffer.ReadLong <> CLIENT_REVISION Then
                Call AlertMsg(index, DIALOGUE_MSG_OUTDATED)
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMsg(index, DIALOGUE_MSG_REBOOTING)
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Then
                Call AlertMsg(index, DIALOGUE_MSG_USERLENGTH, MENU_LOGIN)
                Exit Sub
            End If
            
            If Password = vbNullString Or Len(Password) < 1 Then
                Call AlertMsg(index, DIALOGUE_MSG_WRONGPASS, MENU_LOGIN)
                Exit Sub
            End If

            If IsMultiAccounts(Name) Then
                Call AlertMsg(index, DIALOGUE_MSG_CONNECTION, MENU_LOGIN)
                Exit Sub
            End If
            
            If Not AccountExist(Name) Then
                Call AlertMsg(index, DIALOGUE_MSG_WRONGPASS, MENU_LOGIN)
                Exit Sub
            End If
            
            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, DIALOGUE_MSG_WRONGPASS, MENU_LOGIN)
                Exit Sub
            End If

            ' Load the account
            Call LoadAccount(index, Name)
            
            ' make sure they're not banned
            If isBanned_Account(index) Then
                Call AlertMsg(index, DIALOGUE_MSG_BANNED)
                Exit Sub
            End If

            ' send them to the character portal
            If Not IsPlaying(index) Then
                Call SendPlayerChars(index)
                Call SendNewCharClasses(index)
            End If
            
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".")
            
            ' Update list players from server
            frmServer.lvwInfo.ListItems(index).SubItems(1) = GetPlayerIP(index)
            frmServer.lvwInfo.ListItems(index).SubItems(2) = GetPlayerLogin(index)
            
            Set Buffer = Nothing
        End If
    End If

End Sub
