Attribute VB_Name = "Account_Handle"
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
        
            Buffer.Flush: Set Buffer = Nothing
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
            
            Buffer.Flush: Set Buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Public Sub HandleAddChar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim Sex As Long
    Dim Class As Long
    Dim Sprite As Long
    Dim i As Long
    Dim n As Long
    Dim charNum As Long

    If Not IsPlaying(index) Then
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        Name = Buffer.ReadString
        Sex = Buffer.ReadLong
        Class = Buffer.ReadLong
        Sprite = Buffer.ReadLong
        charNum = Buffer.ReadLong

        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(index, DIALOGUE_MSG_NAMELENGTH, MENU_NEWCHAR, False)
            Exit Sub
        End If

        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))

            If Not isNameLegal(n) Then
                Call AlertMsg(index, DIALOGUE_MSG_NAMEILLEGAL, MENU_NEWCHAR, False)
                Exit Sub
            End If

        Next

        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
            Call AlertMsg(index, DIALOGUE_MSG_CONNECTION)
            Exit Sub
        End If

        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
            Exit Sub
        End If

        ' Check if char already exists in slot
        If CharExist(index, charNum) Then
            Call AlertMsg(index, DIALOGUE_MSG_CONNECTION)
            Exit Sub
        End If

        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMsg(index, DIALOGUE_MSG_NAMETAKEN, MENU_NEWCHAR, False)
            Exit Sub
        End If

        ' Everything went ok, add the character
        Call AddChar(index, Name, Sex, Class, Sprite, charNum)
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(index) & "'s account.", PLAYER_LOG)
        ' log them in!!
        UseChar index, charNum
        
        Buffer.Flush: Set Buffer = Nothing
    End If

End Sub

Public Sub HandleUseChar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, charNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    charNum = Buffer.ReadLong
    UseChar index, charNum
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub HandleDelChar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, charNum As Long
    Dim Login As String, charName As String, filename As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    charNum = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    
    If charNum < 0 Or charNum > MAX_CHARS Then Exit Sub
    
    ' clear the character
    Login = Trim$(Account(index).Login)
    filename = App.Path & "\data\accounts\" & SanitiseString(Login) & ".ini"
    charName = GetVar(filename, "CHAR" & charNum, "Name")
    DeleteCharacter Login, charNum
    
    ' remove the character name from the list
    DeleteName charName
    
    ' send to portal again
    'AlertMsg index, DIALOGUE_MSG_DELCHAR, MENU_LOGIN
    SendPlayerChars index
End Sub

Public Sub HandleMergeAccounts(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, username As String, Password As String, oldPass As String, oldName As String
    Dim filename As String, i As Long, charNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    username = Buffer.ReadString
    Password = Buffer.ReadString
    Buffer.Flush: Set Buffer = Nothing
    
    ' Check versions
    If Len(Trim$(username)) < 3 Or Len(Trim$(Password)) < 3 Then
        Call AlertMsg(index, DIALOGUE_MSG_USERLENGTH, MENU_MERGE, False)
        Exit Sub
    End If
    
    ' check if the player has a slot free
    filename = App.Path & "\data\accounts\" & SanitiseString(Trim$(Account(index).Login)) & ".ini"
    ' exit out if we can't find the player's ACTUAL account
    If Not FileExist(filename) Then
        AlertMsg index, DIALOGUE_MSG_CONNECTION
        Exit Sub
    End If
    For i = MAX_CHARS To 1 Step -1
        ' check if the chars have a name
        If LenB(Trim$(GetVar(filename, "CHAR" & i, "Name"))) < 1 Then
            charNum = i
        End If
    Next
    ' if charnum is defaulted to 0 then no chars available - exit out
    If charNum = 0 Then
        AlertMsg index, DIALOGUE_MSG_CONNECTION
        Exit Sub
    End If
    
    ' check if the user exists
    If Not OldAccount_Exist(username) Then
        Call AlertMsg(index, DIALOGUE_MSG_WRONGPASS, MENU_MERGE, False)
        Exit Sub
    End If
    
    ' check if passwords match
    filename = App.Path & "\data\accounts\old\" & SanitiseString(username) & ".ini"
    oldPass = GetVar(filename, "ACCOUNT", "Password")
    If Not Password = oldPass Then
        Call AlertMsg(index, DIALOGUE_MSG_WRONGPASS, MENU_MERGE, False)
        Exit Sub
    End If

    ' get the old name
    oldName = GetVar(filename, "ACCOUNT", "Name")
    
    ' make sure it's available
    If FindChar(oldName) Then
        Call AlertMsg(index, DIALOGUE_MSG_MERGENAME, MENU_MERGE, False)
        Exit Sub
    End If
    
    ' fill the character slot with the old character
    MergeAccount index, charNum, username
End Sub
