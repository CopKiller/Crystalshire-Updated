Attribute VB_Name = "Client_Input"
Option Explicit
' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

' Actual input
Public Sub CheckKeys()

    ' exit out if dialogue
    If diaIndex > 0 Then Exit Sub
    If GetAsyncKeyState(VK_UP) >= 0 Then DirUp = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then DirDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then DirLeft = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then DirRight = False
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False
    If GetAsyncKeyState(VK_TAB) >= 0 Then tabDown = False
End Sub

Public Sub CheckInputKeys()

    ' exit out if dialogue
    If diaIndex > 0 Then Exit Sub
    
    ' exit out if talking
    If Windows(GetWindowIndex("winChat")).Window.visible Then Exit Sub
    
    ' continue
    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If

    If GetKeyState(vbKeyControl) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If

    If GetKeyState(vbKeyTab) < 0 Then
        tabDown = True
    Else
        tabDown = False
    End If

    'Move Up
    If Not chatOn Then
        If GetKeyState(vbKeySpace) < 0 Then
            CheckMapGetItem
        End If

        Call SetMoveDirection
    End If

End Sub

Public Sub SetMoveDirection()
    If GetKeyState(VK_UP) < 0 Then
        DirUp = True
    End If
    
    If GetKeyState(VK_DOWN) < 0 Then
        DirDown = True
    End If
    
    If GetKeyState(VK_LEFT) < 0 Then
        DirLeft = True
    End If
    
    If GetKeyState(VK_RIGHT) < 0 Then
        DirRight = True
    End If
End Sub

Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)
    Dim chatText As String, Name As String, i As Long, N As Long, Command() As String, buffer As clsBuffer, tmpNum As Long

    
    ' check if we're skipping video
    If KeyAscii = vbKeyEscape Then
        ' hide options screen
        HideWindow GetWindowIndex("winOptions")
        CloseComboMenu
        ' handle the video
        If videoPlaying Then
            videoPlaying = False
            fadeAlpha = 0
            frmMain.picIntro.visible = False
            StopIntro
            Exit Sub
        End If
        If Windows(GetWindowIndex("winEscMenu")).Window.visible Then
            ' hide it
            HideWindow GetWindowIndex("winBlank")
            HideWindow GetWindowIndex("winEscMenu")
            Exit Sub
        Else
            ' show them
            ShowWindow GetWindowIndex("winBlank"), True
            ShowWindow GetWindowIndex("winEscMenu"), True
            Exit Sub
        End If
    End If
    
    If InGame Then
    chatText = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text
    End If
    
    ' Do we have an active window
    If activeWindow > 0 Then
        ' make sure it's visible
        If Windows(activeWindow).Window.visible Then
            ' Do we have an active control
            If Windows(activeWindow).activeControl > 0 Then
                ' Do our thing
                With Windows(activeWindow).Controls(Windows(activeWindow).activeControl)
                    ' Handle input
                    Select Case KeyAscii
                        Case vbKeyBack
                            If LenB(.text) > 0 Then
                                .text = Left$(.text, Len(.text) - 1)
                            End If
                        Case vbKeyReturn
                            ' override for function callbacks
                            If .entCallBack(EntityStates.Enter) > 0 Then
                                entCallBack .entCallBack(EntityStates.Enter), activeWindow, Windows(activeWindow).activeControl, 0, 0
                                Exit Sub
                            Else
                                N = 0
                                For i = Windows(activeWindow).ControlCount To 1 Step -1
                                    If i > Windows(activeWindow).activeControl Then
                                        If SetActiveControl(activeWindow, i) Then N = i
                                    End If
                                Next
                                If N = 0 Then
                                    For i = Windows(activeWindow).ControlCount To 1 Step -1
                                        SetActiveControl activeWindow, i
                                    Next
                                End If
                            End If
                        Case vbKeyTab
                            N = 0
                            For i = 1 To Windows(activeWindow).ControlCount
                                If i > Windows(activeWindow).activeControl Then
                                    If SetActiveControl(activeWindow, i) Then N = i: Exit Sub
                                End If
                            Next
                            If N = 0 Then
                                For i = Windows(activeWindow).ControlCount To 1 Step -1
                                    SetActiveControl activeWindow, i
                                Next
                            End If
                        Case Else
                            .text = .text & ChrW$(KeyAscii)
                    End Select
                    ' exit out early - if not chatting
                    If Windows(activeWindow).Window.Name <> "winChat" Then Exit Sub
                End With
            End If
        End If
    End If

    ' exit out early if we're not ingame
    If Not InGame Then Exit Sub
    
    Select Case KeyAscii
        Case vbKeyEscape
            ' hide options screen
            HideWindow GetWindowIndex("winOptions")
            CloseComboMenu
            ' hide/show chat window
            If Windows(GetWindowIndex("winChat")).Window.visible Then
                Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
                HideChat
                inSmallChat = True
                Exit Sub
            End If
            
            If Windows(GetWindowIndex("winEscMenu")).Window.visible Then
                ' hide it
                HideWindow GetWindowIndex("winBlank")
                HideWindow GetWindowIndex("winEscMenu")
            Else
                ' show them
                ShowWindow GetWindowIndex("winBlank"), True
                ShowWindow GetWindowIndex("winEscMenu"), True
            End If
            ' exit out early
            Exit Sub
        Case 105
            ' hide/show inventory
            If Not Windows(GetWindowIndex("winChat")).Window.visible Then btnMenu_Inv
        Case 99
            ' hide/show inventory
            If Not Windows(GetWindowIndex("winChat")).Window.visible Then btnMenu_Char
        Case 109
            ' hide/show skills
            If Not Windows(GetWindowIndex("winChat")).Window.visible Then btnMenu_Skills
    End Select
    
    ' handles hotbar
    If inSmallChat Then
        For i = 1 To 9
            If KeyAscii = 48 + i Then
                SendHotbarUse i
            End If
            If KeyAscii = 48 Then SendHotbarUse 10
        Next
    End If

    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
        If Windows(GetWindowIndex("winChatSmall")).Window.visible Then
            ShowChat
            inSmallChat = False
            Exit Sub
        End If
    
        ' Broadcast message
        If Left$(chatText, 1) = "'" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call BroadcastMsg(chatText)
            End If

            Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
            HideChat
            Exit Sub
        End If

        ' Emote message
        If Left$(chatText, 1) = "-" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call EmoteMsg(chatText)
            End If

            Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
            HideChat
            Exit Sub
        End If

        ' Player message
        If Left$(chatText, 1) = "!" Then
            Exit Sub
            chatText = Mid$(chatText, 2, Len(chatText) - 1)
            Name = vbNullString
            ' Get the desired player from the user text
            tmpNum = Len(chatText)

            For i = 1 To tmpNum

                If Mid$(chatText, i, 1) <> Space$(1) Then
                    Name = Name & Mid$(chatText, i, 1)
                Else
                    Exit For
                End If

            Next

            chatText = Mid$(chatText, i, Len(chatText) - 1)

            ' Make sure they are actually sending something
            If Len(chatText) - i > 0 Then
                chatText = Mid$(chatText, i + 1, Len(chatText) - i)
                ' Send the message to the player
                Call PlayerMsg(chatText, Name)
            Else
                Call AddText("Usage: !playername (message)", AlertColor)
            End If

            Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
            HideChat
            Exit Sub
        End If

        If Left$(chatText, 1) = "/" Then
            Command = Split(chatText, Space$(1))

            Select Case Command(0)

                Case "/help"
                    Call AddText("Social Commands:", HelpColor)
                    Call AddText("'msghere = Global Message", HelpColor)
                    Call AddText("-msghere = Emote Message", HelpColor)
                    Call AddText("!namehere msghere = Player Message", HelpColor)
                    Call AddText("Available Commands: /who, /fps, /fpslock, /gui, /maps", HelpColor)

                Case "/maps"
                    ClearMapCache

                Case "/gui"
                    hideGUI = Not hideGUI

                Case "/info"

                    ' Checks to make sure we have more than one string in the array
                    If UBound(Command) < 1 Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo continue
                    End If

                    Set buffer = New clsBuffer
                    buffer.WriteLong CPlayerInfoRequest
                    buffer.WriteString Command(1)
                    SendData buffer.ToArray()
                    buffer.Flush: Set buffer = Nothing

                    ' Whos Online
                Case "/who"
                    SendWhosOnline

                    ' Checking fps
                Case "/fps"
                    BFPS = Not BFPS

                    ' toggle fps lock
                Case "/fpslock"
                    Options.FPSLock = Not Options.FPSLock
                    SaveOptions

                    ' Request stats
                Case "/stats"
                    Set buffer = New clsBuffer
                    buffer.WriteLong CGetStats
                    SendData buffer.ToArray()
                    buffer.Flush: Set buffer = Nothing

                    ' // Monitor Admin Commands //
                    ' Kicking a player
                Case "/kick"

                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo continue
                    If UBound(Command) < 1 Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo continue
                    End If

                    SendKick Command(1)

                    ' // Mapper Admin Commands //
                    ' Location
                Case "/loc"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    BLoc = Not BLoc

                    ' Map Editor
                Case "/editmap"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    SendRequestEditMap

                    ' Warping to a player
                Case "/warpmeto"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo continue
                    End If

                    GettingMap = True
                    WarpMeTo Command(1)

                    ' Warping a player to you
                Case "/warptome"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo continue
                    End If

                    WarpToMe Command(1)

                    ' Warping to a map
                Case "/warpto"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo continue
                    End If

                    N = CLng(Command(1))

                    ' Check to make sure its a valid map #
                    If N > 0 And N <= MAX_MAPS Then
                        GettingMap = True
                        Call WarpTo(N)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If

                    ' Setting sprite
                Case "/setsprite"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    If UBound(Command) < 1 Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo continue
                    End If

                    SendSetSprite CLng(Command(1))

                    ' Map report
                Case "/mapreport"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    SendMapReport

                    ' Respawn request
                Case "/respawn"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    SendMapRespawn

                    ' MOTD change
                Case "/motd"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd (new motd)", AlertColor
                        GoTo continue
                    End If

                    SendMOTDChange Right$(chatText, Len(chatText) - 5)

                    ' Check the ban list
                Case "/banlist"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    SendBanList

                    ' Banning a player
                Case "/ban"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban (name)", AlertColor
                        GoTo continue
                    End If

                    SendBan Command(1)

                    ' // Developer Admin Commands //
                    ' Editing item request
                Case "/edititem"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                    SendRequestEditItem

                    ' editing conv request
                Case "/editconv"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                    SendRequestEditConv

                    ' Editing animation request
                Case "/editanimation"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                    SendRequestEditAnimation

                    ' Editing npc request
                Case "/editnpc"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                    SendRequestEditNpc

                Case "/editresource"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                    SendRequestEditResource

                    ' Editing shop request
                Case "/editshop"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                    SendRequestEditShop

                    ' Editing spell request
                Case "/editspell"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                    SendRequestEditSpell

                    ' // Creator Admin Commands //
                    ' Giving another player access
                Case "/setaccess"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue
                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo continue
                    End If

                    SendSetAccess Command(1), CLng(Command(2))

                    ' Ban destroy
                Case "/destroybanlist"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue
                    SendBanDestroy

                    ' Packet debug mode
                Case "/debug"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue
                    DEBUG_MODE = (Not DEBUG_MODE)

                Case Else
                    AddText "Not a valid command!", HelpColor
            End Select

            'continue label where we go instead of exiting the sub
continue:
            Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
            HideChat
            Exit Sub
        End If

        ' Say message
        If Len(chatText) > 0 Then
            Call SayMsg(chatText)
        End If

        Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
        
        ' hide/show chat window
        If Windows(GetWindowIndex("winChat")).Window.visible Then HideChat
        Exit Sub
    End If
    
    ' hide/show chat window
    If Windows(GetWindowIndex("winChatSmall")).Window.visible Then
        Exit Sub
    End If
End Sub
