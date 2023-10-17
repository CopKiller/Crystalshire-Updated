Attribute VB_Name = "modInterfaceEvents"
Option Explicit
Public Declare Sub GetCursorPos Lib "user32" (lpPoint As POINTAPI)
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function entCallBack Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Window As Long, ByRef Control As Long, ByVal forced As Long, ByVal lParam As Long) As Long
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public lastMouseX As Long, lastMouseY As Long
Public currMouseX As Long, currMouseY As Long
Public clickedX As Long, clickedY As Long
Public mouseClick(1 To 2) As Long
Public lastMouseClick(1 To 2) As Long

Public GlobalCaptcha As Long


Public Function MouseX(Optional ByVal hwnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint

    If hwnd Then ScreenToClient hwnd, lpPoint
    MouseX = lpPoint.x
End Function

Public Function MouseY(Optional ByVal hwnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint

    If hwnd Then ScreenToClient hwnd, lpPoint
    MouseY = lpPoint.y
End Function

Public Sub HandleMouseInput()
    Dim entState As entStates, I As Long, x As Long
    
    ' exit out if we're playing video
    If videoPlaying Then Exit Sub
    
    ' set values
    lastMouseX = currMouseX
    lastMouseY = currMouseY
    currMouseX = MouseX(frmMain.hwnd)
    currMouseY = MouseY(frmMain.hwnd)
    GlobalX = currMouseX
    GlobalY = currMouseY
    lastMouseClick(VK_LBUTTON) = mouseClick(VK_LBUTTON)
    lastMouseClick(VK_RBUTTON) = mouseClick(VK_RBUTTON)
    mouseClick(VK_LBUTTON) = GetAsyncKeyState(VK_LBUTTON)
    mouseClick(VK_RBUTTON) = GetAsyncKeyState(VK_RBUTTON)
    
    ' Hover
    entState = entStates.Hover

    ' MouseDown
    If (mouseClick(VK_LBUTTON) And lastMouseClick(VK_LBUTTON) = 0) Or (mouseClick(VK_RBUTTON) And lastMouseClick(VK_RBUTTON) = 0) Then
        clickedX = currMouseX
        clickedY = currMouseY
        entState = entStates.MouseDown
        ' MouseUp
    ElseIf (mouseClick(VK_LBUTTON) = 0 And lastMouseClick(VK_LBUTTON)) Or (mouseClick(VK_RBUTTON) = 0 And lastMouseClick(VK_RBUTTON)) Then
        entState = entStates.MouseUp
        ' MouseMove
    ElseIf (currMouseX <> lastMouseX) Or (currMouseY <> lastMouseY) Then
        entState = entStates.MouseMove
    End If

    ' Handle everything else
    If Not HandleGuiMouse(entState) Then
        ' reset /all/ control mouse events
        For I = 1 To WindowCount
            For x = 1 To Windows(I).ControlCount
                Windows(I).Controls(x).state = Normal
            Next
        Next
        If InGame Then
            If entState = entStates.MouseDown Then
                ' Handle events
                If currMouseX >= 0 And currMouseX <= frmMain.ScaleWidth Then
                    If currMouseY >= 0 And currMouseY <= frmMain.ScaleHeight Then
                    
                        If InMapEditor Then
                            If (mouseClick(VK_LBUTTON) And lastMouseClick(VK_LBUTTON) = 0) Then
                                If frmEditor_Map.optEvents.value Then
                                    selTileX = CurX
                                    selTileY = CurY
                                Else
                                    Call MapEditorMouseDown(vbLeftButton, GlobalX, GlobalY, False)
                                End If
                            ElseIf (mouseClick(VK_RBUTTON) And lastMouseClick(VK_RBUTTON) = 0) Then
                                If Not frmEditor_Map.optEvents.value Then
                                    Call MapEditorMouseDown(vbRightButton, GlobalX, GlobalY, False)
                                End If
                            End If
                        Else
                            ' left click
                            If (mouseClick(VK_LBUTTON) And lastMouseClick(VK_LBUTTON) = 0) Then
                                ' targetting
                                FindTarget
                                ' right click
                            ElseIf (mouseClick(VK_RBUTTON) And lastMouseClick(VK_RBUTTON) = 0) Then
                                If ShiftDown Then
                                    ' admin warp if we're pressing shift and right clicking
                                    If GetPlayerAccess(MyIndex) >= 2 Then AdminWarp CurX, CurY
                                    Exit Sub
                                End If
                                ' right-click menu
                                For I = 1 To Player_HighIndex
                                    If IsPlaying(I) Then
                                        If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                            If GetPlayerX(I) = CurX And GetPlayerY(I) = CurY Then
                                                ShowPlayerMenu I, currMouseX, currMouseY
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
            ElseIf entState = entStates.MouseMove Then
                GlobalX_Map = GlobalX + (TileView.Left * PIC_X) + Camera.Left
                GlobalY_Map = GlobalY + (TileView.Top * PIC_Y) + Camera.Top
                ' Handle the events
                CurX = TileView.Left + ((currMouseX + Camera.Left) \ PIC_X)
                CurY = TileView.Top + ((currMouseY + Camera.Top) \ PIC_Y)

                If InMapEditor Then
                    If (mouseClick(VK_LBUTTON)) Then
                        If Not frmEditor_Map.optEvents.value Then
                            Call MapEditorMouseDown(vbLeftButton, CurX, CurY, False)
                        End If
                    ElseIf (mouseClick(VK_RBUTTON)) Then
                        If Not frmEditor_Map.optEvents.value Then
                            Call MapEditorMouseDown(vbRightButton, CurX, CurY, False)
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Function HandleGuiMouse(entState As entStates) As Boolean
    Dim I As Long, curWindow As Long, curControl As Long, callback As Long, x As Long
    
    ' if hiding gui
    If hideGUI = True Or InMapEditor Then Exit Function

    ' Find the container
    For I = 1 To WindowCount
        With Windows(I).Window
            If .enabled And .visible Then
                If .state <> entStates.MouseDown Then .state = entStates.Normal
                If currMouseX >= .Left And currMouseX <= .Width + .Left Then
                    If currMouseY >= .Top And currMouseY <= .Height + .Top Then
                        ' set the combomenu
                        If .design(0) = DesignTypes.desComboMenuNorm Then
                            ' set the hover menu
                            If entState = MouseMove Or entState = Hover Then
                                ComboMenu_MouseMove I
                            ElseIf entState = MouseDown Then
                                ComboMenu_MouseDown I
                            End If
                        End If
                        ' everything else
                        If curWindow = 0 Then curWindow = I
                        If .zOrder > Windows(curWindow).Window.zOrder Then curWindow = I
                    End If
                End If
                If entState = entStates.MouseMove Then
                    If .canDrag Then
                        If .state = entStates.MouseDown Then
                            .Left = Clamp(.Left + ((currMouseX - .Left) - .movedX), 0, ScreenWidth - .Width)
                            .Top = Clamp(.Top + ((currMouseY - .Top) - .movedY), 0, ScreenHeight - .Height)
                        End If
                    End If
                End If
            End If
        End With
    Next

    ' Handle any controls first
    If curWindow Then
        ' reset /all other/ control mouse events
        For I = 1 To WindowCount
            If I <> curWindow Then
                For x = 1 To Windows(I).ControlCount
                    Windows(I).Controls(x).state = Normal
                Next
            End If
        Next
        For I = 1 To Windows(curWindow).ControlCount
            With Windows(curWindow).Controls(I)
                If .enabled And .visible Then
                    If .state <> entStates.MouseDown Then .state = entStates.Normal
                    If currMouseX >= .Left + Windows(curWindow).Window.Left And currMouseX <= .Left + .Width + Windows(curWindow).Window.Left Then
                        If currMouseY >= .Top + Windows(curWindow).Window.Top And currMouseY <= .Top + .Height + Windows(curWindow).Window.Top Then
                            If curControl = 0 Then curControl = I
                            If .zOrder > Windows(curWindow).Controls(curControl).zOrder Then curControl = I
                        End If
                    End If
                    If entState = entStates.MouseMove Then
                        If .canDrag Then
                            If .state = entStates.MouseDown Then
                                .Left = Clamp(.Left + ((currMouseX - .Left) - .movedX), 0, Windows(curWindow).Window.Width - .Width)
                                .Top = Clamp(.Top + ((currMouseY - .Top) - .movedY), 0, Windows(curWindow).Window.Height - .Height)
                            End If
                        End If
                    End If
                End If
            End With
        Next
        ' Handle control
        If curControl Then
            HandleGuiMouse = True
            With Windows(curWindow).Controls(curControl)
                If .state <> entStates.MouseDown Then
                    If entState <> entStates.MouseMove Then
                        .state = entState
                    Else
                        .state = entStates.Hover
                    End If
                End If
                If entState = entStates.MouseDown Then
                    If .canDrag Then
                        .movedX = clickedX - .Left
                        .movedY = clickedY - .Top
                    End If
                    ' toggle boxes
                    Select Case .Type
                        Case EntityTypes.entCheckbox
                            ' grouped boxes
                            If .group > 0 Then
                                If .value = 0 Then
                                    For I = 1 To Windows(curWindow).ControlCount
                                        If Windows(curWindow).Controls(I).Type = EntityTypes.entCheckbox Then
                                            If Windows(curWindow).Controls(I).group = .group Then
                                                Windows(curWindow).Controls(I).value = 0
                                            End If
                                        End If
                                    Next
                                    .value = 1
                                End If
                            Else
                                If .value = 0 Then
                                    .value = 1
                                Else
                                    .value = 0
                                End If
                            End If
                        Case EntityTypes.entCombobox
                            ShowComboMenu curWindow, curControl
                    End Select
                    ' set active input
                    SetActiveControl curWindow, curControl
                End If
                callback = .entCallBack(entState)
            End With
        Else
            ' Handle container
            With Windows(curWindow).Window
                HandleGuiMouse = True
                If .state <> entStates.MouseDown Then
                    If entState <> entStates.MouseMove Then
                        .state = entState
                    Else
                        .state = entStates.Hover
                    End If
                End If
                If entState = entStates.MouseDown Then
                    If .canDrag Then
                        .movedX = clickedX - .Left
                        .movedY = clickedY - .Top
                    End If
                End If
                callback = .entCallBack(entState)
            End With
        End If
        ' bring to front
        If entState = entStates.MouseDown Then
            UpdateZOrder curWindow
            activeWindow = curWindow
        End If
        ' call back
        If callback <> 0 Then entCallBack callback, curWindow, curControl, 0, 0
    End If

    ' Reset
    If entState = entStates.MouseUp Then ResetMouseDown
End Function

Public Sub ResetGUI()
    Dim I As Long, x As Long

    For I = 1 To WindowCount

        If Windows(I).Window.state <> MouseDown Then Windows(I).Window.state = Normal

        For x = 1 To Windows(I).ControlCount

            If Windows(I).Controls(x).state <> MouseDown Then Windows(I).Controls(x).state = Normal
        Next
    Next

End Sub

Public Sub ResetMouseDown()
    Dim callback As Long
    Dim I As Long, x As Long

    For I = 1 To WindowCount

        With Windows(I)
            .Window.state = entStates.Normal
            callback = .Window.entCallBack(entStates.Normal)

            If callback <> 0 Then entCallBack callback, I, 0, 0, 0

            For x = 1 To .ControlCount
                .Controls(x).state = entStates.Normal
                callback = .Controls(x).entCallBack(entStates.Normal)

                If callback <> 0 Then entCallBack callback, I, x, 0, 0
            Next

        End With

    Next

End Sub
' ################## ##
' ## REGISTER WINDOW ##
' #####################
Public Sub btnRegister_Click()
    HideWindows
    RenCaptcha
    ClearRegisterTexts
    ShowWindow GetWindowIndex("winRegister")
End Sub
Sub ClearRegisterTexts()
Dim I As Long
    With Windows(GetWindowIndex("winRegister"))
        .Controls(GetControlIndex("winRegister", "txtAccount")).text = vbNullString
        .Controls(GetControlIndex("winRegister", "txtPass")).text = vbNullString
        .Controls(GetControlIndex("winRegister", "txtPass2")).text = vbNullString
        .Controls(GetControlIndex("winRegister", "txtCode")).text = vbNullString
        .Controls(GetControlIndex("winRegister", "txtCaptcha")).text = vbNullString
    For I = 0 To 6
        .Controls(GetControlIndex("winRegister", "picCaptcha")).image(I) = Tex_Captcha(GlobalCaptcha)
    Next
    End With
End Sub
Sub RenCaptcha()
Dim N As Long
    N = Int(Rnd * (Count_Captcha - 1)) + 1
GlobalCaptcha = N
End Sub
Public Sub btnSendRegister_Click()
    Dim User As String, Pass As String, pass2 As String, Code As String, Captcha As String

    With Windows(GetWindowIndex("winRegister"))
        User = .Controls(GetControlIndex("winRegister", "txtAccount")).text
        Pass = .Controls(GetControlIndex("winRegister", "txtPass")).text
        pass2 = .Controls(GetControlIndex("winRegister", "txtPass2")).text
        Code = .Controls(GetControlIndex("winRegister", "txtCode")).text
        Captcha = .Controls(GetControlIndex("winRegister", "txtCaptcha")).text
    End With

If Trim$(Pass) <> Trim$(pass2) Then
   Call Dialogue("Register", "Falha ao criar conta.", "A confirmação não confere com a senha!", TypeDELCHAR, StyleOKAY, 1)
    Exit Sub
End If

If User = vbNullString Or Pass = vbNullString Or pass2 = vbNullString Or Code = vbNullString Then
    Call Dialogue("Register", "Falha ao criar conta.", "Nenhum campo pode ficar em branco!", TypeDELCHAR, StyleOKAY, 1)
    Exit Sub
End If

If Trim$(Captcha) <> Trim$(GetCaptcha) Then
    RenCaptcha
    ClearRegisterTexts
    Call Dialogue("Register", "Falha ao criar conta.", "Captcha Incorreto!", TypeDELCHAR, StyleOKAY, 1)
    Exit Sub
End If

 SendRegister User, Pass, Code
End Sub
Public Sub btnReturnMain_Click()
    HideWindows
    ShowWindow GetWindowIndex("winLogin")
End Sub
Function GetCaptcha() As String
Select Case GlobalCaptcha
    
    Case 1
        GetCaptcha = "EVqu"
    Case 2
        GetCaptcha = "8Nmv"
    Case 3
         GetCaptcha = "1Swi"
    Case 4
         GetCaptcha = "Vunk"
    Case 5
         GetCaptcha = "isKD"
    Case 6
         GetCaptcha = "eYX2"
End Select
End Function

' ##################
' ## Login Window ##
' ##################

Public Sub btnLogin_Click()
    Dim User As String, Pass As String
    
    With Windows(GetWindowIndex("winLogin"))
        User = .Controls(GetControlIndex("winLogin", "txtUser")).text
        Pass = .Controls(GetControlIndex("winLogin", "txtPass")).text
    End With
    
    Login User, Pass
End Sub

Public Sub chkSaveUser_Click()

    With Windows(GetWindowIndex("winLogin")).Controls(GetControlIndex("winLogin", "chkSaveUser"))
        If .value = 0 Then ' set as false
            Options.SaveUser = 0
            Options.Username = vbNullString
            SaveOptions
        Else
            Options.SaveUser = 1
            SaveOptions
        End If
    End With
End Sub

' #######################
' ## Characters Window ##
' #######################

Public Sub Chars_DrawFace()
Dim Xo As Long, Yo As Long, imageFace As Long, imageChar As Long, x As Long, I As Long
    
    Xo = Windows(GetWindowIndex("winCharacters")).Window.Left
    Yo = Windows(GetWindowIndex("winCharacters")).Window.Top
    
    x = Xo + 24
    For I = 1 To MAX_CHARS
        If LenB(Trim$(CharName(I))) > 0 Then
            If CharSprite(I) > 0 Then
                If Not CharSprite(I) > Count_Char And Not CharSprite(I) > Count_Face Then
                    imageFace = Tex_Face(CharSprite(I))
                    imageChar = Tex_Char(CharSprite(I))
                    RenderTexture imageFace, x, Yo + 56, 0, 0, 94, 94, 94, 94
                    RenderTexture imageChar, x - 1, Yo + 117, 32, 0, 32, 32, 32, 32
                End If
            End If
        End If
        x = x + 110
    Next
End Sub

Public Sub btnAcceptChar_1()
    SendUseChar 1
End Sub

Public Sub btnAcceptChar_2()
    SendUseChar 2
End Sub

Public Sub btnAcceptChar_3()
    SendUseChar 3
End Sub

Public Sub btnDelChar_1()
    Dialogue "Delete Character", "Deleting this character is permanent.", "Are you sure you want to delete this character?", TypeDELCHAR, StyleYESNO, 1
End Sub

Public Sub btnDelChar_2()
    Dialogue "Delete Character", "Deleting this character is permanent.", "Are you sure you want to delete this character?", TypeDELCHAR, StyleYESNO, 2
End Sub

Public Sub btnDelChar_3()
    Dialogue "Delete Character", "Deleting this character is permanent.", "Are you sure you want to delete this character?", TypeDELCHAR, StyleYESNO, 3
End Sub

Public Sub btnCreateChar_1()
    CharNum = 1
    ShowClasses
End Sub

Public Sub btnCreateChar_2()
    CharNum = 2
    ShowClasses
End Sub

Public Sub btnCreateChar_3()
    CharNum = 3
    ShowClasses
End Sub

Public Sub btnCharacters_Close()
    DestroyTCP
    HideWindows
    ShowWindow GetWindowIndex("winLogin")
End Sub

' #####################
' ## Dialogue Window ##
' #####################

Public Sub btnDialogue_Close()
    If diaStyle = StyleOKAY Then
        dialogueHandler 1
    ElseIf diaStyle = StyleYESNO Then
        dialogueHandler 3
    End If
End Sub

Public Sub Dialogue_Okay()
    dialogueHandler 1
End Sub

Public Sub Dialogue_Yes()
    dialogueHandler 2
End Sub

Public Sub Dialogue_No()
    dialogueHandler 3
End Sub

' ####################
' ## Classes Window ##
' ####################

Public Sub Classes_DrawFace()
Dim imageFace As Long, Xo As Long, Yo As Long

    Xo = Windows(GetWindowIndex("winClasses")).Window.Left
    Yo = Windows(GetWindowIndex("winClasses")).Window.Top
    
    Max_Classes = 3
    
    If newCharClass = 0 Then newCharClass = 1

    Select Case newCharClass
        Case 1 ' Warrior
            imageFace = Tex_GUI(18)
        Case 2 ' Wizard
            imageFace = Tex_GUI(19)
        Case 3 ' Whisperer
            imageFace = Tex_GUI(20)
    End Select
    
    ' render face
    RenderTexture imageFace, Xo + 14, Yo - 41, 0, 0, 256, 256, 256, 256
End Sub

Public Sub Classes_DrawText()
Dim image As Long, text As String, Xo As Long, Yo As Long, textArray() As String, I As Long, count As Long, y As Long, x As Long

    Xo = Windows(GetWindowIndex("winClasses")).Window.Left
    Yo = Windows(GetWindowIndex("winClasses")).Window.Top

    Select Case newCharClass
        Case 1 ' Warrior
            text = "The way of a warrior has never been an easy one. Skilled use of a sword is not something learnt overnight. Being able to take a decent amount of hits is important for these characters and as such they weigh a lot of importance on endurance and strength."
        Case 2 ' Wizard
            text = "Wizards are often mistrusted characters who have mastered the practise of using their own spirit to create elemental entities. Generally seen as playful and almost childish because of the huge amounts of pleasure they take from setting things on fire."
        Case 3 ' Whisperer
            text = "The art of healing is one which comes with tremendous amounts of pressure and guilt. Constantly being put under high-pressure situations where their abilities could mean the difference between life and death leads many Whisperers to insanity."
    End Select
    
    ' wrap text
    WordWrap_Array text, 200, textArray()
    ' render text
    count = UBound(textArray)
    y = Yo + 60
    For I = 1 To count
        x = Xo + 132 + (200 \ 2) - (TextWidth(font(Fonts.rockwell_15), textArray(I)) \ 2)
        RenderText font(Fonts.rockwell_15), textArray(I), x, y, White
        y = y + 14
    Next
End Sub

Public Sub btnClasses_Left()
Dim text As String
    newCharClass = newCharClass - 1
    If newCharClass <= 0 Then
        newCharClass = Max_Classes
    End If
    Windows(GetWindowIndex("winClasses")).Controls(GetControlIndex("winClasses", "lblClassName")).text = Trim$(Class(newCharClass).Name)
End Sub

Public Sub btnClasses_Right()
Dim text As String
    newCharClass = newCharClass + 1
    If newCharClass > Max_Classes Then
        newCharClass = 1
    End If
    Windows(GetWindowIndex("winClasses")).Controls(GetControlIndex("winClasses", "lblClassName")).text = Trim$(Class(newCharClass).Name)
End Sub

Public Sub btnClasses_Accept()
    HideWindow GetWindowIndex("winClasses")
    ShowWindow GetWindowIndex("winNewChar")
End Sub

Public Sub btnClasses_Close()
    HideWindows
    ShowWindow GetWindowIndex("winCharacters")
End Sub

' ###################
' ## New Character ##
' ###################

Public Sub NewChar_OnDraw()
Dim imageFace As Long, imageChar As Long, Xo As Long, Yo As Long
    
    Xo = Windows(GetWindowIndex("winNewChar")).Window.Left
    Yo = Windows(GetWindowIndex("winNewChar")).Window.Top
    
    If newCharGender = SEX_MALE Then
        imageFace = Tex_Face(Class(newCharClass).MaleSprite(newCharSprite))
        imageChar = Tex_Char(Class(newCharClass).MaleSprite(newCharSprite))
    Else
        imageFace = Tex_Face(Class(newCharClass).FemaleSprite(newCharSprite))
        imageChar = Tex_Char(Class(newCharClass).FemaleSprite(newCharSprite))
    End If
    
    ' render face
    RenderTexture imageFace, Xo + 166, Yo + 56, 0, 0, 94, 94, 94, 94
    ' render char
    RenderTexture imageChar, Xo + 166, Yo + 116, 32, 0, 32, 32, 32, 32
End Sub

Public Sub btnNewChar_Left()
Dim spriteCount As Long

    If newCharGender = SEX_MALE Then
        spriteCount = UBound(Class(newCharClass).MaleSprite)
    Else
        spriteCount = UBound(Class(newCharClass).FemaleSprite)
    End If

    If newCharSprite <= 0 Then
        newCharSprite = spriteCount
    Else
        newCharSprite = newCharSprite - 1
    End If
End Sub

Public Sub btnNewChar_Right()
Dim spriteCount As Long

    If newCharGender = SEX_MALE Then
        spriteCount = UBound(Class(newCharClass).MaleSprite)
    Else
        spriteCount = UBound(Class(newCharClass).FemaleSprite)
    End If

    If newCharSprite >= spriteCount Then
        newCharSprite = 0
    Else
        newCharSprite = newCharSprite + 1
    End If
End Sub

Public Sub chkNewChar_Male()
    newCharSprite = 1
    newCharGender = SEX_MALE
End Sub

Public Sub chkNewChar_Female()
    newCharSprite = 1
    newCharGender = SEX_FEMALE
End Sub

Public Sub btnNewChar_Cancel()
    Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "txtName")).text = vbNullString
    Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "chkMale")).value = 1
    Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "chkFemale")).value = 0
    newCharSprite = 1
    newCharGender = SEX_MALE
    HideWindows
    ShowWindow GetWindowIndex("winClasses")
End Sub

Public Sub btnNewChar_Accept()
Dim Name As String
    Name = Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "txtName")).text
    HideWindows
    AddChar Name, newCharGender, newCharClass, newCharSprite
End Sub

' ##############
' ## Esc Menu ##
' ##############

Public Sub btnEscMenu_Return()
    HideWindow GetWindowIndex("winBlank")
    HideWindow GetWindowIndex("winEscMenu")
End Sub

Public Sub btnEscMenu_Options()
    HideWindow GetWindowIndex("winEscMenu")
    ShowWindow GetWindowIndex("winOptions"), True, True
End Sub

Public Sub btnEscMenu_MainMenu()
    HideWindows
    ShowWindow GetWindowIndex("winLogin")
    Stop_Music
    ' play the menu music
    If Len(Trim$(MenuMusic)) > 0 Then Play_Music Trim$(MenuMusic)
    logoutGame
End Sub

Public Sub btnEscMenu_Exit()
    HideWindow GetWindowIndex("winBlank")
    HideWindow GetWindowIndex("winEscMenu")
    DestroyGame
End Sub

' ##########
' ## Bars ##
' ##########

Public Sub Bars_OnDraw()
    Dim Xo As Long, Yo As Long, Width As Long
    
    Xo = Windows(GetWindowIndex("winBars")).Window.Left
    Yo = Windows(GetWindowIndex("winBars")).Window.Top
    
    ' Bars
    RenderTexture Tex_GUI(27), Xo + 15, Yo + 15, 0, 0, BarWidth_GuiHP, 13, BarWidth_GuiHP, 13
    RenderTexture Tex_GUI(28), Xo + 15, Yo + 32, 0, 0, BarWidth_GuiSP, 13, BarWidth_GuiSP, 13
    RenderTexture Tex_GUI(29), Xo + 15, Yo + 49, 0, 0, BarWidth_GuiEXP, 13, BarWidth_GuiEXP, 13
End Sub

' ##########
' ## Menu ##
' ##########

Public Sub btnMenu_Char()
Dim curWindow As Long
    curWindow = GetWindowIndex("winCharacter")
    If Windows(curWindow).Window.visible Then
        HideWindow curWindow
    Else
        ShowWindow curWindow, , False
    End If
End Sub

Public Sub btnMenu_Inv()
Dim curWindow As Long
    curWindow = GetWindowIndex("winInventory")
    If Windows(curWindow).Window.visible Then
        HideWindow curWindow
    Else
        ShowWindow curWindow, , False
    End If
End Sub

Public Sub btnMenu_Skills()
Dim curWindow As Long
    curWindow = GetWindowIndex("winSkills")
    If Windows(curWindow).Window.visible Then
        HideWindow curWindow
    Else
        ShowWindow curWindow, , False
    End If
End Sub

Public Sub btnMenu_Map()
    'Windows(GetWindowIndex("winCharacter")).Window.visible = Not Windows(GetWindowIndex("winCharacter")).Window.visible
End Sub

Public Sub btnMenu_Guild()
    'Windows(GetWindowIndex("winCharacter")).Window.visible = Not Windows(GetWindowIndex("winCharacter")).Window.visible
End Sub

Public Sub btnMenu_Quest()
    'Windows(GetWindowIndex("winCharacter")).Window.visible = Not Windows(GetWindowIndex("winCharacter")).Window.visible
    Windows(GetWindowIndex("winOffer")).Window.visible = Not Windows(GetWindowIndex("winOffer")).Window.visible
End Sub

' ###############
' ##    Bank   ##
' ###############

Public Sub btnMenu_Bank()
    If Windows(GetWindowIndex("winBank")).Window.visible Then
        CloseBank
    End If

    Windows(GetWindowIndex("winBank")).Window.visible = Not Windows(GetWindowIndex("winBank")).Window.visible
End Sub

Public Sub Bank_MouseMove()
    Dim ItemNum As Long, x As Long, y As Long, I As Long

    ' exit out early if dragging
    If DragBox.Type <> part_None Then Exit Sub

    ItemNum = IsBankItem(Windows(GetWindowIndex("winBank")).Window.Left, Windows(GetWindowIndex("winBank")).Window.Top)

    If ItemNum > 0 Then

        ' make sure we're not dragging the item
        If DragBox.Type = Part_Item And DragBox.value = ItemNum Then Exit Sub
        ' calc position
        x = Windows(GetWindowIndex("winBank")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        y = Windows(GetWindowIndex("winBank")).Window.Top - 4

        ' offscreen?
        If x < 0 Then
            ' switch to right
            x = Windows(GetWindowIndex("winBank")).Window.Left + Windows(GetWindowIndex("winBank")).Window.Width
        End If

        ' go go go
        ShowItemDesc x, y, Bank.Item(ItemNum).num, False
    End If
End Sub

Public Sub Bank_MouseDown()
    Dim BankSlot As Long, winIndex As Long, I As Long

    ' is there an item?
    BankSlot = IsBankItem(Windows(GetWindowIndex("winBank")).Window.Left, Windows(GetWindowIndex("winBank")).Window.Top)

    If BankSlot > 0 Then
        ' exit out if we're offering that item

        ' drag it
        With DragBox
            .Type = Part_Item
            .value = Bank.Item(BankSlot).num
            .Origin = origin_Bank
            .Slot = BankSlot
        End With

        winIndex = GetWindowIndex("winDragBox")
        With Windows(winIndex).Window
            .state = MouseDown
            .Left = lastMouseX - 16
            .Top = lastMouseY - 16
            .movedX = clickedX - .Left
            .movedY = clickedY - .Top
        End With

        ShowWindow winIndex, , False
        ' stop dragging inventory
        Windows(GetWindowIndex("winBank")).Window.state = Normal
    End If

    ' show desc. if needed
    Bank_MouseMove
End Sub

' ###############
' ## Inventory ##
' ###############

Public Sub Inventory_MouseDown()
Dim invNum As Long, winIndex As Long, I As Long
    
    ' is there an item?
    invNum = IsItem(Windows(GetWindowIndex("winInventory")).Window.Left, Windows(GetWindowIndex("winInventory")).Window.Top)
    
    If invNum Then
        ' exit out if we're offering that item
        If InTrade > 0 Then
            For I = 1 To MAX_INV
                If TradeYourOffer(I).num = invNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).num)).Type = ITEM_TYPE_CURRENCY Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(I).value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(I).num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
            ' currency handler
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Then
                Dialogue "Select Amount", "Please choose how many to offer", "", TypeTRADEAMOUNT, StyleINPUT, invNum
                Exit Sub
            End If
            ' trade the normal item
            Call TradeItem(invNum, 0)
            Exit Sub
        End If
        
        ' drag it
        With DragBox
            .Type = Part_Item
            .value = GetPlayerInvItemNum(MyIndex, invNum)
            .Origin = origin_Inventory
            .Slot = invNum
        End With
        
        winIndex = GetWindowIndex("winDragBox")
        With Windows(winIndex).Window
            .state = MouseDown
            .Left = lastMouseX - 16
            .Top = lastMouseY - 16
            .movedX = clickedX - .Left
            .movedY = clickedY - .Top
        End With
        ShowWindow winIndex, , False
        ' stop dragging inventory
        Windows(GetWindowIndex("winInventory")).Window.state = Normal
    End If

    ' show desc. if needed
    Inventory_MouseMove
End Sub

Public Sub Inventory_DblClick()
Dim ItemNum As Long, I As Long

    If InTrade > 0 Then Exit Sub

    ItemNum = IsItem(Windows(GetWindowIndex("winInventory")).Window.Left, Windows(GetWindowIndex("winInventory")).Window.Top)
    
    If ItemNum Then
            SendUseItem ItemNum
    End If
    
    ' show desc. if needed
    Inventory_MouseMove
End Sub

Public Sub Inventory_MouseMove()
Dim ItemNum As Long, x As Long, y As Long, I As Long

    ' exit out early if dragging
    If DragBox.Type <> part_None Then Exit Sub

    ItemNum = IsItem(Windows(GetWindowIndex("winInventory")).Window.Left, Windows(GetWindowIndex("winInventory")).Window.Top)
    
    If ItemNum Then
        ' exit out if we're offering that item
        If InTrade > 0 Then
            For I = 1 To MAX_INV
                If TradeYourOffer(I).num = ItemNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).num)).Type = ITEM_TYPE_CURRENCY Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(I).value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(I).num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
        End If
        ' make sure we're not dragging the item
        If DragBox.Type = Part_Item And DragBox.value = ItemNum Then Exit Sub
        ' calc position
        x = Windows(GetWindowIndex("winInventory")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        y = Windows(GetWindowIndex("winInventory")).Window.Top - 4
        ' offscreen?
        If x < 0 Then
            ' switch to right
            x = Windows(GetWindowIndex("winInventory")).Window.Left + Windows(GetWindowIndex("winInventory")).Window.Width
        End If
        ' go go go
        ShowInvDesc x, y, ItemNum
    End If
End Sub

' ###############
' ## Character ##
' ###############

Public Sub Character_MouseDown()
Dim ItemNum As Long
    
    ItemNum = IsEqItem(Windows(GetWindowIndex("winCharacter")).Window.Left, Windows(GetWindowIndex("winCharacter")).Window.Top)
    
    If ItemNum Then
        SendUnequip ItemNum
    End If
    
    ' show desc. if needed
    Character_MouseMove
End Sub

Public Sub Character_MouseMove()
Dim ItemNum As Long, x As Long, y As Long

    ' exit out early if dragging
    If DragBox.Type <> part_None Then Exit Sub

    ItemNum = IsEqItem(Windows(GetWindowIndex("winCharacter")).Window.Left, Windows(GetWindowIndex("winCharacter")).Window.Top)
    
    If ItemNum Then
        ' calc position
        x = Windows(GetWindowIndex("winCharacter")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        y = Windows(GetWindowIndex("winCharacter")).Window.Top - 4
        ' offscreen?
        If x < 0 Then
            ' switch to right
            x = Windows(GetWindowIndex("winCharacter")).Window.Left + Windows(GetWindowIndex("winCharacter")).Window.Width
        End If
        ' go go go
        ShowEqDesc x, y, ItemNum
    End If
End Sub

Public Sub Character_SpendPoint1()
    SendTrainStat 1
End Sub

Public Sub Character_SpendPoint2()
    SendTrainStat 2
End Sub

Public Sub Character_SpendPoint3()
    SendTrainStat 3
End Sub

Public Sub Character_SpendPoint4()
    SendTrainStat 4
End Sub

Public Sub Character_SpendPoint5()
    SendTrainStat 5
End Sub

' #################
' ## Description ##
' #################

Public Sub Description_OnDraw()
Dim Xo As Long, Yo As Long, texNum As Long, y As Long, I As Long, count As Long

    ' exit out if we don't have a num
    If descItem = 0 Or descType = 0 Then Exit Sub

    Xo = Windows(GetWindowIndex("winDescription")).Window.Left
    Yo = Windows(GetWindowIndex("winDescription")).Window.Top
    
    Select Case descType
        Case 1 ' Inventory Item
            texNum = Tex_Item(Item(descItem).Pic)
        Case 2 ' Spell Icon
            texNum = Tex_Spellicon(Spell(descItem).icon)
            ' render bar
            With Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "picBar"))
                If .visible Then RenderTexture Tex_GUI(45), Xo + .Left, Yo + .Top, 0, 12, .value, 12, .value, 12
            End With
    End Select
    
    ' render sprite
    RenderTexture texNum, Xo + 20, Yo + 34, 0, 0, 64, 64, 32, 32
    
    ' render text array
    y = 18
    count = UBound(descText)
    For I = 1 To count
        RenderText font(Fonts.verdana_12), descText(I).text, Xo + 141 - (TextWidth(font(Fonts.verdana_12), descText(I).text) \ 2), Yo + y, descText(I).Colour
        y = y + 12
    Next
    
    ' close
    HideWindow GetWindowIndex("winDescription")
End Sub

' ##############
' ## Drag Box ##
' ##############

Public Sub DragBox_OnDraw()
Dim Xo As Long, Yo As Long, texNum As Long, winIndex As Long

    winIndex = GetWindowIndex("winDragBox")
    Xo = Windows(winIndex).Window.Left
    Yo = Windows(winIndex).Window.Top
    
    ' get texture num
    With DragBox
        Select Case .Type
            Case Part_Item
                If .value Then
                    texNum = Tex_Item(Item(.value).Pic)
                End If
            Case Part_spell
                If .value Then
                    texNum = Tex_Spellicon(Spell(.value).icon)
                End If
        End Select
    End With
    
    ' draw texture
    RenderTexture texNum, Xo, Yo, 0, 0, 32, 32, 32, 32
End Sub

Public Sub DragBox_Check()
Dim winIndex As Long, I As Long, curWindow As Long, curControl As Long, tmpRec As RECT
    
    winIndex = GetWindowIndex("winDragBox")
    
    ' can't drag nuthin'
    If DragBox.Type = part_None Then Exit Sub
    
    ' check for other windows
    For I = 1 To WindowCount
        With Windows(I).Window
            If .visible Then
                ' can't drag to self
                If .Name <> "winDragBox" Then
                    If currMouseX >= .Left And currMouseX <= .Left + .Width Then
                        If currMouseY >= .Top And currMouseY <= .Top + .Height Then
                            If curWindow = 0 Then curWindow = I
                            If .zOrder > Windows(curWindow).Window.zOrder Then curWindow = I
                        End If
                    End If
                End If
            End If
        End With
    Next
    
    ' we have a window - check if we can drop
    If curWindow Then
        Select Case Windows(curWindow).Window.Name
            Case "winBank"
                If DragBox.Origin = origin_Bank Then
                    ' it's from the inventory!
                    If DragBox.Type = Part_Item Then
                        ' find the slot to switch with
                        For I = 1 To MAX_BANK
                            With tmpRec
                                .Top = Windows(curWindow).Window.Top + BankTop + ((BankOffsetY + 32) * ((I - 1) \ BankColumns))
                                .bottom = .Top + 32
                                .Left = Windows(curWindow).Window.Left + BankLeft + ((BankOffsetX + 32) * (((I - 1) Mod BankColumns)))
                                .Right = .Left + 32
                            End With
    
                            If currMouseX >= tmpRec.Left And currMouseX <= tmpRec.Right Then
                                If currMouseY >= tmpRec.Top And currMouseY <= tmpRec.bottom Then
                                    ' switch the slots
                                    If DragBox.Slot <> I Then
                                        ChangeBankSlots DragBox.Slot, I
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
                
                ' se o item saiu do inventario
                If DragBox.Origin = origin_Inventory Then
                    If DragBox.Type = Part_Item Then
    
                        If Item(GetPlayerInvItemNum(MyIndex, DragBox.Slot)).Type <> ITEM_TYPE_CURRENCY Then
                            DepositItem DragBox.Slot, 1
                        Else
                            Dialogue "Depositar Item", "Insira a quantidade para depósito.", "", TypeDEPOSITITEM, StyleINPUT, DragBox.Slot
                        End If
    
                    End If
                End If
                
            Case "winInventory"
                If DragBox.Origin = origin_Inventory Then
                    ' it's from the inventory!
                    If DragBox.Type = Part_Item Then
                        ' find the slot to switch with
                        For I = 1 To MAX_INV
                            With tmpRec
                                .Top = Windows(curWindow).Window.Top + InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                                .bottom = .Top + 32
                                .Left = Windows(curWindow).Window.Left + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                                .Right = .Left + 32
                            End With
                            
                            If currMouseX >= tmpRec.Left And currMouseX <= tmpRec.Right Then
                                If currMouseY >= tmpRec.Top And currMouseY <= tmpRec.bottom Then
                                    ' switch the slots
                                    If DragBox.Slot <> I Then SendChangeInvSlots DragBox.Slot, I
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                End If
                
                ' se o item saiu do bank
                If DragBox.Origin = origin_Bank Then
                    If DragBox.Type = Part_Item Then
    
                        If Item(Bank.Item(DragBox.Slot).num).Type <> ITEM_TYPE_CURRENCY Then
                            WithdrawItem DragBox.Slot, 0
                        Else
                            Dialogue "Retirar Item", "Insira a quantidade que deseja retirar", "", TypeWITHDRAWITEM, StyleINPUT, DragBox.Slot
                        End If
    
                    End If
                End If
            Case "winSkills"
                If DragBox.Origin = origin_Spells Then
                    ' it's from the spells!
                    If DragBox.Type = Part_spell Then
                        ' find the slot to switch with
                        For I = 1 To MAX_PLAYER_SPELLS
                            With tmpRec
                                .Top = Windows(curWindow).Window.Top + SkillTop + ((SkillOffsetY + 32) * ((I - 1) \ SkillColumns))
                                .bottom = .Top + 32
                                .Left = Windows(curWindow).Window.Left + SkillLeft + ((SkillOffsetX + 32) * (((I - 1) Mod SkillColumns)))
                                .Right = .Left + 32
                            End With
                            
                            If currMouseX >= tmpRec.Left And currMouseX <= tmpRec.Right Then
                                If currMouseY >= tmpRec.Top And currMouseY <= tmpRec.bottom Then
                                    ' switch the slots
                                    If DragBox.Slot <> I Then SendChangeSpellSlots DragBox.Slot, I
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                End If
            Case "winHotbar"
                If DragBox.Origin <> origin_None Then
                    If DragBox.Type <> part_None Then
                        ' find the slot
                        For I = 1 To MAX_HOTBAR
                            With tmpRec
                                .Top = Windows(curWindow).Window.Top + HotbarTop
                                .bottom = .Top + 32
                                .Left = Windows(curWindow).Window.Left + HotbarLeft + ((I - 1) * HotbarOffsetX)
                                .Right = .Left + 32
                            End With
                            
                            If currMouseX >= tmpRec.Left And currMouseX <= tmpRec.Right Then
                                If currMouseY >= tmpRec.Top And currMouseY <= tmpRec.bottom Then
                                    ' set the hotbar slot
                                    If DragBox.Origin <> origin_Hotbar Then
                                        If DragBox.Type = Part_Item Then
                                            SendHotbarChange 1, DragBox.Slot, I
                                        ElseIf DragBox.Type = Part_spell Then
                                            SendHotbarChange 2, DragBox.Slot, I
                                        End If
                                    Else
                                        ' SWITCH the hotbar slots
                                        If DragBox.Slot <> I Then SwitchHotbar DragBox.Slot, I
                                    End If
                                    ' exit early
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                End If
        End Select
    Else
        ' no windows found - dropping on bare map
        Select Case DragBox.Origin
            Case PartTypeOrigins.origin_Inventory
                If Item(GetPlayerInvItemNum(MyIndex, DragBox.Slot)).Type <> ITEM_TYPE_CURRENCY Then
                    SendDropItem DragBox.Slot, GetPlayerInvItemNum(MyIndex, DragBox.Slot)
                Else
                    Dialogue "Drop Item", "Please choose how many to drop", "", TypeDROPITEM, StyleINPUT, GetPlayerInvItemNum(MyIndex, DragBox.Slot)
                End If
            Case PartTypeOrigins.origin_Spells
                ' dialogue
            Case PartTypeOrigins.origin_Hotbar
                SendHotbarChange 0, 0, DragBox.Slot
        End Select
    End If
    
    ' close window
    HideWindow winIndex
    With DragBox
        .Type = part_None
        .Slot = 0
        .Origin = origin_None
        .value = 0
    End With
End Sub

' ############
' ## Skills ##
' ############

Public Sub Skills_MouseDown()
Dim slotNum As Long, winIndex As Long
    
    ' is there an item?
    slotNum = IsSkill(Windows(GetWindowIndex("winSkills")).Window.Left, Windows(GetWindowIndex("winSkills")).Window.Top)
    
    If slotNum Then
        With DragBox
            .Type = Part_spell
            .value = PlayerSpells(slotNum).Spell
            .Origin = origin_Spells
            .Slot = slotNum
        End With
        
        winIndex = GetWindowIndex("winDragBox")
        With Windows(winIndex).Window
            .state = MouseDown
            .Left = lastMouseX - 16
            .Top = lastMouseY - 16
            .movedX = clickedX - .Left
            .movedY = clickedY - .Top
        End With
        ShowWindow winIndex, , False
        ' stop dragging inventory
        Windows(GetWindowIndex("winSkills")).Window.state = Normal
    End If

    ' show desc. if needed
    Skills_MouseMove
End Sub

Public Sub Skills_DblClick()
Dim slotNum As Long

    slotNum = IsSkill(Windows(GetWindowIndex("winSkills")).Window.Left, Windows(GetWindowIndex("winSkills")).Window.Top)
    
    If slotNum Then
        CastSpell slotNum
    End If
    
    ' show desc. if needed
    Skills_MouseMove
End Sub

Public Sub Skills_MouseMove()
Dim slotNum As Long, x As Long, y As Long

    ' exit out early if dragging
    If DragBox.Type <> part_None Then Exit Sub

    slotNum = IsSkill(Windows(GetWindowIndex("winSkills")).Window.Left, Windows(GetWindowIndex("winSkills")).Window.Top)
    
    If slotNum Then
        ' make sure we're not dragging the item
        If DragBox.Type = Part_Item And DragBox.value = slotNum Then Exit Sub
        ' calc position
        x = Windows(GetWindowIndex("winSkills")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        y = Windows(GetWindowIndex("winSkills")).Window.Top - 4
        ' offscreen?
        If x < 0 Then
            ' switch to right
            x = Windows(GetWindowIndex("winSkills")).Window.Left + Windows(GetWindowIndex("winSkills")).Window.Width
        End If
        ' go go go
        ShowPlayerSpellDesc x, y, slotNum
    End If
End Sub

' ############
' ## Hotbar ##
' ############

Public Sub Hotbar_MouseDown()
Dim slotNum As Long, winIndex As Long
    
    ' is there an item?
    slotNum = IsHotbar(Windows(GetWindowIndex("winHotbar")).Window.Left, Windows(GetWindowIndex("winHotbar")).Window.Top)
    
    If slotNum Then
        With DragBox
            If Hotbar(slotNum).sType = 1 Then ' inventory
                .Type = Part_Item
            ElseIf Hotbar(slotNum).sType = 2 Then ' spell
                .Type = Part_spell
            End If
            .value = Hotbar(slotNum).Slot
            .Origin = origin_Hotbar
            .Slot = slotNum
        End With
        
        winIndex = GetWindowIndex("winDragBox")
        With Windows(winIndex).Window
            .state = MouseDown
            .Left = lastMouseX - 16
            .Top = lastMouseY - 16
            .movedX = clickedX - .Left
            .movedY = clickedY - .Top
        End With
        ShowWindow winIndex, , False
        ' stop dragging inventory
        Windows(GetWindowIndex("winHotbar")).Window.state = Normal
    End If

    ' show desc. if needed
    Hotbar_MouseMove
End Sub

Public Sub Hotbar_DblClick()
Dim slotNum As Long

    slotNum = IsHotbar(Windows(GetWindowIndex("winHotbar")).Window.Left, Windows(GetWindowIndex("winHotbar")).Window.Top)
    
    If slotNum Then
        SendHotbarUse slotNum
    End If
    
    ' show desc. if needed
    Hotbar_MouseMove
End Sub

Public Sub Hotbar_MouseMove()
Dim slotNum As Long, x As Long, y As Long

    ' exit out early if dragging
    If DragBox.Type <> part_None Then Exit Sub

    slotNum = IsHotbar(Windows(GetWindowIndex("winHotbar")).Window.Left, Windows(GetWindowIndex("winHotbar")).Window.Top)
    
    If slotNum Then
        ' make sure we're not dragging the item
        If DragBox.Origin = origin_Hotbar And DragBox.Slot = slotNum Then Exit Sub
        ' calc position
        x = Windows(GetWindowIndex("winHotbar")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        y = Windows(GetWindowIndex("winHotbar")).Window.Top - 4
        ' offscreen?
        If x < 0 Then
            ' switch to right
            x = Windows(GetWindowIndex("winHotbar")).Window.Left + Windows(GetWindowIndex("winHotbar")).Window.Width
        End If
        ' go go go
        Select Case Hotbar(slotNum).sType
            Case 1 ' inventory
                ShowItemDesc x, y, Hotbar(slotNum).Slot, False
            Case 2 ' spells
                ShowSpellDesc x, y, Hotbar(slotNum).Slot, 0
        End Select
    End If
End Sub

' Chat
Public Sub btnSay_Click()
    HandleKeyPresses vbKeyReturn
End Sub

Public Sub OnDraw_Chat()
Dim winIndex As Long, Xo As Long, Yo As Long

    winIndex = GetWindowIndex("winChat")
    Xo = Windows(winIndex).Window.Left
    Yo = Windows(winIndex).Window.Top + 16
    
    ' draw the box
    RenderDesign DesignTypes.desWin_Desc, Xo, Yo, 352, 152
    ' draw the input box
    RenderTexture Tex_GUI(46), Xo + 7, Yo + 123, 0, 0, 171, 22, 171, 22
    RenderTexture Tex_GUI(46), Xo + 174, Yo + 123, 0, 22, 171, 22, 171, 22
    ' call the chat render
    RenderChat
End Sub

Public Sub OnDraw_ChatSmall()
Dim winIndex As Long, Xo As Long, Yo As Long

    winIndex = GetWindowIndex("winChatSmall")
    
    If actChatWidth < 160 Then actChatWidth = 160
    If actChatHeight < 10 Then actChatHeight = 10
    
    Xo = Windows(winIndex).Window.Left + 10
    Yo = ScreenHeight - 16 - actChatHeight - 8
    
    ' draw the background
    RenderDesign DesignTypes.desWin_Shadow, Xo, Yo, actChatWidth, actChatHeight
    ' call the chat render
    RenderChat
End Sub

Public Sub chkChat_Game()
    Options.channelState(ChatChannel.chGame) = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "chkGame")).value
    UpdateChat
End Sub

Public Sub chkChat_Map()
    Options.channelState(ChatChannel.chMap) = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "chkMap")).value
    UpdateChat
End Sub

Public Sub chkChat_Global()
    Options.channelState(ChatChannel.chGlobal) = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "chkGlobal")).value
    UpdateChat
End Sub

Public Sub chkChat_Party()
    Options.channelState(ChatChannel.chParty) = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "chkParty")).value
    UpdateChat
End Sub

Public Sub chkChat_Guild()
    Options.channelState(ChatChannel.chGuild) = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "chkGuild")).value
    UpdateChat
End Sub

Public Sub chkChat_Private()
    Options.channelState(ChatChannel.chPrivate) = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "chkPrivate")).value
    UpdateChat
End Sub

Public Sub btnChat_Up()
    ChatButtonUp = True
End Sub

Public Sub btnChat_Down()
    ChatButtonDown = True
End Sub

Public Sub btnChat_Up_MouseUp()
    ChatButtonUp = False
End Sub

Public Sub btnChat_Down_MouseUp()
    ChatButtonDown = False
End Sub

' Options
Public Sub btnOptions_Close()
    HideWindow GetWindowIndex("winOptions")
    ShowWindow GetWindowIndex("winEscMenu")
End Sub

Sub btnOptions_Confirm()
Dim I As Long, value As Long, Width As Long, Height As Long, message As Boolean, musicFile As String

    ' music
    value = Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "chkMusic")).value
    If Options.Music <> value Then
        Options.Music = value
        ' let them know
        If value = 0 Then
            AddText "Music turned off.", BrightGreen
            Stop_Music
        Else
            AddText "Music tured on.", BrightGreen
            ' play music
            If InGame Then musicFile = Trim$(Map.MapData.Music) Else musicFile = Trim$(MenuMusic)
            If Not musicFile = "None." Then
                Play_Music musicFile
            Else
                Stop_Music
            End If
        End If
    End If
    
    ' sound
    value = Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "chkSound")).value
    If Options.sound <> value Then
        Options.sound = value
        ' let them know
        If value = 0 Then
            AddText "Sound turned off.", BrightGreen
        Else
            AddText "Sound tured on.", BrightGreen
        End If
    End If
    
    ' autotiles
    value = Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "chkAutotiles")).value
    If value = 1 Then value = 0 Else value = 1
    If Options.NoAuto <> value Then
        Options.NoAuto = value
        ' let them know
        If value = 0 Then
            If InGame Then
                AddText "Autotiles turned on.", BrightGreen
                initAutotiles
            End If
        Else
            If InGame Then
                AddText "Autotiles turned off.", BrightGreen
                initAutotiles
            End If
        End If
    End If
    
    ' fullscreen
    value = Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "chkFullscreen")).value
    If Options.Fullscreen <> value Then
        Options.Fullscreen = value
        message = True
    End If
    
    ' resolution
    With Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRes"))
        If .value > 0 And .value <= RES_COUNT Then
            If Options.Resolution <> .value Then
                Options.Resolution = .value
                If Not isFullscreen Then
                    SetResolution
                Else
                    message = True
                End If
            End If
        End If
    End With
    
    ' render
    With Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRender"))
        If .value > 0 And .value <= 3 Then
            If Options.Render <> .value - 1 Then
                Options.Render = .value - 1
                message = True
            End If
        End If
    End With
    
    ' save options
    SaveOptions
    ' let them know
    If InGame Then
        If message Then AddText "Some changes will take effect next time you load the game.", BrightGreen
    End If
    ' close
    btnOptions_Close
End Sub

' OfferWindow
Public Sub AcceptOffer1()

End Sub

Public Sub RecuseOffer1()
    Call UpdateOffers(1)
End Sub

Public Sub AcceptOffer2()

End Sub

Public Sub RecuseOffer2()
    Call UpdateOffers(2)
End Sub

Public Sub AcceptOffer3()

End Sub

Public Sub RecuseOffer3()
    Call UpdateOffers(3)
End Sub

' Npc Chat
Public Sub btnNpcChat_Close()
    HideWindow GetWindowIndex("winNpcChat")
End Sub

Public Sub btnOpt1()
    SendChatOption 1
End Sub
Public Sub btnOpt2()
    SendChatOption 2
End Sub
Public Sub btnOpt3()
    SendChatOption 3
End Sub
Public Sub btnOpt4()
    SendChatOption 4
End Sub

' Shop
Public Sub btnShop_Close()
    CloseShop
End Sub

Public Sub chkShopBuying()
    With Windows(GetWindowIndex("winShop"))
        If .Controls(GetControlIndex("winShop", "chkBuying")).value = 1 Then
            .Controls(GetControlIndex("winShop", "chkSelling")).value = 0
        Else
            .Controls(GetControlIndex("winShop", "chkSelling")).value = 0
            .Controls(GetControlIndex("winShop", "chkBuying")).value = 1
            Exit Sub
        End If
    End With
    ' show buy button, hide sell
    With Windows(GetWindowIndex("winShop"))
        .Controls(GetControlIndex("winShop", "btnSell")).visible = False
        .Controls(GetControlIndex("winShop", "btnBuy")).visible = True
    End With
    ' update the shop
    shopIsSelling = False
    shopSelectedSlot = 1
    UpdateShop
End Sub

Public Sub chkShopSelling()
    With Windows(GetWindowIndex("winShop"))
        If .Controls(GetControlIndex("winShop", "chkSelling")).value = 1 Then
            .Controls(GetControlIndex("winShop", "chkBuying")).value = 0
        Else
            .Controls(GetControlIndex("winShop", "chkBuying")).value = 0
            .Controls(GetControlIndex("winShop", "chkSelling")).value = 1
            Exit Sub
        End If
    End With
    ' show sell button, hide buy
    With Windows(GetWindowIndex("winShop"))
        .Controls(GetControlIndex("winShop", "btnBuy")).visible = False
        .Controls(GetControlIndex("winShop", "btnSell")).visible = True
    End With
    ' update the shop
    shopIsSelling = True
    shopSelectedSlot = 1
    UpdateShop
End Sub

Public Sub btnShopBuy()
    BuyItem shopSelectedSlot
End Sub

Public Sub btnShopSell()
    SellItem shopSelectedSlot
End Sub

Public Sub Shop_MouseDown()
Dim shopNum As Long
    
    ' is there an item?
    shopNum = IsShopSlot(Windows(GetWindowIndex("winShop")).Window.Left, Windows(GetWindowIndex("winShop")).Window.Top)
    
    If shopNum Then
        ' set the active slot
        shopSelectedSlot = shopNum
        UpdateShop
    End If
    
    Shop_MouseMove
End Sub

Public Sub Shop_MouseMove()
Dim shopSlot As Long, ItemNum As Long, x As Long, y As Long

    If InShop = 0 Then Exit Sub

    shopSlot = IsShopSlot(Windows(GetWindowIndex("winShop")).Window.Left, Windows(GetWindowIndex("winShop")).Window.Top)
    
    If shopSlot Then
        ' calc position
        x = Windows(GetWindowIndex("winShop")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        y = Windows(GetWindowIndex("winShop")).Window.Top - 4
        ' offscreen?
        If x < 0 Then
            ' switch to right
            x = Windows(GetWindowIndex("winShop")).Window.Left + Windows(GetWindowIndex("winShop")).Window.Width
        End If
        ' selling/buying
        If Not shopIsSelling Then
            ' get the itemnum
            ItemNum = Shop(InShop).TradeItem(shopSlot).Item
            If ItemNum = 0 Then Exit Sub
            ShowShopDesc x, y, ItemNum
        Else
            ' get the itemnum
            ItemNum = GetPlayerInvItemNum(MyIndex, shopSlot)
            If ItemNum = 0 Then Exit Sub
            ShowShopDesc x, y, ItemNum
        End If
    End If
End Sub

' Right Click Menu
Sub RightClick_Close()
    ' close all menus
    HideWindow GetWindowIndex("winRightClickBG")
    HideWindow GetWindowIndex("winPlayerMenu")
End Sub

' Player Menu
Sub PlayerMenu_Party()
    RightClick_Close
    If PlayerMenuIndex = 0 Then Exit Sub
    SendPartyRequest PlayerMenuIndex
End Sub

Sub PlayerMenu_Trade()
    RightClick_Close
    If PlayerMenuIndex = 0 Then Exit Sub
    SendTradeRequest PlayerMenuIndex
End Sub

Sub PlayerMenu_Guild()
    RightClick_Close
    If PlayerMenuIndex = 0 Then Exit Sub
    AddText "System not yet in place.", BrightRed
End Sub

Sub PlayerMenu_PM()
    RightClick_Close
    If PlayerMenuIndex = 0 Then Exit Sub
    AddText "System not yet in place.", BrightRed
End Sub

' Trade
Sub btnTrade_Close()
    HideWindow GetWindowIndex("winTrade")
    DeclineTrade
End Sub

Sub btnTrade_Accept()
    AcceptTrade
End Sub

Sub TradeMouseDown_Your()
Dim Xo As Long, Yo As Long, ItemNum As Long
    Xo = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).Left
    Yo = Windows(GetWindowIndex("winTrade")).Window.Top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).Top
    ItemNum = IsTrade(Xo, Yo)
    
    ' make sure it exists
    If ItemNum > 0 Then
        If TradeYourOffer(ItemNum).num = 0 Then Exit Sub
        If GetPlayerInvItemNum(MyIndex, TradeYourOffer(ItemNum).num) = 0 Then Exit Sub
        
        ' unoffer the item
        UntradeItem ItemNum
    End If
End Sub

Sub TradeMouseMove_Your()
Dim Xo As Long, Yo As Long, ItemNum As Long, x As Long, y As Long
    Xo = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).Left
    Yo = Windows(GetWindowIndex("winTrade")).Window.Top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).Top
    ItemNum = IsTrade(Xo, Yo)
    
    ' make sure it exists
    If ItemNum > 0 Then
        If TradeYourOffer(ItemNum).num = 0 Then Exit Sub
        If GetPlayerInvItemNum(MyIndex, TradeYourOffer(ItemNum).num) = 0 Then Exit Sub
        
        ' calc position
        x = Windows(GetWindowIndex("winTrade")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        y = Windows(GetWindowIndex("winTrade")).Window.Top - 4
        ' offscreen?
        If x < 0 Then
            ' switch to right
            x = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Window.Width
        End If
        ' go go go
        ShowItemDesc x, y, GetPlayerInvItemNum(MyIndex, TradeYourOffer(ItemNum).num), False
    End If
End Sub

Sub TradeMouseMove_Their()
Dim Xo As Long, Yo As Long, ItemNum As Long, x As Long, y As Long
    Xo = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picTheir")).Left
    Yo = Windows(GetWindowIndex("winTrade")).Window.Top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picTheir")).Top
    ItemNum = IsTrade(Xo, Yo)
    
    ' make sure it exists
    If ItemNum > 0 Then
        If TradeTheirOffer(ItemNum).num = 0 Then Exit Sub
        
        ' calc position
        x = Windows(GetWindowIndex("winTrade")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        y = Windows(GetWindowIndex("winTrade")).Window.Top - 4
        ' offscreen?
        If x < 0 Then
            ' switch to right
            x = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Window.Width
        End If
        ' go go go
        ShowItemDesc x, y, TradeTheirOffer(ItemNum).num, False
    End If
End Sub

' combobox
Sub CloseComboMenu()
    HideWindow GetWindowIndex("winComboMenuBG")
    HideWindow GetWindowIndex("winComboMenu")
End Sub
