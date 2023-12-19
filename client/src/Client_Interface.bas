Attribute VB_Name = "Client_Interface"
Option Explicit

' Entity Types
Public Enum EntityTypes
    EntityLabel = 1
    EntityWindow
    EntityButton
    EntityTextBox
    EntityPictureBox
    EntityCheckbox
    entityCombo
End Enum

' Design Types
Public Enum DesignTypes

    ' Boxes
    DesignWoodNormal = 1
    DesignWoodEmpty

    DesignGreenNormal
    DesignGreenHover
    DesignGreenClick

    DesignRedNormal
    DesignRedHover
    DesignRedClick

    DesignBlueNormal
    DesignBlueHover
    DesignBlueClick

    DesignGoldNormal
    DesignGoldHover
    DesignGoldClick

    DesignGrey

    ' Windows
    DesignWindowNormal
    DesignWindowWithoutBar
    DesignWindowClear
    designWindowDescription
    designWindowShadow

    DesignParchment
    DesignBlackParchment

    ' Textboxes
    DesignTextInput

    ' Checkboxes
    DesignCheckbox
    DesignCheckChat
    DesignCheckBuy
    DesignCheckSell

    ' Right-click Menu
    DesignMenuHeader
    DesignMenuHover

    ' Color
    DesignColor

    ' Comboboxes
    DesignCombo
    DesignComboBackground

    ' tile Selection
    designTilesetGrid
End Enum

' Button States
Public Enum EntityStates
    Normal = 0
    Hover
    MouseDown
    MouseMove
    MouseUp
    DoubleClick
    Enter

    ' Count
    StateCount
End Enum

' Alignment
Public Enum Alignment
    AlignLeft = 0
    AlignRight
    AlignCentre
End Enum

' Part Types
Public Enum PartType
    PartNone = 0
    PartItem
    Partspell
End Enum

' Origins
Public Enum PartTypeOrigins
    OriginNone = 0
    OriginInventory
    OriginHotbar
    OriginSpells
    OriginBank
End Enum

' Entity UDT
Public Type EntityRec
    ' constants
    Name As String
    ' values
Type As Byte
    Top As Long
    Left As Long
    Width As Long
    Height As Long
    enabled As Boolean
    visible As Boolean
    canDrag As Boolean
    max As Long
    min As Long
    Value As Long
    text As String
    image(0 To EntityStates.StateCount - 1) As Long
    design(0 To EntityStates.StateCount - 1) As Long
    entCallBack(0 To EntityStates.StateCount - 1) As Long
    alpha As Long
    clickThrough As Boolean
    xOffset As Long
    yOffset As Long
    align As Byte
    font As Long
    textColour As Long
    textColourHover As Long
    textColourClick As Long
    zChange As Byte
    onDraw As Long
    origLeft As Long
    origTop As Long
    tooltip As String
    group As Long
    list() As String
    activated As Boolean
    linkedToWin As Long
    linkedToCon As Long
    ' window
    icon As Long
    ' textbox
    isCensor As Boolean
    ' temp
    state As EntityStates
    movedX As Long
    movedY As Long
    zOrder As Long
End Type

' For small parts
Public Type EntityPartRec
Type As PartType
    Origin As PartTypeOrigins
    Value As Long
    Slot As Long
End Type

' Window UDT
Public Type WindowRec
    Window As EntityRec
    Controls() As EntityRec
    ControlCount As Long
    activeControl As Long
End Type

' actual GUI
Public Windows() As WindowRec
Public windowCount As Long

Public windowUpdated As Boolean
Public controlUpdated As Boolean

Public activeWindow As Long

' GUI parts
Public DragBox As EntityPartRec
Private zOrder_Win As Long
Private zOrder_Con As Long

Public Function SetzOrder_Win(ByVal Value As Long)
    zOrder_Win = Value
End Function

Public Sub CreateEntity(winNum As Long, zOrder As Long, Name As String, tType As EntityTypes, ByRef design() As Long, ByRef image() As Long, ByRef entCallBack() As Long, _
                        Optional Left As Long, Optional Top As Long, Optional Width As Long, Optional Height As Long, Optional visible As Boolean = True, Optional canDrag As Boolean, Optional max As Long, _
                        Optional min As Long, Optional Value As Long, Optional text As String, Optional align As Byte, Optional font As Long = Fonts.georgia_16, Optional textColour As Long = White, _
                        Optional alpha As Long = 255, Optional clickThrough As Boolean, Optional xOffset As Long, Optional yOffset As Long, Optional zChange As Byte, Optional ByVal icon As Long, _
                        Optional ByVal onDraw As Long, Optional isActive As Boolean, Optional isCensor As Boolean, Optional textColourHover As Long, Optional textColourClick As Long, _
                        Optional tooltip As String, Optional group As Long)
    Dim i As Long

    ' check if it's a legal number
    If winNum <= 0 Or winNum > windowCount Then
        Exit Sub
    End If

    If windowUpdated Then
        If controlUpdated Then
            Windows(winNum).ControlCount = 0
            controlUpdated = False
        End If
    End If
    
    ' re-dim the control array
    With Windows(winNum)
        .ControlCount = .ControlCount + 1
        ReDim Preserve .Controls(1 To .ControlCount) As EntityRec
    End With

    ' Set the new control values
    With Windows(winNum).Controls(Windows(winNum).ControlCount)
        .Name = Name
        .Type = tType

        ' loop through states
        For i = 0 To EntityStates.StateCount - 1
            .design(i) = design(i)
            .image(i) = image(i)
            .entCallBack(i) = entCallBack(i)
        Next

        .Left = Left
        .Top = Top
        .origLeft = Left
        .origTop = Top
        .Width = Width
        .Height = Height
        .visible = visible
        .canDrag = canDrag
        .max = max
        .min = min
        .Value = Value
        .text = text
        .align = align
        .font = font
        .textColour = textColour
        .textColourHover = textColourHover
        .textColourClick = textColourClick
        .alpha = alpha
        .clickThrough = clickThrough
        .xOffset = xOffset
        .yOffset = yOffset
        .zChange = zChange
        .zOrder = zOrder
        .enabled = True
        .icon = icon
        .onDraw = onDraw
        .isCensor = isCensor
        .tooltip = tooltip
        .group = group
        ReDim .list(0 To 0) As String
    End With

    ' set the active control
    If isActive Then Windows(winNum).activeControl = Windows(winNum).ControlCount

    ' set the zOrder
    zOrder_Con = zOrder_Con + 1
End Sub

Public Sub UpdateZOrder(winNum As Long, Optional forced As Boolean = False)
    Dim i As Long
    Dim oldZOrder As Long

    With Windows(winNum).Window

        If Not forced Then If .zChange = 0 Then Exit Sub
        If .zOrder = windowCount Then Exit Sub
        oldZOrder = .zOrder

        For i = 1 To windowCount

            If Windows(i).Window.zOrder > oldZOrder Then
                Windows(i).Window.zOrder = Windows(i).Window.zOrder - 1
            End If

        Next

        .zOrder = windowCount
    End With

End Sub

Public Sub SortWindows()
    Dim tempWindow As WindowRec
    Dim i As Long, X As Long
    X = 1

    While X <> 0
        X = 0

        For i = 1 To windowCount - 1

            If Windows(i).Window.zOrder > Windows(i + 1).Window.zOrder Then
                tempWindow = Windows(i)
                Windows(i) = Windows(i + 1)
                Windows(i + 1) = tempWindow
                X = 1
            End If

        Next

    Wend

End Sub

Public Sub RenderEntities()
    Dim i As Long, X As Long, curZOrder As Long

    ' don't render anything if we don't have any containers
    If windowCount = 0 Then Exit Sub
    ' reset zOrder
    curZOrder = 1

    ' loop through windows
    Do While curZOrder <= windowCount
        For i = 1 To windowCount
            If curZOrder = Windows(i).Window.zOrder Then
                ' increment
                curZOrder = curZOrder + 1
                ' make sure it's visible
                If Windows(i).Window.visible Then
                    ' render container
                    RenderWindow i
                    ' render controls
                    For X = 1 To Windows(i).ControlCount
                        If Windows(i).Controls(X).visible Then RenderEntity i, X
                    Next
                End If
            End If
        Next
    Loop
End Sub

Public Sub RenderEntity(winNum As Long, entNum As Long)
    Dim Xo As Long, Yo As Long, hor_centre As Long, ver_centre As Long, Height As Long, Width As Long, Left As Long, texNum As Long, xOffset As Long
    Dim Callback As Long, taddText As String, Colour As Long, textArray() As String, Count As Long, yOffset As Long, i As Long, Y As Long, X As Long

    ' check if the window exists
    If winNum <= 0 Or winNum > windowCount Then
        Exit Sub
    End If

    ' check if the entity exists
    If entNum <= 0 Or entNum > Windows(winNum).ControlCount Then
        Exit Sub
    End If

    ' check the container's position
    Xo = Windows(winNum).Window.Left
    Yo = Windows(winNum).Window.Top

    With Windows(winNum).Controls(entNum)

        ' find the control type
        Select Case .Type
            ' picture box

        Case EntityTypes.EntityPictureBox

            ' render specific designs
            If .design(.state) > 0 Then RenderDesign .design(.state), .Left + Xo, .Top + Yo, .Width, .Height, .alpha
            ' render image
            If .image(.state) > 0 Then RenderTexture .image(.state), .Left + Xo, .Top + Yo, 0, 0, .Width, .Height, .Width, .Height, DX8Colour(White, .alpha)

            ' textbox
        Case EntityTypes.EntityTextBox
            ' render specific designs

            If .design(.state) > 0 Then RenderDesign .design(.state), .Left + Xo, .Top + Yo, .Width, .Height, .alpha
            ' render image
            If .image(.state) > 0 Then RenderTexture .image(.state), .Left + Xo, .Top + Yo, 0, 0, .Width, .Height, .Width, .Height, DX8Colour(White, .alpha)
            ' render text
            If activeWindow = winNum And Windows(winNum).activeControl = entNum Then taddText = chatShowLine
            ' if it's censored then render censored
            If Not .isCensor Then
                RenderText font(.font), .text & taddText, .Left + Xo + .xOffset, .Top + Yo + .yOffset, .textColour
            Else
                RenderText font(.font), CensorWord(.text) & taddText, .Left + Xo + .xOffset, .Top + Yo + .yOffset, .textColour
            End If

            ' buttons
        Case EntityTypes.EntityButton

            ' render specific designs
            If .design(.state) > 0 Then
                If .design(.state) > 0 Then
                    RenderDesign .design(.state), .Left + Xo, .Top + Yo, .Width, .Height
                End If
            End If
            ' render image
            If .image(.state) > 0 Then
                If .image(.state) > 0 Then
                    RenderTexture .image(.state), .Left + Xo, .Top + Yo, 0, 0, .Width, .Height, .Width, .Height
                End If
            End If
            ' render icon
            If .icon > 0 Then
                Width = mTexture(.icon).Width
                Height = mTexture(.icon).Height
                RenderTexture .icon, .Left + Xo + .xOffset, .Top + Yo + .yOffset, 0, 0, Width, Height, Width, Height
            End If
            ' for changing the text space
            xOffset = Width
            ' calculate the vertical centre
            Height = TextHeight(font(Fonts.georgiaDec_16))
            If Height > .Height Then
                ver_centre = .Top + Yo
            Else
                ver_centre = .Top + Yo + ((.Height - Height) \ 2) + 1
            End If
            ' calculate the horizontal centre
            Width = TextWidth(font(.font), .text)
            If Width > .Width Then
                hor_centre = .Left + Xo + xOffset
            Else
                hor_centre = .Left + Xo + xOffset + ((.Width - Width - xOffset) \ 2)
            End If
            ' get the colour
            If .state = Hover Then
                Colour = .textColourHover
            ElseIf .state = MouseDown Then
                Colour = .textColourClick

            Else
                Colour = .textColour
            End If
            RenderText font(.font), .text, hor_centre, ver_centre, Colour

            ' labels
        Case EntityTypes.EntityLabel

            If Len(.text) > 0 Then
                Select Case .align
                Case Alignment.AlignLeft
                    ' check if need to word wrap
                    If TextWidth(font(.font), .text) > .Width Then
                        ' wrap text
                        WordWrap_Array .text, .Width, textArray()
                        ' render text
                        Count = UBound(textArray)
                        For i = 1 To Count
                            RenderText font(.font), textArray(i), .Left + Xo, .Top + Yo + yOffset, .textColour, .alpha
                            yOffset = yOffset + 14
                        Next
                    Else
                        ' just one line
                        RenderText font(.font), .text, .Left + Xo, .Top + Yo, .textColour, .alpha
                    End If
                Case Alignment.AlignRight
                    ' check if need to word wrap
                    If TextWidth(font(.font), .text) > .Width Then
                        ' wrap text
                        WordWrap_Array .text, .Width, textArray()
                        ' render text
                        Count = UBound(textArray)
                        For i = 1 To Count
                            Left = .Left + .Width - TextWidth(font(.font), textArray(i))
                            RenderText font(.font), textArray(i), Left + Xo, .Top + Yo + yOffset, .textColour, .alpha
                            yOffset = yOffset + 14
                        Next
                    Else
                        ' just one line
                        Left = .Left + .Width - TextWidth(font(.font), .text)
                        RenderText font(.font), .text, Left + Xo, .Top + Yo, .textColour, .alpha
                    End If
                Case Alignment.AlignCentre
                    ' check if need to word wrap
                    If TextWidth(font(.font), .text) > .Width Then
                        ' wrap text
                        WordWrap_Array .text, .Width, textArray()
                        ' render text
                        Count = UBound(textArray)
                        For i = 1 To Count
                            Left = .Left + (.Width \ 2) - (TextWidth(font(.font), textArray(i)) \ 2)
                            RenderText font(.font), textArray(i), Left + Xo, .Top + Yo + yOffset, .textColour, .alpha
                            yOffset = yOffset + 14
                        Next
                    Else
                        ' just one line
                        Left = .Left + (.Width \ 2) - (TextWidth(font(.font), .text) \ 2)
                        RenderText font(.font), .text, Left + Xo, .Top + Yo, .textColour, .alpha
                    End If
                End Select
            End If

            ' checkboxes
        Case EntityTypes.EntityCheckbox

            Select Case .design(0)
            Case DesignTypes.DesignCheckbox

                ' empty?
                If .Value = 0 Then texNum = TextureGUI(32) Else texNum = TextureGUI(33)
                ' render box
                RenderTexture texNum, .Left + Xo, .Top + Yo, 0, 0, 14, 14, 14, 14
                ' find text position
                Select Case .align
                Case Alignment.AlignLeft
                    Left = .Left + 18 + Xo
                Case Alignment.AlignRight
                    Left = .Left + 18 + (.Width - 18) - TextWidth(font(.font), .text) + Xo
                Case Alignment.AlignCentre
                    Left = .Left + 18 + ((.Width - 18) / 2) - (TextWidth(font(.font), .text) / 2) + Xo
                End Select
                ' render text
                RenderText font(.font), .text, Left, .Top + Yo, .textColour, .alpha
            Case DesignTypes.DesignCheckChat

                If .Value = 0 Then .alpha = 150 Else .alpha = 255

                ' render box
                RenderEntity_Square TextureDesign(1), .Left + Xo, .Top + Yo, 49, 23, 4, .alpha

                '
                Left = .Left + (49 / 2) - (TextWidth(font(.font), .text) / 2) + Xo
                ' render text

                RenderText font(.font), .text, Left, .Top + Yo + 4, .textColour, .alpha

            Case DesignTypes.DesignCheckBuy

                If .Value = 0 Then texNum = TextureGradient(1) Else texNum = TextureGradient(2)
                RenderEntity_Square TextureDesign(9), .Left + Xo, .Top + Yo, 49, 20, 2, 255

                RenderTexture texNum, .Left + Xo + 2, .Top + Yo + 2, 0, 0, 45, 16, 45, 16

                Left = .Left + (49 / 2) - (TextWidth(font(.font), .text) / 2) + Xo
                RenderText font(.font), .text, Left, .Top + Yo + 4, .textColour, .alpha
            Case DesignTypes.DesignCheckSell

                If .Value = 0 Then texNum = TextureGradient(4) Else texNum = TextureGradient(5)
                RenderEntity_Square TextureDesign(10), .Left + Xo, .Top + Yo, 49, 20, 2, 255

                RenderTexture texNum, .Left + Xo + 2, .Top + Yo + 2, 0, 0, 45, 16, 45, 16

                Left = .Left + (49 / 2) - (TextWidth(font(.font), .text) / 2) + Xo
                RenderText font(.font), .text, Left, .Top + Yo + 4, .textColour, .alpha
            End Select

            ' comboboxes
        Case EntityTypes.entityCombo
            Select Case .design(0)
            Case DesignTypes.DesignCombo
                ' draw the background
                RenderDesign DesignTypes.DesignBlackParchment, .Left + Xo, .Top + Yo, .Width, .Height
                ' render the text
                If .Value > 0 Then
                    If .Value <= UBound(.list) Then
                        RenderText font(.font), .list(.Value), .Left + Xo + 5, .Top + Yo + 3, White
                    End If
                End If
                ' draw the little arow
                RenderTexture TextureGUI(5), .Left + Xo + .Width - 11, .Top + Yo + 7, 0, 0, 5, 4, 5, 4
            End Select
        End Select

        ' callback draw
        Callback = .onDraw

        If Callback <> 0 Then entCallBack Callback, winNum, entNum, 0, 0
    End With

End Sub

Public Sub RenderWindow(winNum As Long)
    Dim Width As Long, Height As Long, Callback As Long, X As Long, Y As Long, i As Long, Left As Long

    ' check if the window exists
    If winNum <= 0 Or winNum > windowCount Then
        Exit Sub
    End If

    With Windows(winNum).Window

        Select Case .design(0)
        Case DesignTypes.DesignComboBackground


            RenderDesign DesignTypes.DesignBlackParchment, .Left, .Top + 2, .Width, .Height

            ' text
            If UBound(.list) > 0 Then
                Y = .Top + 2

                X = .Left
                For i = 1 To UBound(.list)
                    ' render select
                    If i = .Value Or i = .group Then RenderDesign DesignTypes.DesignBlackParchment, X, Y - 1, .Width, 15
                    ' render text
                    Left = X + (.Width \ 2) - (TextWidth(font(.font), .list(i)) \ 2)
                    If i = .Value Or i = .group Then
                        RenderText font(.font), .list(i), Left, Y, Yellow
                    Else
                        RenderText font(.font), .list(i), Left, Y, White
                    End If
                    Y = Y + 16
                Next
            End If
            Exit Sub
        End Select

        Select Case .design(.state)

        Case DesignTypes.DesignWindowNormal

            ' Render do background da janela
            RenderDesign DesignTypes.DesignWoodNormal, .Left, .Top, .Width, .Height

            ' Render da top bar da janela
            RenderDesign DesignTypes.DesignGreenNormal, .Left, .Top, .Width, 40

            ' render the caption
            RenderText font(.font), Trim$(.text), .Left + Height + 20, .Top + 15, .textColour

        Case DesignTypes.DesignWindowWithoutBar
            ' render window
            RenderDesign DesignTypes.DesignWoodNormal, .Left, .Top, .Width, .Height

        Case DesignTypes.DesignWindowClear
            ' render window
            RenderDesign DesignTypes.DesignWoodEmpty, .Left, .Top, .Width, .Height
            RenderDesign DesignTypes.DesignGreenNormal, .Left, .Top, .Width, 40

            ' render the caption
            RenderText font(.font), Trim$(.text), .Left + Height + 20, .Top + 15, .textColour

        Case DesignTypes.designWindowDescription
            ' render window
            RenderDesign DesignTypes.designWindowDescription, .Left, .Top, .Width, .Height

        Case designWindowShadow
            ' render window
            RenderDesign DesignTypes.designWindowShadow, .Left, .Top, .Width, .Height

        End Select

        ' OnDraw call back
        Callback = .onDraw

        If Callback <> 0 Then entCallBack Callback, winNum, 0, 0, 0
    End With

End Sub

Public Sub RenderDesign(design As Long, Left As Long, Top As Long, Width As Long, Height As Long, Optional alpha As Long = 255, Optional ByVal Color As Long = -1)
    Dim bs As Long, Colour As Long

    If Color = -1 Then
        Colour = DX8Colour(White, alpha)
    Else
        Colour = Color
    End If

    Select Case design

    Case DesignTypes.DesignMenuHeader
        ' render the header
        RenderTexture TextureBlank, Left, Top, 0, 0, Width, Height, 32, 32, D3DColorARGB(200, 47, 77, 29)

    Case DesignTypes.DesignMenuHover

        ' render the option
        RenderTexture TextureBlank, Left, Top, 0, 0, Width, Height, 32, 32, D3DColorARGB(200, 98, 98, 98)

    Case DesignTypes.DesignColor

        ' render the option
        RenderTexture TextureBlank, Left, Top, 0, 0, Width, Height, 32, 32, Colour

    Case DesignTypes.DesignWoodNormal
        bs = 2
        ' render the wood box
        RenderEntity_Square TextureDesign(1), Left, Top, Width, Height, bs, alpha
        ' render wood texture
        RenderTexture TextureGUI(1), Left + bs, Top + bs, 100, 100, Width - (bs * 2), Height - (bs * 2), Width - (bs * 2), Height - (bs * 2), Colour

    Case DesignTypes.DesignWoodEmpty

        bs = 2
        ' render the wood box
        RenderEntity_Square TextureDesign(2), Left, Top, Width, Height, bs, alpha

    Case DesignTypes.DesignGreenNormal
        bs = 2
        ' render the green box
        RenderEntity_Square TextureDesign(9), Left, Top, Width, Height, bs, alpha
        ' render green gradient overlay
        RenderTexture TextureGradient(1), Left + bs, Top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.DesignGreenHover
        bs = 2
        ' render the green box
        RenderEntity_Square TextureDesign(9), Left, Top, Width, Height, bs, alpha
        ' render green gradient overlay
        RenderTexture TextureGradient(2), Left + bs, Top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.DesignGreenClick
        bs = 2
        ' render the green box
        RenderEntity_Square TextureDesign(9), Left, Top, Width, Height, bs, alpha
        ' render green gradient overlay
        RenderTexture TextureGradient(3), Left + bs, Top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.DesignRedNormal
        bs = 2
        ' render the red box
        RenderEntity_Square TextureDesign(10), Left, Top, Width, Height, bs, alpha
        ' render red gradient overlay
        RenderTexture TextureGradient(4), Left + bs, Top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.DesignRedHover
        bs = 2
        ' render the red box
        RenderEntity_Square TextureDesign(10), Left, Top, Width, Height, bs, alpha
        ' render red gradient overlay
        RenderTexture TextureGradient(5), Left + bs, Top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.DesignRedClick
        bs = 2
        ' render the red box
        RenderEntity_Square TextureDesign(10), Left, Top, Width, Height, bs, alpha
        ' render red gradient overlay
        RenderTexture TextureGradient(6), Left + bs, Top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.DesignBlueNormal
        bs = 2
        ' render the Blue box
        RenderEntity_Square TextureDesign(11), Left, Top, Width, Height, bs, alpha
        ' render Blue gradient overlay
        RenderTexture TextureGradient(7), Left + bs, Top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.DesignBlueHover
        bs = 2
        ' render the Blue box
        RenderEntity_Square TextureDesign(11), Left, Top, Width, Height, bs, alpha
        ' render Blue gradient overlay
        RenderTexture TextureGradient(8), Left + bs, Top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.DesignBlueClick
        bs = 2
        ' render the Blue box
        RenderEntity_Square TextureDesign(11), Left, Top, Width, Height, bs, alpha
        ' render Blue gradient overlay
        RenderTexture TextureGradient(9), Left + bs, Top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.DesignGoldNormal
        bs = 2
        ' render the Orange box
        RenderEntity_Square TextureDesign(12), Left, Top, Width, Height, bs, alpha
        ' render Orange gradient overlay
        RenderTexture TextureGradient(10), Left + bs, Top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.DesignGoldHover
        bs = 2
        ' render the Orange box
        RenderEntity_Square TextureDesign(12), Left, Top, Width, Height, bs, alpha
        ' render Orange gradient overlay
        RenderTexture TextureGradient(11), Left + bs, Top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.DesignGoldClick
        bs = 2
        ' render the Orange box
        RenderEntity_Square TextureDesign(12), Left, Top, Width, Height, bs, alpha
        ' render Orange gradient overlay
        RenderTexture TextureGradient(12), Left + bs, Top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.DesignGrey
        bs = 2
        ' render the Orange box
        RenderEntity_Square TextureDesign(13), Left, Top, Width, Height, bs, alpha
        ' render Orange gradient overlay
        RenderTexture TextureGradient(13), Left + bs, Top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.DesignParchment
        bs = 2
        ' render the parchment box
        RenderEntity_Square TextureDesign(7), Left, Top, Width, Height, bs, alpha

    Case DesignTypes.DesignBlackParchment
        bs = 4
        ' render the black oval
        RenderEntity_Square TextureDesign(5), Left, Top, Width, Height, bs, alpha

    Case DesignTypes.DesignTextInput
        bs = 5
        ' render the black oval
        RenderEntity_Square TextureDesign(6), Left, Top, Width, Height, bs, alpha

    Case DesignTypes.designWindowDescription
        bs = 8
        ' render black square
        RenderEntity_Square TextureDesign(3), Left, Top, Width, Height, bs, alpha

    Case DesignTypes.designWindowShadow
        bs = 35
        ' render the green box
        RenderEntity_Square TextureDesign(4), Left - bs, Top - bs, Width + (bs * 2), Height + (bs * 2), bs, alpha

    Case DesignTypes.designTilesetGrid
        bs = 16
        ' render box
        RenderEntity_Square TextureDesign(8), Left, Top, Width, Height, bs, alpha
    End Select

End Sub

Public Sub RenderEntity_Square(texNum As Long, X As Long, Y As Long, Width As Long, Height As Long, borderSize As Long, Optional alpha As Long = 255)
    Dim bs As Long, Colour As Long
    ' change colour for alpha
    Colour = DX8Colour(White, alpha)
    ' Set the border size
    bs = borderSize
    ' Draw centre
    RenderTexture texNum, X + bs, Y + bs, bs + 1, bs + 1, Width - (bs * 2), Height - (bs * 2), 1, 1, Colour
    ' Draw top side
    RenderTexture texNum, X + bs, Y, bs, 0, Width - (bs * 2), bs, 1, bs, Colour
    ' Draw left side
    RenderTexture texNum, X, Y + bs, 0, bs, bs, Height - (bs * 2), bs, 1, Colour
    ' Draw right side
    RenderTexture texNum, X + Width - bs, Y + bs, bs + 3, bs, bs, Height - (bs * 2), bs, 1, Colour
    ' Draw bottom side
    RenderTexture texNum, X + bs, Y + Height - bs, bs, bs + 3, Width - (bs * 2), bs, 1, bs, Colour
    ' Draw top left corner
    RenderTexture texNum, X, Y, 0, 0, bs, bs, bs, bs, Colour
    ' Draw top right corner
    RenderTexture texNum, X + Width - bs, Y, bs + 3, 0, bs, bs, bs, bs, Colour
    ' Draw bottom left corner
    RenderTexture texNum, X, Y + Height - bs, 0, bs + 3, bs, bs, bs, bs, Colour
    ' Draw bottom right corner
    RenderTexture texNum, X + Width - bs, Y + Height - bs, bs + 3, bs + 3, bs, bs, bs, bs, Colour
End Sub

Sub Combobox_AddItem(winIndex As Long, controlIndex As Long, text As String)
    Dim Count As Long
    Count = UBound(Windows(winIndex).Controls(controlIndex).list)
    ReDim Preserve Windows(winIndex).Controls(controlIndex).list(0 To Count + 1)
    Windows(winIndex).Controls(controlIndex).list(Count + 1) = text
End Sub

Public Sub CreateWindow(Name As String, caption As String, zOrder As Long, Left As Long, Top As Long, Width As Long, Height As Long, icon As Long, _
                        Optional visible As Boolean = True, Optional font As Long = Fonts.georgia_16, Optional textColour As Long = White, Optional xOffset As Long, _
                        Optional yOffset As Long, Optional design_norm As Long, Optional design_hover As Long, Optional design_mousedown As Long, Optional image_norm As Long, _
                        Optional image_hover As Long, Optional image_mousedown As Long, Optional entCallBack_norm As Long, Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, _
                        Optional entCallBack_mousemove As Long, Optional entCallBack_DblClick As Long, Optional canDrag As Boolean = True, Optional zChange As Byte = True, Optional ByVal onDraw As Long, _
                        Optional isActive As Boolean, Optional clickThrough As Boolean)

    Dim i As Long
    Dim design(0 To EntityStates.StateCount - 1) As Long
    Dim image(0 To EntityStates.StateCount - 1) As Long
    Dim entCallBack(0 To EntityStates.StateCount - 1) As Long

    ' fill temp arrays
    design(EntityStates.Normal) = design_norm
    design(EntityStates.Hover) = design_hover
    design(EntityStates.MouseDown) = design_mousedown
    design(EntityStates.DoubleClick) = design_norm
    design(EntityStates.MouseUp) = design_norm
    image(EntityStates.Normal) = image_norm
    image(EntityStates.Hover) = image_hover
    image(EntityStates.MouseDown) = image_mousedown
    image(EntityStates.DoubleClick) = image_norm
    image(EntityStates.MouseUp) = image_norm
    entCallBack(EntityStates.Normal) = entCallBack_norm
    entCallBack(EntityStates.Hover) = entCallBack_hover
    entCallBack(EntityStates.MouseDown) = entCallBack_mousedown
    entCallBack(EntityStates.MouseMove) = entCallBack_mousemove
    entCallBack(EntityStates.DoubleClick) = entCallBack_DblClick
    ' redim the windows

    If Not windowUpdated Then
        windowCount = windowCount + 1
        ReDim Preserve Windows(1 To windowCount) As WindowRec
    End If
    ' set the properties
    With Windows(windowCount).Window
        .Name = Name
        .Type = EntityTypes.EntityWindow

        ' loop through states
        For i = 0 To EntityStates.StateCount - 1

            .design(i) = design(i)
            .image(i) = image(i)
            .entCallBack(i) = entCallBack(i)
        Next

        .Left = Left
        .Top = Top
        .origLeft = Left
        .origTop = Top
        .Width = Width
        .Height = Height
        .visible = visible
        .canDrag = canDrag
        .text = caption
        .font = font
        .textColour = textColour
        .xOffset = xOffset
        .yOffset = yOffset
        .icon = icon
        .enabled = True
        .zChange = zChange
        .zOrder = zOrder
        .onDraw = onDraw
        .clickThrough = clickThrough
        ' set active
        If .visible Then activeWindow = windowCount
    End With

    ' set the zOrder
    zOrder_Win = zOrder_Win + 1
End Sub

Public Sub CreateTextbox(winNum As Long, Name As String, Left As Long, Top As Long, Width As Long, Height As Long, Optional text As String, Optional font As Long = Fonts.georgia_16, _
                         Optional textColour As Long = White, Optional align As Byte = Alignment.AlignLeft, Optional visible As Boolean = True, Optional alpha As Long = 255, Optional image_norm As Long, _
                         Optional image_hover As Long, Optional image_mousedown As Long, Optional design_norm As Long, Optional design_hover As Long, Optional design_mousedown As Long, _
                         Optional entCallBack_norm As Long, Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, Optional entCallBack_mousemove As Long, Optional entCallBack_DblClick As Long, _
                         Optional isActive As Boolean, Optional xOffset As Long, Optional yOffset As Long, Optional isCensor As Boolean, Optional entCallBack_enter As Long)
    Dim design(0 To EntityStates.StateCount - 1) As Long
    Dim image(0 To EntityStates.StateCount - 1) As Long
    Dim entCallBack(0 To EntityStates.StateCount - 1) As Long
    ' fill temp arrays
    design(EntityStates.Normal) = design_norm
    design(EntityStates.Hover) = design_hover
    design(EntityStates.MouseDown) = design_mousedown
    image(EntityStates.Normal) = image_norm
    image(EntityStates.Hover) = image_hover
    image(EntityStates.MouseDown) = image_mousedown
    entCallBack(EntityStates.Normal) = entCallBack_norm
    entCallBack(EntityStates.Hover) = entCallBack_hover
    entCallBack(EntityStates.MouseDown) = entCallBack_mousedown
    entCallBack(EntityStates.MouseMove) = entCallBack_mousemove
    entCallBack(EntityStates.DoubleClick) = entCallBack_DblClick
    entCallBack(EntityStates.Enter) = entCallBack_enter
    ' create the textbox
    CreateEntity winNum, zOrder_Con, Name, EntityTextBox, design(), image(), entCallBack(), Left, Top, Width, Height, visible, , , , , text, align, font, textColour, alpha, , xOffset, yOffset, , , , isActive, isCensor
End Sub

Public Sub CreatePictureBox(winNum As Long, Name As String, Left As Long, Top As Long, Width As Long, Height As Long, Optional visible As Boolean = True, Optional canDrag As Boolean, _
                            Optional alpha As Long = 255, Optional clickThrough As Boolean, Optional image_norm As Long, Optional image_hover As Long, Optional image_mousedown As Long, Optional design_norm As Long, _
                            Optional design_hover As Long, Optional design_mousedown As Long, Optional entCallBack_norm As Long, Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, _
                            Optional entCallBack_mousemove As Long, Optional entCallBack_DblClick As Long, Optional onDraw As Long)
    Dim design(0 To EntityStates.StateCount - 1) As Long
    Dim image(0 To EntityStates.StateCount - 1) As Long
    Dim entCallBack(0 To EntityStates.StateCount - 1) As Long
    ' fill temp arrays
    design(EntityStates.Normal) = design_norm
    design(EntityStates.Hover) = design_hover
    design(EntityStates.MouseDown) = design_mousedown
    image(EntityStates.Normal) = image_norm
    image(EntityStates.Hover) = image_hover
    image(EntityStates.MouseDown) = image_mousedown
    entCallBack(EntityStates.Normal) = entCallBack_norm
    entCallBack(EntityStates.Hover) = entCallBack_hover
    entCallBack(EntityStates.MouseDown) = entCallBack_mousedown
    entCallBack(EntityStates.MouseMove) = entCallBack_mousemove
    entCallBack(EntityStates.DoubleClick) = entCallBack_DblClick
    ' create the box
    CreateEntity winNum, zOrder_Con, Name, EntityPictureBox, design(), image(), entCallBack(), Left, Top, Width, Height, visible, canDrag, , , , , , , , alpha, clickThrough, , , , , onDraw

End Sub

Public Sub CreateButton(winNum As Long, Name As String, Left As Long, Top As Long, Width As Long, Height As Long, Optional text As String, Optional font As Fonts = Fonts.georgia_16, _
                        Optional textColour As Long = White, Optional icon As Long, Optional visible As Boolean = True, Optional alpha As Long = 255, Optional image_norm As Long, Optional image_hover As Long, _
                        Optional image_mousedown As Long, Optional design_norm As Long, Optional design_hover As Long, Optional design_mousedown As Long, Optional entCallBack_norm As Long, _
                        Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, Optional entCallBack_mousemove As Long, Optional entCallBack_DblClick As Long, Optional xOffset As Long, _
                        Optional yOffset As Long, Optional textColourHover As Long = -1, Optional textColourClick As Long = -1, Optional tooltip As String)
    Dim design(0 To EntityStates.StateCount - 1) As Long
    Dim image(0 To EntityStates.StateCount - 1) As Long
    Dim entCallBack(0 To EntityStates.StateCount - 1) As Long

    ' default the colours
    If textColourHover = -1 Then textColourHover = textColour
    If textColourClick = -1 Then textColourClick = textColour
    ' fill temp arrays
    design(EntityStates.Normal) = design_norm
    design(EntityStates.Hover) = design_hover
    design(EntityStates.MouseDown) = design_mousedown
    image(EntityStates.Normal) = image_norm
    image(EntityStates.Hover) = image_hover
    image(EntityStates.MouseDown) = image_mousedown
    entCallBack(EntityStates.Normal) = entCallBack_norm
    entCallBack(EntityStates.Hover) = entCallBack_hover
    entCallBack(EntityStates.MouseDown) = entCallBack_mousedown
    entCallBack(EntityStates.MouseMove) = entCallBack_mousemove
    entCallBack(EntityStates.DoubleClick) = entCallBack_DblClick
    ' create the box
    CreateEntity winNum, zOrder_Con, Name, EntityButton, design(), image(), entCallBack(), Left, Top, Width, Height, visible, , , , , text, , font, textColour, alpha, , xOffset, yOffset, , icon, , , , textColourHover, textColourClick, tooltip
End Sub

Public Sub CreateLabel(winNum As Long, Name As String, Left As Long, Top As Long, Width As Long, Optional Height As Long, Optional text As String, Optional font As Fonts = Fonts.georgia_16, _
                       Optional textColour As Long = White, Optional align As Byte = Alignment.AlignLeft, Optional visible As Boolean = True, Optional alpha As Long = 255, Optional clickThrough As Boolean, _
                       Optional entCallBack_norm As Long, Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, Optional entCallBack_mousemove As Long, Optional entCallBack_DblClick As Long)
    Dim design(0 To EntityStates.StateCount - 1) As Long
    Dim image(0 To EntityStates.StateCount - 1) As Long
    Dim entCallBack(0 To EntityStates.StateCount - 1) As Long
    ' fill temp arrays
    entCallBack(EntityStates.Normal) = entCallBack_norm
    entCallBack(EntityStates.Hover) = entCallBack_hover
    entCallBack(EntityStates.MouseDown) = entCallBack_mousedown
    entCallBack(EntityStates.MouseMove) = entCallBack_mousemove
    entCallBack(EntityStates.DoubleClick) = entCallBack_DblClick
    ' create the box
    CreateEntity winNum, zOrder_Con, Name, EntityLabel, design(), image(), entCallBack(), Left, Top, Width, Height, visible, , , , , text, align, font, textColour, alpha, clickThrough
End Sub

Public Sub CreateCheckbox(winNum As Long, Name As String, Left As Long, Top As Long, Width As Long, Optional Height As Long = 15, Optional Value As Long, Optional text As String, _
                          Optional font As Fonts = Fonts.georgia_16, Optional textColour As Long = White, Optional align As Byte = Alignment.AlignLeft, Optional visible As Boolean = True, Optional alpha As Long = 255, _
                          Optional theDesign As Long, Optional entCallBack_norm As Long, Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, Optional entCallBack_mousemove As Long, _
                          Optional entCallBack_DblClick As Long, Optional group As Long)
    Dim design(0 To EntityStates.StateCount - 1) As Long
    Dim image(0 To EntityStates.StateCount - 1) As Long
    Dim entCallBack(0 To EntityStates.StateCount - 1) As Long
    ' fill temp arrays
    entCallBack(EntityStates.Normal) = entCallBack_norm
    entCallBack(EntityStates.Hover) = entCallBack_hover
    entCallBack(EntityStates.MouseDown) = entCallBack_mousedown
    entCallBack(EntityStates.MouseMove) = entCallBack_mousemove
    entCallBack(EntityStates.DoubleClick) = entCallBack_DblClick
    ' fill temp array
    design(0) = theDesign
    ' create the box
    CreateEntity winNum, zOrder_Con, Name, EntityCheckbox, design(), image(), entCallBack(), Left, Top, Width, Height, visible, , , , Value, text, align, font, textColour, alpha, , , , , , , , , , , , group
End Sub

Public Sub CreateComboBox(winNum As Long, Name As String, Left As Long, Top As Long, Width As Long, Height As Long, design As Long, Optional font As Fonts = Fonts.georgia_16)
    Dim theDesign(0 To EntityStates.StateCount - 1) As Long
    Dim image(0 To EntityStates.StateCount - 1) As Long
    Dim entCallBack(0 To EntityStates.StateCount - 1) As Long
    theDesign(0) = design
    ' create the box
    CreateEntity winNum, zOrder_Con, Name, entityCombo, theDesign(), image(), entCallBack(), Left, Top, Width, Height, , , , , , , , font
End Sub

Public Function GetWindowIndex(winName As String) As Long
    Dim i As Long

    For i = 1 To windowCount

        If LCase$(Windows(i).Window.Name) = LCase$(winName) Then
            GetWindowIndex = i
            Exit Function
        End If

    Next

    GetWindowIndex = 0
End Function

Public Function GetControlIndex(winName As String, controlName As String) As Long
    Dim i As Long, winIndex As Long

    winIndex = GetWindowIndex(winName)

    If Not winIndex > 0 Or Not winIndex <= windowCount Then Exit Function

    For i = 1 To Windows(winIndex).ControlCount

        If LCase$(Windows(winIndex).Controls(i).Name) = LCase$(controlName) Then
            GetControlIndex = i
            Exit Function
        End If

    Next

    GetControlIndex = 0
End Function

Public Function SetActiveControl(curWindow As Long, curControl As Long) As Boolean
' make sure it's something which CAN be active
    Select Case Windows(curWindow).Controls(curControl).Type
    Case EntityTypes.EntityTextBox

        Windows(curWindow).activeControl = curControl
        SetActiveControl = True
    End Select
End Function

Public Sub CentraliseWindow(curWindow As Long)
    With Windows(curWindow).Window
        .Left = (ScreenWidth / 2) - (.Width / 2)
        .Top = (ScreenHeight / 2) - (.Height / 2)
        .origLeft = .Left
        .origTop = .Top
    End With
End Sub

Public Sub HideWindows()
    Dim i As Long
    For i = 1 To windowCount
        HideWindow i
    Next
End Sub

Public Sub ShowWindow(curWindow As Long, Optional forced As Boolean, Optional resetPosition As Boolean = True)
    Windows(curWindow).Window.visible = True
    If forced Then
        UpdateZOrder curWindow, forced
        activeWindow = curWindow
    ElseIf Windows(curWindow).Window.zChange Then
        UpdateZOrder curWindow
        activeWindow = curWindow
    End If
    If resetPosition Then
        With Windows(curWindow).Window
            .Left = .origLeft
            .Top = .origTop
        End With
    End If
End Sub

Public Sub HideWindow(curWindow As Long)
    Dim i As Long
    Windows(curWindow).Window.visible = False
    ' find next window to set as active
    For i = windowCount To 1 Step -1
        If Windows(i).Window.visible And Windows(i).Window.zChange Then
            'UpdateZOrder i
            activeWindow = i
            Exit Sub
        End If
    Next
End Sub

Public Sub CreateWindow_Login()
' Definição da Janela
    CreateWindow "winLogin", "Login", zOrder_Win, 0, 0, 272, 227, TextureItem(45), , Fonts.Default, White, 3, 5, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal

    ' Centralizar a Janela
    CentraliseWindow windowCount

    ' Ordem da Janela
    zOrder_Con = 1

    ' Botão de Fechar
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , , , , GetAddress(AddressOf DestroyGame)

    ' Pergaminho
    CreatePictureBox windowCount, "picParchment", 8, WindowTopBar + 6, 256, 173, , , , , , , , DesignTypes.DesignParchment, DesignTypes.DesignParchment, DesignTypes.DesignParchment

    ' Sombras
    CreatePictureBox windowCount, "picShadow_1", 15, WindowTopBar + 14, 242, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreatePictureBox windowCount, "picShadow_2", 15, WindowTopBar + 56, 242, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment

    ' Textos
    CreateLabel windowCount, "lblUsername", 15, WindowTopBar + 10, 242, , "Usuário", Default, White, Alignment.AlignCentre
    CreateLabel windowCount, "lblPassword", 15, WindowTopBar + 52, 242, , "Senha", Default, White, Alignment.AlignCentre

    ' Textboxes
    CreateTextbox windowCount, "txtUser", 15, WindowTopBar + 27, 242, 24, Options.Username, Fonts.Default, DarkGrey, Alignment.AlignLeft, , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, , , , , , , 6, 4
    CreateTextbox windowCount, "txtPass", 15, WindowTopBar + 69, 242, 24, vbNullString, Fonts.Default, DarkGrey, Alignment.AlignLeft, , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, , , , , , , 6, 4, True, GetAddress(AddressOf btnLogin_Click)

    ' Botões
    CreateButton windowCount, "btnAccept", 15, WindowTopBar + 108, 242, 30, "Acessar", Default, White, , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnLogin_Click)
    CreateButton windowCount, "btnRegister", 15, WindowTopBar + 142, 242, 30, "Criar uma nova Conta", Default, White, , , , , , , DesignTypes.DesignGoldNormal, DesignTypes.DesignGoldHover, DesignTypes.DesignGoldClick, , , GetAddress(AddressOf btnRegister_Click)

    ' Set the active control
    If Not Len(Windows(GetWindowIndex("winLogin")).Controls(GetControlIndex("winLogin", "txtUser")).text) > 0 Then
        '    SetActiveControl GetWindowIndex("winLogin"), GetControlIndex("winLogin", "txtUser")
    Else
        '    SetActiveControl GetWindowIndex("winLogin"), GetControlIndex("winLogin", "txtPass")
    End If
End Sub

Public Sub StrongPassword()
    Dim X, Y, passwordLenght, Colour As Long

    X = Windows(GetWindowIndex("winRegister")).Window.Left
    Y = Windows(GetWindowIndex("winRegister")).Window.Top
    passwordLenght = Len(Windows(GetWindowIndex("winRegister")).Controls(GetControlIndex("winRegister", "txtPass2")).text)

    If passwordLenght <= 6 Then
        Colour = D3DColorARGB(200, 200, 60, 60)
    ElseIf passwordLenght <= 10 Then
        Colour = D3DColorARGB(200, 220, 200, 50)
    ElseIf passwordLenght > 10 Then
        Colour = D3DColorARGB(200, 80, 220, 50)
    End If

    RenderDesign DesignTypes.DesignColor, X + 15, Y + (WindowTopBar + 136), 242, 4, , Colour
End Sub

Public Sub CreateWindow_Register()

' Definição da Janela
    CreateWindow "winRegister", "Cadastrar uma nova conta", zOrder_Win, 0, 0, 272, 359, TextureItem(45), , Fonts.Default, , 3, 5, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal

    ' Centralizar a Janela
    CentraliseWindow windowCount

    ' Ordem da Janela
    zOrder_Con = 1

    ' Botão de Fechar
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , , , , GetAddress(AddressOf btnReturnMain_Click)

    ' Pergaminho
    CreatePictureBox windowCount, "picParchment", 8, WindowTopBar + 6, 256, 305, , , , , , , , DesignTypes.DesignParchment, DesignTypes.DesignParchment, DesignTypes.DesignParchment

    ' Sombras
    CreatePictureBox windowCount, "picShadow_1", 15, WindowTopBar + 14, 242, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreatePictureBox windowCount, "picShadow_2", 15, WindowTopBar + 56, 242, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreatePictureBox windowCount, "picShadow_3", 15, WindowTopBar + 98, 242, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreatePictureBox windowCount, "picShadow_4", 15, WindowTopBar + 146, 242, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreatePictureBox windowCount, "picShadow_5", 15, WindowTopBar + 188, 242, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment

    ' Textos
    CreateLabel windowCount, "lblUsername", 15, (WindowTopBar + 14) - 4, 242, , "Usuário", Fonts.Default, White, Alignment.AlignCentre
    CreateLabel windowCount, "lblPassword", 15, (WindowTopBar + 56) - 4, 242, , "Senha", Fonts.Default, White, Alignment.AlignCentre
    CreateLabel windowCount, "lblPassword2", 15, (WindowTopBar + 98) - 4, 242, , "Repetir Senha", Fonts.Default, White, Alignment.AlignCentre
    CreateLabel windowCount, "lblCode", 15, (WindowTopBar + 146) - 4, 242 - 4, , "Código Secreto", Fonts.Default, White, Alignment.AlignCentre
    CreateLabel windowCount, "lblCaptcha", 15, (WindowTopBar + 188) - 4, 242, , "Captcha", Fonts.Default, White, Alignment.AlignCentre

    ' Textboxes
    CreateTextbox windowCount, "txtAccount", 15, WindowTopBar + 27, 242, 23, vbNullString, Fonts.Default, DarkGrey, Alignment.AlignLeft, , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, , , , , , , 6, 4, False, GetAddress(AddressOf btnSendRegister_Click)
    CreateTextbox windowCount, "txtPass", 15, WindowTopBar + 69, 242, 23, vbNullString, Fonts.Default, DarkGrey, Alignment.AlignLeft, , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, , , , , , , 6, 4, True, GetAddress(AddressOf btnSendRegister_Click)
    CreateTextbox windowCount, "txtPass2", 15, WindowTopBar + 111, 242, 23, vbNullString, Fonts.Default, DarkGrey, Alignment.AlignLeft, , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, , , , , , , 6, 4, True, GetAddress(AddressOf btnSendRegister_Click)
    CreateTextbox windowCount, "txtCode", 15, WindowTopBar + 160, 242, 23, vbNullString, Fonts.Default, DarkGrey, Alignment.AlignLeft, , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, , , , , , , 6, 4, False, GetAddress(AddressOf btnSendRegister_Click)
    CreateTextbox windowCount, "txtCaptcha", 15, WindowTopBar + 235, 242, 23, vbNullString, Fonts.Default, DarkGrey, Alignment.AlignLeft, , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, , , , , , , 6, 4, False, GetAddress(AddressOf btnSendRegister_Click)


    ' Captcha
    CreatePictureBox windowCount, "picCaptcha", 15, WindowTopBar + 201, 242, 30, , , , , TextureCaptcha(GlobalCaptcha), TextureCaptcha(GlobalCaptcha), TextureCaptcha(GlobalCaptcha), DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment

    ' Botões
    CreateButton windowCount, "btnAccept", 15, WindowTopBar + 274, 242, 30, "Criar a Conta", Fonts.Default, White, , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnSendRegister_Click)

    ' Senha Forte
    CreatePictureBox windowCount, "strongPassword", 15, WindowTopBar + 136, 242, 4, True, False, , , , , , DesignColor, DesignColor, DesignColor, , , , , , GetAddress(AddressOf StrongPassword)

    SetActiveControl GetWindowIndex("winRegister"), GetControlIndex("winRegister", "txtAccount")
End Sub

Public Sub CreateWindow_Characters()
' Create the window
    CreateWindow "winCharacters", "Characters", zOrder_Win, 0, 0, 364, 229, TextureItem(62), False, Fonts.rockwellDec_15, , 3, 5, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal
    ' Centralise it
    CentraliseWindow windowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , , , , GetAddress(AddressOf btnCharacters_Close)
    ' Parchment
    CreatePictureBox windowCount, "picParchment", 6, 26, 352, 197, , , , , , , , DesignTypes.DesignParchment, DesignTypes.DesignParchment, DesignTypes.DesignParchment
    ' Names
    CreatePictureBox windowCount, "picShadow_1", 22, 41, 98, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblCharName_1", 22, 37, 98, , "Blank Slot", rockwellDec_15, White, Alignment.AlignCentre
    CreatePictureBox windowCount, "picShadow_2", 132, 41, 98, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblCharName_2", 132, 37, 98, , "Blank Slot", rockwellDec_15, White, Alignment.AlignCentre
    CreatePictureBox windowCount, "picShadow_3", 242, 41, 98, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblCharName_3", 242, 37, 98, , "Blank Slot", rockwellDec_15, White, Alignment.AlignCentre
    ' Scenery Boxes
    CreatePictureBox windowCount, "picScene_1", 23, 55, 96, 96, , , , , TextureGUI(2), TextureGUI(2), TextureGUI(2)
    CreatePictureBox windowCount, "picScene_2", 133, 55, 96, 96, , , , , TextureGUI(2), TextureGUI(2), TextureGUI(2)
    CreatePictureBox windowCount, "picScene_3", 243, 55, 96, 96, , , , , TextureGUI(2), TextureGUI(2), TextureGUI(2), , , , , , , , , GetAddress(AddressOf Chars_DrawFace)
    ' Create Buttons
    CreateButton windowCount, "btnSelectChar_1", 22, 155, 98, 24, "Select", rockwellDec_15, , , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnAcceptChar_1)
    CreateButton windowCount, "btnCreateChar_1", 22, 155, 98, 24, "Create", rockwellDec_15, , , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnCreateChar_1)
    CreateButton windowCount, "btnDelChar_1", 22, 183, 98, 24, "Delete", rockwellDec_15, , , , , , , , DesignTypes.DesignRedNormal, DesignTypes.DesignRedHover, DesignTypes.DesignRedClick, , , GetAddress(AddressOf btnDelChar_1)
    CreateButton windowCount, "btnSelectChar_2", 132, 155, 98, 24, "Select", rockwellDec_15, , , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnAcceptChar_2)
    CreateButton windowCount, "btnCreateChar_2", 132, 155, 98, 24, "Create", rockwellDec_15, , , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnCreateChar_2)
    CreateButton windowCount, "btnDelChar_2", 132, 183, 98, 24, "Delete", rockwellDec_15, , , , , , , , DesignTypes.DesignRedNormal, DesignTypes.DesignRedHover, DesignTypes.DesignRedClick, , , GetAddress(AddressOf btnDelChar_2)
    CreateButton windowCount, "btnSelectChar_3", 242, 155, 98, 24, "Select", rockwellDec_15, , , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnAcceptChar_3)
    CreateButton windowCount, "btnCreateChar_3", 242, 155, 98, 24, "Create", rockwellDec_15, , , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnCreateChar_3)
    CreateButton windowCount, "btnDelChar_3", 242, 183, 98, 24, "Delete", rockwellDec_15, , , , , , , , DesignTypes.DesignRedNormal, DesignTypes.DesignRedHover, DesignTypes.DesignRedClick, , , GetAddress(AddressOf btnDelChar_3)
End Sub

Public Sub CreateWindow_Loading()
' Create the window
    CreateWindow "winLoading", "Loading", zOrder_Win, 0, 0, 278, 79, TextureItem(104), True, Fonts.rockwellDec_15, , 2, 7, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal
    ' Centralise it
    CentraliseWindow windowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Parchment
    CreatePictureBox windowCount, "picParchment", 6, 26, 266, 47, , , , , , , , DesignTypes.DesignParchment, DesignTypes.DesignParchment, DesignTypes.DesignParchment
    ' Text background
    CreatePictureBox windowCount, "picRecess", 26, 39, 226, 22, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    ' Label
    CreateLabel windowCount, "lblLoading", 6, 43, 266, , "Loading Game Data...", rockwell_15, , Alignment.AlignCentre
End Sub

Public Sub CreateWindow_Dialogue()
' Create black background
    CreateWindow "winBlank", "", zOrder_Win, 0, 0, 800, 600, 0, , , , , , DesignTypes.designWindowDescription, DesignTypes.designWindowDescription, DesignTypes.designWindowDescription, , , , , , , , , False, False
    ' Create dialogue window
    CreateWindow "winDialogue", "Warning", zOrder_Win, 0, 0, 348, 145, TextureItem(38), , Fonts.rockwellDec_15, , 3, 5, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal, , , , , , , , , , False
    ' Centralise it
    CentraliseWindow windowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , , , , GetAddress(AddressOf btnDialogue_Close)

    ' Parchment
    CreatePictureBox windowCount, "picParchment", 6, 26, 335, 113, , , , , , , , DesignTypes.DesignParchment, DesignTypes.DesignParchment, DesignTypes.DesignParchment
    ' Header
    CreatePictureBox windowCount, "picShadow", 103, 44, 144, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblHeader", 103, 41, 144, , "Header", rockwellDec_15, White, Alignment.AlignCentre
    ' Labels
    CreateLabel windowCount, "lblBody_1", 15, 60, 314, , "Invalid username or password.", rockwell_15, , Alignment.AlignCentre
    CreateLabel windowCount, "lblBody_2", 15, 75, 314, , "Please try again.", rockwell_15, , Alignment.AlignCentre
    ' Buttons
    CreateButton windowCount, "btnYes", 104, 98, 68, 24, "Yes", rockwellDec_15, , , False, , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf Dialogue_Yes)
    CreateButton windowCount, "btnNo", 180, 98, 68, 24, "No", rockwellDec_15, , , False, , , , , DesignTypes.DesignRedNormal, DesignTypes.DesignRedHover, DesignTypes.DesignRedClick, , , GetAddress(AddressOf Dialogue_No)
    CreateButton windowCount, "btnOkay", 140, 98, 68, 24, "Okay", rockwellDec_15, , , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf Dialogue_Okay)
    ' Input
    CreateTextbox windowCount, "txtInput", 93, 75, 162, 18, , rockwell_15, White, Alignment.AlignCentre, , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , , , , , , 4, 2
    ' set active control
    SetActiveControl windowCount, GetControlIndex("winDialogue", "txtInput")
End Sub

Public Sub CreateWindow_Classes()
' Create window
    CreateWindow "winClasses", "Select Class", zOrder_Win, 0, 0, 364, 229, TextureItem(17), False, Fonts.rockwellDec_15, , 2, 6, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal

    ' Centralise it
    CentraliseWindow windowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , , , , GetAddress(AddressOf btnClasses_Close)
    ' Parchment
    CreatePictureBox windowCount, "picParchment", 6, 26, 352, 197, , , , , , , , DesignTypes.DesignParchment, DesignTypes.DesignParchment, DesignTypes.DesignParchment, , , , , , GetAddress(AddressOf Classes_DrawFace)
    ' Class Name
    CreatePictureBox windowCount, "picShadow", 183, 42, 98, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblClassName", 183, 39, 98, , "Warrior", rockwellDec_15, White, Alignment.AlignCentre
    ' Select Buttons
    CreateButton windowCount, "btnLeft", 171, 40, 11, 13, , , , , , , TextureGUI(12), TextureGUI(13), TextureGUI(14), , , , , , GetAddress(AddressOf btnClasses_Left)
    CreateButton windowCount, "btnRight", 282, 40, 11, 13, , , , , , , TextureGUI(15), TextureGUI(16), TextureGUI(17), , , , , , GetAddress(AddressOf btnClasses_Right)
    ' Accept Button
    CreateButton windowCount, "btnAccept", 183, 185, 98, 22, "Accept", rockwellDec_15, , , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnClasses_Accept)
    ' Text background
    CreatePictureBox windowCount, "picBackground", 127, 55, 210, 124, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    ' Overlay
    CreatePictureBox windowCount, "picOverlay", 6, 26, 0, 0, , , , , , , , , , , , , , , , GetAddress(AddressOf Classes_DrawText)
End Sub

Public Sub CreateWindow_NewChar()
' Create window
    CreateWindow "winNewChar", "Create Character", zOrder_Win, 0, 0, 291, 172, TextureItem(17), False, Fonts.rockwellDec_15, , 2, 6, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal

    ' Centralise it
    CentraliseWindow windowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , , , , GetAddress(AddressOf btnNewChar_Cancel)
    ' Parchment
    CreatePictureBox windowCount, "picParchment", 6, 26, 278, 140, , , , , , , , DesignTypes.DesignParchment, DesignTypes.DesignParchment, DesignTypes.DesignParchment
    ' Name
    CreatePictureBox windowCount, "picShadow_1", 29, 42, 124, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblName", 29, 39, 124, , "Name", rockwellDec_15, White, Alignment.AlignCentre
    ' Textbox
    CreateTextbox windowCount, "txtName", 29, 55, 124, 19, , Fonts.rockwell_15, , Alignment.AlignLeft, , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, , , , , , , 5, 3
    ' Gender
    CreatePictureBox windowCount, "picShadow_2", 29, 85, 124, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblGender", 29, 82, 124, , "Gender", rockwellDec_15, White, Alignment.AlignCentre
    ' Checkboxes
    CreateCheckbox windowCount, "chkMale", 29, 103, 55, , 1, "Male", rockwell_15, , Alignment.AlignCentre, , , DesignTypes.DesignCheckbox, , , GetAddress(AddressOf chkNewChar_Male), , , 1
    CreateCheckbox windowCount, "chkFemale", 90, 103, 62, , 0, "Female", rockwell_15, , Alignment.AlignCentre, , , DesignTypes.DesignCheckbox, , , GetAddress(AddressOf chkNewChar_Female), , , 1

    ' Buttons
    CreateButton windowCount, "btnAccept", 29, 127, 60, 24, "Accept", rockwellDec_15, , , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnNewChar_Accept)
    CreateButton windowCount, "btnCancel", 93, 127, 60, 24, "Cancel", rockwellDec_15, , , , , , , , DesignTypes.DesignRedNormal, DesignTypes.DesignRedHover, DesignTypes.DesignRedClick, , , GetAddress(AddressOf btnNewChar_Cancel)
    ' Sprite
    CreatePictureBox windowCount, "picShadow_3", 175, 42, 76, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblSprite", 175, 39, 76, , "Sprite", rockwellDec_15, White, Alignment.AlignCentre
    ' Scene
    CreatePictureBox windowCount, "picScene", 165, 55, 96, 96, , , , , TextureGUI(2), TextureGUI(2), TextureGUI(2), , , , , , , , , GetAddress(AddressOf NewChar_OnDraw)
    ' Buttons
    CreateButton windowCount, "btnLeft", 163, 40, 11, 13, , , , , , , TextureGUI(12), TextureGUI(13), TextureGUI(14), , , , , , GetAddress(AddressOf btnNewChar_Left)
    CreateButton windowCount, "btnRight", 252, 40, 11, 13, , , , , , , TextureGUI(15), TextureGUI(16), TextureGUI(17), , , , , , GetAddress(AddressOf btnNewChar_Right)

    ' Set the active control
    SetActiveControl GetWindowIndex("winNewChar"), GetControlIndex("winNewChar", "txtName")
End Sub

Public Sub CreateWindow_EscMenu()
' Create window
    CreateWindow "winEscMenu", "", zOrder_Win, 0, 0, 210, 156, 0, , , , , , DesignTypes.DesignWindowWithoutBar, DesignTypes.DesignWindowWithoutBar, DesignTypes.DesignWindowWithoutBar, , , , , , , , , False, False
    ' Centralise it
    CentraliseWindow windowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Parchment
    CreatePictureBox windowCount, "picParchment", 6, 6, 198, 144, , , , , , , , DesignTypes.DesignParchment, DesignTypes.DesignParchment, DesignTypes.DesignParchment
    ' Buttons
    CreateButton windowCount, "btnReturn", 16, 16, 178, 28, "Return to Game (Esc)", rockwellDec_15, , , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnEscMenu_Return)
    CreateButton windowCount, "btnOptions", 16, 48, 178, 28, "Options", rockwellDec_15, , , , , , , , DesignTypes.DesignGoldNormal, DesignTypes.DesignGoldHover, DesignTypes.DesignGoldClick, , , GetAddress(AddressOf btnEscMenu_Options)
    CreateButton windowCount, "btnMainMenu", 16, 80, 178, 28, "Back to Main Menu", rockwellDec_15, , , , , , , , DesignTypes.DesignBlueNormal, DesignTypes.DesignBlueHover, DesignTypes.DesignBlueClick, , , GetAddress(AddressOf btnEscMenu_MainMenu)
    CreateButton windowCount, "btnExit", 16, 112, 178, 28, "Exit the Game", rockwellDec_15, , , , , , , , DesignTypes.DesignRedNormal, DesignTypes.DesignRedHover, DesignTypes.DesignRedClick, , , GetAddress(AddressOf btnEscMenu_Exit)
End Sub

Public Sub CreateWindow_Bars()
' Create window
    CreateWindow "winBars", "", zOrder_Win, 10, 10, 239, 77, 0, , , , , , DesignTypes.DesignWindowWithoutBar, DesignTypes.DesignWindowWithoutBar, DesignTypes.DesignWindowWithoutBar, , , , , , , , , False, False

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Parchment
    CreatePictureBox windowCount, "picParchment", 6, 6, 227, 65, , , , , , , , DesignTypes.DesignParchment, DesignTypes.DesignParchment, DesignTypes.DesignParchment
    ' Blank Bars
    CreatePictureBox windowCount, "picHP_Blank", 15, 15, 209, 13, , , , , TextureGUI(26), TextureGUI(26), TextureGUI(26)
    CreatePictureBox windowCount, "picSP_Blank", 15, 32, 209, 13, , , , , TextureGUI(27), TextureGUI(27), TextureGUI(27)
    CreatePictureBox windowCount, "picEXP_Blank", 15, 49, 209, 13, , , , , TextureGUI(28), TextureGUI(28), TextureGUI(28)
    ' Draw the bars
    CreatePictureBox windowCount, "picBlank", 0, 0, 0, 0, , , , , , , , , , , , , , , , GetAddress(AddressOf Bars_OnDraw)
    ' Labels
    CreateLabel windowCount, "lblHP", 15, 14, 209, 12, "999/999", rockwellDec_10, White, Alignment.AlignCentre
    CreateLabel windowCount, "lblMP", 15, 31, 209, 12, "999/999", rockwellDec_10, White, Alignment.AlignCentre
    CreateLabel windowCount, "lblEXP", 15, 48, 209, 12, "999/999", rockwellDec_10, White, Alignment.AlignCentre
End Sub

Public Sub CreateWindow_Menu()
' Create window
    CreateWindow "winMenu", "", zOrder_Win, 564, 563, 229, 31, 0, , , , , , , , , , , , , , , , , False, False

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Wood part
    CreatePictureBox windowCount, "picWood", 0, 5, 228, 21, , , , , , , , DesignTypes.DesignWoodNormal, DesignTypes.DesignWoodNormal, DesignTypes.DesignWoodNormal
    ' Buttons
    CreateButton windowCount, "btnChar", 8, 1, 29, 29, , , , TextureItem(108), , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnMenu_Char), , , -1, -2, , , "Character (C)"
    CreateButton windowCount, "btnInv", 44, 1, 29, 29, , , , TextureItem(1), , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnMenu_Inv), , , -1, -2, , , "Inventory (I)"
    CreateButton windowCount, "btnSkills", 82, 1, 29, 29, , , , TextureItem(109), , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnMenu_Skills), , , -1, -2, , , "Skills (M)"
    CreateButton windowCount, "btnMap", 119, 1, 29, 29, , , , TextureItem(106), , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnMenu_Map), , , -1, -2
    CreateButton windowCount, "btnGuild", 155, 1, 29, 29, , , , TextureItem(107), , , , , , DesignTypes.DesignGrey, DesignTypes.DesignGrey, DesignTypes.DesignGrey, , , GetAddress(AddressOf btnMenu_Guild), , , -1, -1
    CreateButton windowCount, "btnQuest", 191, 1, 29, 29, , , , TextureItem(23), , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnMenu_Quest), , , -1, -2
End Sub

Public Sub CreateWindow_Hotbar()
' Create window
    CreateWindow "winHotbar", "", zOrder_Win, 372, 10, 418, 36, 0, , , , , , , , , , , , , GetAddress(AddressOf Hotbar_MouseMove), GetAddress(AddressOf Hotbar_MouseDown), GetAddress(AddressOf Hotbar_MouseMove), GetAddress(AddressOf Hotbar_DblClick), False, False, GetAddress(AddressOf DrawHotbar)
End Sub

Public Sub CreateWindow_Bank()
    CreateWindow "winBank", "Bank", zOrder_Win, 0, 0, 391, 373, TextureItem(1), True, Fonts.verdana_13, , 2, 5, DesignTypes.DesignWindowClear, DesignTypes.DesignWindowClear, DesignTypes.DesignWindowClear, , , , , GetAddress(AddressOf Bank_MouseMove), GetAddress(AddressOf Bank_MouseDown), GetAddress(AddressOf Bank_MouseMove), 0, , , GetAddress(AddressOf DrawBank)

    ' Centralise it
    CentraliseWindow windowCount

    ' Set the index for spawning controls
    zOrder_Con = 1
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , , , , GetAddress(AddressOf btnMenu_Bank)

End Sub


Public Sub CreateWindow_Inventory()
' Create window
    CreateWindow "winInventory", "Inventory", zOrder_Win, 0, 0, 202, 319, TextureItem(1), False, Fonts.rockwellDec_15, , 2, 7, DesignTypes.DesignWindowClear, DesignTypes.DesignWindowClear, DesignTypes.DesignWindowClear, , , , , GetAddress(AddressOf Inventory_MouseMove), GetAddress(AddressOf Inventory_MouseDown), GetAddress(AddressOf Inventory_MouseMove), GetAddress(AddressOf Inventory_DblClick), , , GetAddress(AddressOf DrawInventory)

    ' Centralise it
    CentraliseWindow windowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , , , , GetAddress(AddressOf btnMenu_Inv)
    CreateLabel windowCount, "lblGold", 42, 296, 100, , "0g", verdana_12
End Sub

Public Sub CreateWindow_Character()
' Create window
    CreateWindow "winCharacter", "Character Status", zOrder_Win, 0, 0, 214, 318, TextureItem(62), False, Fonts.rockwellDec_15, , 2, 6, DesignTypes.DesignWindowClear, DesignTypes.DesignWindowClear, DesignTypes.DesignWindowClear, , , , , GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_MouseDown), GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_MouseMove), , , GetAddress(AddressOf DrawCharacter)

    ' Centralise it
    CentraliseWindow windowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , , , , GetAddress(AddressOf btnMenu_Char)

    ' Parchment
    CreatePictureBox windowCount, "picParchment", 6, 26, 162, 287, , , , , , , , DesignTypes.DesignParchment, DesignTypes.DesignParchment, DesignTypes.DesignParchment
    ' White boxes
    CreatePictureBox windowCount, "picWhiteBox", 13, 34, 148, 19, , , , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput
    CreatePictureBox windowCount, "picWhiteBox", 13, 54, 148, 19, , , , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput
    CreatePictureBox windowCount, "picWhiteBox", 13, 74, 148, 19, , , , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput
    CreatePictureBox windowCount, "picWhiteBox", 13, 94, 148, 19, , , , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput
    CreatePictureBox windowCount, "picWhiteBox", 13, 114, 148, 19, , , , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput
    CreatePictureBox windowCount, "picWhiteBox", 13, 134, 148, 19, , , , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput
    CreatePictureBox windowCount, "picWhiteBox", 13, 154, 148, 19, , , , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput
    ' Labels
    CreateLabel windowCount, "lblName", 18, 36, 147, 16, "Name", rockwellDec_10
    CreateLabel windowCount, "lblClass", 18, 56, 147, 16, "Class", rockwellDec_10
    CreateLabel windowCount, "lblLevel", 18, 76, 147, 16, "Level", rockwellDec_10
    CreateLabel windowCount, "lblGuild", 18, 96, 147, 16, "Guild", rockwellDec_10
    CreateLabel windowCount, "lblHealth", 18, 116, 147, 16, "Health", rockwellDec_10
    CreateLabel windowCount, "lblSpirit", 18, 136, 147, 16, "Spirit", rockwellDec_10
    CreateLabel windowCount, "lblExperience", 18, 156, 147, 16, "Experience", rockwellDec_10
    ' Attributes
    CreatePictureBox windowCount, "picShadow", 18, 176, 138, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblLabel", 18, 173, 138, , "Character Attributes", rockwellDec_15, , Alignment.AlignCentre
    ' Black boxes
    CreatePictureBox windowCount, "picBlackBox", 13, 186, 148, 19, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreatePictureBox windowCount, "picBlackBox", 13, 206, 148, 19, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreatePictureBox windowCount, "picBlackBox", 13, 226, 148, 19, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreatePictureBox windowCount, "picBlackBox", 13, 246, 148, 19, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreatePictureBox windowCount, "picBlackBox", 13, 266, 148, 19, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreatePictureBox windowCount, "picBlackBox", 13, 286, 148, 19, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    ' Labels
    CreateLabel windowCount, "lblLabel", 18, 188, 138, , "Strength", rockwellDec_10, Gold, Alignment.AlignRight
    CreateLabel windowCount, "lblLabel", 18, 208, 138, , "Endurance", rockwellDec_10, Gold, Alignment.AlignRight
    CreateLabel windowCount, "lblLabel", 18, 228, 138, , "Intelligence", rockwellDec_10, Gold, Alignment.AlignRight
    CreateLabel windowCount, "lblLabel", 18, 248, 138, , "Agility", rockwellDec_10, Gold, Alignment.AlignRight
    CreateLabel windowCount, "lblLabel", 18, 268, 138, , "Willpower", rockwellDec_10, Gold, Alignment.AlignRight
    CreateLabel windowCount, "lblLabel", 18, 288, 138, , "Unused Stat Points", rockwellDec_10, LightGreen, Alignment.AlignRight
    ' Buttons
    CreateButton windowCount, "btnStat_1", 15, 188, 15, 15, , , , , , , TextureGUI(18), TextureGUI(19), TextureGUI(20), , , , , , GetAddress(AddressOf Character_SpendPoint1)
    CreateButton windowCount, "btnStat_2", 15, 208, 15, 15, , , , , , , TextureGUI(18), TextureGUI(19), TextureGUI(20), , , , , , GetAddress(AddressOf Character_SpendPoint2)
    CreateButton windowCount, "btnStat_3", 15, 228, 15, 15, , , , , , , TextureGUI(18), TextureGUI(19), TextureGUI(20), , , , , , GetAddress(AddressOf Character_SpendPoint3)
    CreateButton windowCount, "btnStat_4", 15, 248, 15, 15, , , , , , , TextureGUI(18), TextureGUI(19), TextureGUI(20), , , , , , GetAddress(AddressOf Character_SpendPoint4)
    CreateButton windowCount, "btnStat_5", 15, 268, 15, 15, , , , , , , TextureGUI(18), TextureGUI(19), TextureGUI(20), , , , , , GetAddress(AddressOf Character_SpendPoint5)
    ' fake buttons
    CreatePictureBox windowCount, "btnGreyStat_1", 15, 188, 15, 15, , , , , TextureGUI(21), TextureGUI(21), TextureGUI(21)
    CreatePictureBox windowCount, "btnGreyStat_2", 15, 208, 15, 15, , , , , TextureGUI(21), TextureGUI(21), TextureGUI(21)
    CreatePictureBox windowCount, "btnGreyStat_3", 15, 228, 15, 15, , , , , TextureGUI(21), TextureGUI(21), TextureGUI(21)
    CreatePictureBox windowCount, "btnGreyStat_4", 15, 248, 15, 15, , , , , TextureGUI(21), TextureGUI(21), TextureGUI(21)
    CreatePictureBox windowCount, "btnGreyStat_5", 15, 268, 15, 15, , , , , TextureGUI(21), TextureGUI(21), TextureGUI(21)
    ' Labels
    CreateLabel windowCount, "lblStat_1", 32, 188, 100, , "255", rockwellDec_10
    CreateLabel windowCount, "lblStat_2", 32, 208, 100, , "255", rockwellDec_10
    CreateLabel windowCount, "lblStat_3", 32, 228, 100, , "255", rockwellDec_10
    CreateLabel windowCount, "lblStat_4", 32, 248, 100, , "255", rockwellDec_10
    CreateLabel windowCount, "lblStat_5", 32, 268, 100, , "255", rockwellDec_10
    CreateLabel windowCount, "lblPoints", 18, 288, 100, , "255", rockwellDec_10
End Sub

Public Sub CreateWindow_PlayerQuest()
' Create window
    CreateWindow "winPlayerQuests", "Quests", zOrder_Win, 0, 0, 450, 412, TextureItem(23), False, Fonts.rockwellDec_15, , 2, 7, DesignTypes.DesignWindowClear, DesignTypes.DesignWindowClear, DesignTypes.DesignWindowClear, , , , , GetAddress(AddressOf Inventory_MouseMove), GetAddress(AddressOf Inventory_MouseDown), GetAddress(AddressOf Inventory_MouseMove), GetAddress(AddressOf Inventory_DblClick), , , GetAddress(AddressOf DrawWinQuest)

    ' Centralise it
    CentraliseWindow windowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , , , , GetAddress(AddressOf btnMenu_Quest)

    CreateButton windowCount, "btnMission1", 5, 24, 125, 30, "Quest 1", verdana_12, White, , False, , , , , DesignTypes.DesignMenuHover, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , , GetAddress(AddressOf btnQuest1)
    CreateButton windowCount, "btnMission2", 5, 56, 125, 30, "Quest 2", verdana_12, White, , False, , , , , DesignTypes.DesignMenuHover, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , , GetAddress(AddressOf btnQuest2)
    CreateButton windowCount, "btnMission3", 5, 88, 125, 30, "Quest 3", verdana_12, White, , False, , , , , DesignTypes.DesignMenuHover, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , , GetAddress(AddressOf btnQuest3)
    CreateButton windowCount, "btnMission4", 5, 120, 125, 30, "Quest 4", verdana_12, White, , False, , , , , DesignTypes.DesignMenuHover, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , , GetAddress(AddressOf btnQuest4)
    CreateButton windowCount, "btnMission5", 5, 152, 125, 30, "Quest 5", verdana_12, White, , False, , , , , DesignTypes.DesignMenuHover, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , , GetAddress(AddressOf btnQuest5)
    CreateButton windowCount, "btnMission6", 5, 184, 125, 30, "Quest 6", verdana_12, White, , False, , , , , DesignTypes.DesignMenuHover, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , , GetAddress(AddressOf btnQuest6)
    CreateButton windowCount, "btnMission7", 5, 216, 125, 30, "Quest 7", verdana_12, White, , False, , , , , DesignTypes.DesignMenuHover, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , , GetAddress(AddressOf btnQuest7)
    CreateButton windowCount, "btnMission8", 5, 248, 125, 30, "Quest 8", verdana_12, White, , False, , , , , DesignTypes.DesignMenuHover, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , , GetAddress(AddressOf btnQuest8)
    CreateButton windowCount, "btnMission9", 5, 280, 125, 30, "Quest 9", verdana_12, White, , False, , , , , DesignTypes.DesignMenuHover, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , , GetAddress(AddressOf btnQuest9)
    CreateButton windowCount, "btnMission10", 5, 312, 125, 30, "Quest 10", verdana_12, White, , False, , , , , DesignTypes.DesignMenuHover, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , , GetAddress(AddressOf btnQuest10)
    CreateButton windowCount, "btnMission11", 5, 344, 125, 30, "Quest 11", verdana_12, White, , False, , , , , DesignTypes.DesignMenuHover, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , , GetAddress(AddressOf btnQuest11)
    CreateButton windowCount, "btnMission12", 5, 376, 125, 30, "Quest 12", verdana_12, White, , False, , , , , DesignTypes.DesignMenuHover, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , , GetAddress(AddressOf btnQuest12)

    ' Description
    CreateLabel windowCount, "lblDescrip", 142, 27, 150, 20, "Description:", Fonts.verdana_12, Grey, Alignment.AlignLeft, True
    CreatePictureBox windowCount, "picDescription", 137, 43, 307, 150, True, False, , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput
    CreateLabel windowCount, "lblDescription", 142, 47, 297, 140, "", Fonts.verdana_12, White, Alignment.AlignLeft, True

    ' Goal
    CreateLabel windowCount, "lblGo", 142, 197, 150, 20, "Goal:", Fonts.verdana_12, Grey, Alignment.AlignLeft, True
    CreatePictureBox windowCount, "picGoal", 137, 212, 307, 150, True, False, , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput
    CreateLabel windowCount, "lblGoal", 142, 217, 297, 140, "", Fonts.verdana_12, White, Alignment.AlignLeft, True

    CreateLabel windowCount, "lblExp", 335, 381, 101, , "1.000.000 EXP", rockwellDec_15, LightGreen, Alignment.AlignCentre
End Sub

Public Sub CreateWindow_Description()
' Create window
    CreateWindow "winDescription", "", zOrder_Win, 0, 0, 193, 142, 0, , , , , , DesignTypes.designWindowDescription, DesignTypes.designWindowDescription, DesignTypes.designWindowDescription, , , , , , , , , False

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Name
    CreateLabel windowCount, "lblName", 8, 12, 177, , "(SB) Flame Sword", rockwellDec_15, BrightBlue, Alignment.AlignCentre
    ' Sprite box
    CreatePictureBox windowCount, "picSprite", 18, 32, 68, 68, , , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenNormal, , , , , , GetAddress(AddressOf Description_OnDraw)
    ' Sep
    CreatePictureBox windowCount, "picSep", 96, 28, 1, 92, , , , , TextureGUI(44), TextureGUI(44), TextureGUI(44)
    ' Requirements
    CreateLabel windowCount, "lblClass", 5, 102, 92, , "Warrior", verdana_12, LightGreen, Alignment.AlignCentre
    CreateLabel windowCount, "lblLevel", 5, 114, 92, , "Level 20", verdana_12, BrightRed, Alignment.AlignCentre
    CreateLabel windowCount, "lblDescription", 100, 28, 85, 112, "Level 20", verdana_12, White, Alignment.AlignCentre, False
    ' Bar
    CreatePictureBox windowCount, "picBar", 19, 114, 66, 12, False, , , , TextureGUI(45), TextureGUI(45), TextureGUI(45)
End Sub

Public Sub CreateWindow_DragBox()
' Create window
    CreateWindow "winDragBox", "", zOrder_Win, 0, 0, 32, 32, 0, , , , , , , , , , , , GetAddress(AddressOf DragBox_Check), , , , , , , GetAddress(AddressOf DragBox_OnDraw)
    ' Need to set up unique mouseup event
    Windows(windowCount).Window.entCallBack(EntityStates.MouseUp) = GetAddress(AddressOf DragBox_Check)
End Sub

Public Sub CreateWindow_Skills()
' Create window
    CreateWindow "winSkills", "Skills", zOrder_Win, 0, 0, 202, 297, TextureItem(109), False, Fonts.rockwellDec_15, , 2, 7, DesignTypes.DesignWindowClear, DesignTypes.DesignWindowClear, DesignTypes.DesignWindowClear, , , , , GetAddress(AddressOf Skills_MouseMove), GetAddress(AddressOf Skills_MouseDown), GetAddress(AddressOf Skills_MouseMove), GetAddress(AddressOf Skills_DblClick), , , GetAddress(AddressOf DrawSkills)

    ' Centralise it
    CentraliseWindow windowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , , , , GetAddress(AddressOf btnMenu_Skills)
End Sub

Public Sub CreateWindow_Chat()
' Create window
    CreateWindow "winChat", "", zOrder_Win, 8, 422, 352, 152, 0, False, , , , , , , , , , , , , , , , False

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Channel boxes
    CreateCheckbox windowCount, "chkGame", 10, 2, 49, 23, 1, "Game", rockwellDec_10, , , , , DesignTypes.DesignCheckChat, , , GetAddress(AddressOf chkChat_Game)
    CreateCheckbox windowCount, "chkMap", 60, 2, 49, 23, 1, "Map", rockwellDec_10, , , , , DesignTypes.DesignCheckChat, , , GetAddress(AddressOf chkChat_Map)
    CreateCheckbox windowCount, "chkGlobal", 110, 2, 49, 23, 1, "Global", rockwellDec_10, , , , , DesignTypes.DesignCheckChat, , , GetAddress(AddressOf chkChat_Global)
    CreateCheckbox windowCount, "chkParty", 160, 2, 49, 23, 1, "Party", rockwellDec_10, , , , , DesignTypes.DesignCheckChat, , , GetAddress(AddressOf chkChat_Party)
    CreateCheckbox windowCount, "chkGuild", 210, 2, 49, 23, 1, "Guild", rockwellDec_10, , , , , DesignTypes.DesignCheckChat, , , GetAddress(AddressOf chkChat_Guild)
    CreateCheckbox windowCount, "chkPrivate", 260, 2, 49, 23, 1, "Private", rockwellDec_10, , , , , DesignTypes.DesignCheckChat, , , GetAddress(AddressOf chkChat_Private)

    ' Blank picturebox - ondraw wrapper
    CreatePictureBox windowCount, "picNull", 0, 0, 0, 0, , , , , , , , , , , , , , , , GetAddress(AddressOf OnDraw_Chat)
    ' Chat button
    CreateButton windowCount, "btnChat", 296, 124 + 16, 48, 20, "Say", rockwellDec_15, , , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnSay_Click)
    ' Chat Textbox
    CreateTextbox windowCount, "txtChat", 12, 127 + 16, 286, 25, , Fonts.verdana_12
    ' buttons
    CreateButton windowCount, "btnUp", 328, 28, 11, 13, , , , , , , TextureGUI(6), TextureGUI(7), TextureGUI(8), , , , , , GetAddress(AddressOf btnChat_Up)
    CreateButton windowCount, "btnDown", 327, 122, 11, 13, , , , , , , TextureGUI(9), TextureGUI(10), TextureGUI(11), , , , , , GetAddress(AddressOf btnChat_Down)

    ' Custom Handlers for mouse up
    Windows(windowCount).Controls(GetControlIndex("winChat", "btnUp")).entCallBack(EntityStates.MouseUp) = GetAddress(AddressOf btnChat_Up_MouseUp)
    Windows(windowCount).Controls(GetControlIndex("winChat", "btnDown")).entCallBack(EntityStates.MouseUp) = GetAddress(AddressOf btnChat_Down_MouseUp)

    ' Set the active control
    SetActiveControl GetWindowIndex("winChat"), GetControlIndex("winChat", "txtChat")

    ' sort out the tabs
    With Windows(GetWindowIndex("winChat"))
        .Controls(GetControlIndex("winChat", "chkGame")).Value = Options.channelState(ChatChannel.chGame)
        .Controls(GetControlIndex("winChat", "chkMap")).Value = Options.channelState(ChatChannel.chMap)
        .Controls(GetControlIndex("winChat", "chkGlobal")).Value = Options.channelState(ChatChannel.chGlobal)
        .Controls(GetControlIndex("winChat", "chkParty")).Value = Options.channelState(ChatChannel.chParty)
        .Controls(GetControlIndex("winChat", "chkGuild")).Value = Options.channelState(ChatChannel.chGuild)
        .Controls(GetControlIndex("winChat", "chkPrivate")).Value = Options.channelState(ChatChannel.chPrivate)
    End With
End Sub

Public Sub CreateWindow_ChatSmall()
' Create window
    CreateWindow "winChatSmall", "", zOrder_Win, 8, 438, 0, 0, 0, False, , , , , , , , , , , , , , , , False, , GetAddress(AddressOf OnDraw_ChatSmall), , True

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Chat Label
    CreateLabel windowCount, "lblMsg", 12, 127, 286, 25, "Press 'Enter' to open chatbox.", verdana_12, Grey
End Sub

Public Sub CreateWindow_Options()
' Create window
    CreateWindow "winOptions", "", zOrder_Win, 0, 0, 210, 212, 0, , , , , , DesignTypes.DesignWindowWithoutBar, DesignTypes.DesignWindowWithoutBar, DesignTypes.DesignWindowWithoutBar, , , , , , , , , False, False

    ' Centralise it
    CentraliseWindow windowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Parchment
    CreatePictureBox windowCount, "picParchment", 6, 6, 198, 200, , , , , , , , DesignTypes.DesignParchment, DesignTypes.DesignParchment, DesignTypes.DesignParchment
    ' General
    CreatePictureBox windowCount, "picBlank", 35, 25, 140, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblBlank", 35, 22, 140, , "General Options", rockwellDec_15, White, Alignment.AlignCentre
    ' Check boxes
    CreateCheckbox windowCount, "chkMusic", 35, 40, 80, , , "Music", rockwellDec_10, , , , , DesignTypes.DesignCheckbox
    CreateCheckbox windowCount, "chkSound", 115, 40, 80, , , "Sound", rockwellDec_10, , , , , DesignTypes.DesignCheckbox
    CreateCheckbox windowCount, "chkAutotiles", 35, 60, 80, , , "Autotiles", rockwellDec_10, , , , , DesignTypes.DesignCheckbox
    CreateCheckbox windowCount, "chkFullscreen", 115, 60, 80, , , "Fullscreen", rockwellDec_10, , , , , DesignTypes.DesignCheckbox

    ' Resolution
    CreatePictureBox windowCount, "picBlank", 35, 85, 140, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblBlank", 35, 82, 140, , "Select Resolution", rockwellDec_15, White, Alignment.AlignCentre
    ' combobox
    CreateComboBox windowCount, "cmbRes", 30, 100, 150, 18, DesignTypes.DesignCombo, verdana_12
    ' Renderer
    CreatePictureBox windowCount, "picBlank", 35, 125, 140, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblBlank", 35, 122, 140, , "DirectX Mode", rockwellDec_15, White, Alignment.AlignCentre
    ' Check boxes
    CreateComboBox windowCount, "cmbRender", 30, 140, 150, 18, DesignTypes.DesignCombo, verdana_12
    ' Button
    CreateButton windowCount, "btnConfirm", 65, 168, 80, 22, "Confirm", rockwellDec_15, , , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnOptions_Confirm)

    ' Populate the options screen
    SetOptionsScreen
End Sub

Public Sub CreateWindow_Shop()
' Create window
    CreateWindow "winShop", "Shop", zOrder_Win, 0, 0, 278, 293, TextureItem(17), False, Fonts.rockwellDec_15, , 2, 5, DesignTypes.DesignWindowClear, DesignTypes.DesignWindowClear, DesignTypes.DesignWindowClear, , , , , GetAddress(AddressOf Shop_MouseMove), GetAddress(AddressOf Shop_MouseDown), GetAddress(AddressOf Shop_MouseMove), GetAddress(AddressOf Shop_MouseMove), , , GetAddress(AddressOf DrawShopBackground)

    ' additional mouse event
    Windows(windowCount).Window.entCallBack(EntityStates.MouseUp) = GetAddress(AddressOf Shop_MouseMove)
    ' Centralise it
    CentraliseWindow windowCount

    ' Close button
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , , , , GetAddress(AddressOf btnShop_Close)
    ' Parchment
    CreatePictureBox windowCount, "picParchment", 6, 215, 266, 50, , , , , , , , DesignTypes.DesignParchment, DesignTypes.DesignParchment, DesignTypes.DesignParchment, , , , , , GetAddress(AddressOf DrawShop)
    ' Picture Box
    CreatePictureBox windowCount, "picItemBG", 13, 222, 36, 36, , , , , TextureGUI(30), TextureGUI(30), TextureGUI(30)
    CreatePictureBox windowCount, "picItem", 15, 224, 32, 32
    ' Buttons
    CreateButton windowCount, "btnBuy", 190, 228, 70, 24, "Buy", rockwellDec_15, White, , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnShopBuy)
    CreateButton windowCount, "btnSell", 190, 228, 70, 24, "Sell", rockwellDec_15, White, , False, , , , , DesignTypes.DesignRedNormal, DesignTypes.DesignRedHover, DesignTypes.DesignRedClick, , , GetAddress(AddressOf btnShopSell)
    ' Buying/Selling
    CreateCheckbox windowCount, "chkBuying", 173, 265, 49, 20, 1, , , , , , , DesignTypes.DesignCheckBuy, , , GetAddress(AddressOf chkShopBuying)
    CreateCheckbox windowCount, "chkSelling", 222, 265, 49, 20, 0, , , , , , , DesignTypes.DesignCheckSell, , , GetAddress(AddressOf chkShopSelling)

    ' Labels
    CreateLabel windowCount, "lblName", 56, 226, 300, , "Test Item", verdanaBold_12, Black, Alignment.AlignLeft
    CreateLabel windowCount, "lblCost", 56, 240, 300, , "1000g", verdana_12, Black, Alignment.AlignLeft
    ' Gold
    CreateLabel windowCount, "lblGold", 44, 269, 300, , "0g", verdana_12
End Sub

Public Sub CreateWindow_Offer()
    Dim WidthWindow As Long, HeightWindow As Long
    Dim Yo As Long, Xo As Long
    ' Create window
    CreateWindow "winOffer", "", zOrder_Win, 10, 90, 535, 285, TextureItem(111), False, Fonts.rockwellDec_15, , 2, 11, , , , , , , , GetAddress(AddressOf Offer_MouseMove), , GetAddress(AddressOf Offer_MouseMove), , False, , GetAddress(AddressOf DrawInviteBackground)

    CreatePictureBox windowCount, "picBGOffer1", 0, 0, 485, 45, False, , , , , , , DesignTypes.designWindowDescription, DesignTypes.designWindowDescription, DesignTypes.designWindowDescription
    WidthWindow = Windows(windowCount).Controls(GetControlIndex("winOffer", "picBGOffer1")).Width
    HeightWindow = Windows(windowCount).Controls(GetControlIndex("winOffer", "picBGOffer1")).Height - 18
    Yo = Windows(windowCount).Controls(GetControlIndex("winOffer", "picBGOffer1")).Top + 10
    Xo = Windows(windowCount).Controls(GetControlIndex("winOffer", "picBGOffer1")).Left
    ' Offer BG
    CreatePictureBox windowCount, "picOfferBG1", 10, Yo, 334, 25, False, , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    ' Title Offer
    CreateLabel windowCount, "lblTitleOffer1", 7 + Xo + ((334 - 318) / 2), Yo + 5, 318, 25, "[Offer]", Fonts.georgia_16, White, Alignment.AlignLeft, False
    ' Action buttons
    CreateButton windowCount, "btnAccept1", 349, Yo, 60, 25, "Accept", verdana_12, Grey, , False, , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf AcceptOffer1), , , , , DarkGrey
    CreateButton windowCount, "btnRecuse1", 414, Yo, 60, 25, "Refuse", verdana_12, Grey, , False, , , , , DesignTypes.DesignRedNormal, DesignTypes.DesignRedHover, DesignTypes.DesignRedClick, , , GetAddress(AddressOf RecuseOffer1), , , , , DarkGrey
    ' Offer BG#################################################################################
    CreatePictureBox windowCount, "picBGOffer2", 0, Yo + HeightWindow, 485, 45, False, , , , , , , DesignTypes.designWindowDescription, DesignTypes.designWindowDescription, DesignTypes.designWindowDescription
    Yo = Windows(windowCount).Controls(GetControlIndex("winOffer", "picBGOffer2")).Top + 10
    Xo = Windows(windowCount).Controls(GetControlIndex("winOffer", "picBGOffer2")).Left
    CreatePictureBox windowCount, "picOfferBG2", 10, Yo, 334, 25, False, , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    ' Title Offer
    CreateLabel windowCount, "lblTitleOffer2", 7 + Xo + ((334 - 318) / 2), Yo + 5, 318, 25, "[Offer]", Fonts.georgia_16, White, Alignment.AlignLeft, False
    ' Action buttons
    CreateButton windowCount, "btnAccept2", 349, Yo, 60, 25, "Accept", verdana_12, Grey, , False, , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf AcceptOffer2), , , , , DarkGrey
    CreateButton windowCount, "btnRecuse2", 414, Yo, 60, 25, "Refuse", verdana_12, Grey, , False, , , , , DesignTypes.DesignRedNormal, DesignTypes.DesignRedHover, DesignTypes.DesignRedClick, , , GetAddress(AddressOf RecuseOffer2), , , , , DarkGrey
    ' Offer BG#################################################################################
    CreatePictureBox windowCount, "picBGOffer3", 0, Yo + HeightWindow, 485, 45, False, , , , , , , DesignTypes.designWindowDescription, DesignTypes.designWindowDescription, DesignTypes.designWindowDescription
    Yo = Windows(windowCount).Controls(GetControlIndex("winOffer", "picBGOffer3")).Top + 10
    Xo = Windows(windowCount).Controls(GetControlIndex("winOffer", "picBGOffer3")).Left
    CreatePictureBox windowCount, "picOfferBG3", 10, Yo, 334, 25, False, , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    ' Title Offer
    CreateLabel windowCount, "lblTitleOffer3", 7 + Xo + ((334 - 318) / 2), Yo + 5, 318, 25, "[Offer]", Fonts.georgia_16, White, Alignment.AlignLeft, False
    ' Action buttons
    CreateButton windowCount, "btnAccept3", 349, Yo, 60, 25, "Accept", verdana_12, Grey, , False, , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf AcceptOffer3), , , , , DarkGrey
    CreateButton windowCount, "btnRecuse3", 414, Yo, 60, 25, "Refuse", verdana_12, Grey, , False, , , , , DesignTypes.DesignRedNormal, DesignTypes.DesignRedHover, DesignTypes.DesignRedClick, , , GetAddress(AddressOf RecuseOffer3), , , , , DarkGrey
End Sub

Public Sub CreateWindow_NpcChat()
' Create window

    CreateWindow "winNpcChat", "Conversation with [Name]", zOrder_Win, 0, 0, 480, 228, TextureItem(111), False, Fonts.rockwellDec_15, , 2, 11, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal

    ' Centralise it
    CentraliseWindow windowCount

    ' Close Button
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , , , , GetAddress(AddressOf btnNpcChat_Close)

    ' Parchment
    CreatePictureBox windowCount, "picParchment", 6, 26, 468, 198, , , , , , , , DesignTypes.DesignParchment, DesignTypes.DesignParchment, DesignTypes.DesignParchment
    ' Face background
    CreatePictureBox windowCount, "picFaceBG", 20, 40, 102, 102, , , , , TextureGUI(36), TextureGUI(36), TextureGUI(36)
    ' Actual Face
    CreatePictureBox windowCount, "picFace", 23, 43, 96, 96, , , , , TextureFace(1), TextureFace(1), TextureFace(1)
    ' Chat BG
    CreatePictureBox windowCount, "picChatBG", 128, 39, 334, 104, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    ' Chat
    CreateLabel windowCount, "lblChat", 136, 44, 318, 102, "[Text]", rockwellDec_15, White, Alignment.AlignCentre
    ' Reply buttons
    CreateButton windowCount, "btnOpt4", 69, 145, 343, 15, "[Text]", verdana_12, Black, , , , , , , , , , , , GetAddress(AddressOf btnOpt4), , , , , DarkGrey
    CreateButton windowCount, "btnOpt3", 69, 162, 343, 15, "[Text]", verdana_12, Black, , , , , , , , , , , , GetAddress(AddressOf btnOpt3), , , , , DarkGrey
    CreateButton windowCount, "btnOpt2", 69, 179, 343, 15, "[Text]", verdana_12, Black, , , , , , , , , , , , GetAddress(AddressOf btnOpt2), , , , , DarkGrey
    CreateButton windowCount, "btnOpt1", 69, 196, 343, 15, "[Text]", verdana_12, Black, , , , , , , , , , , , GetAddress(AddressOf btnOpt1), , , , , DarkGrey

    ' Cache positions
    optPos(1) = 196
    optPos(2) = 179
    optPos(3) = 162
    optPos(4) = 145
    optHeight = 228
End Sub

Public Sub CreateWindow_RightClick()
' Create window
    CreateWindow "winRightClickBG", "", zOrder_Win, 0, 0, 800, 600, 0, , , , , , , , , , , , , , GetAddress(AddressOf RightClick_Close), , , False
    ' Centralise it
    CentraliseWindow windowCount
End Sub

Public Sub CreateWindow_PlayerMenu()
' Create window
    CreateWindow "winPlayerMenu", "", zOrder_Win, 0, 0, 110, 106, 0, , , , , , DesignTypes.designWindowDescription, DesignTypes.designWindowDescription, DesignTypes.designWindowDescription, , , , , , GetAddress(AddressOf RightClick_Close), , , False
    ' Centralise it
    CentraliseWindow windowCount

    ' Name
    CreateButton windowCount, "btnName", 8, 8, 94, 18, "[Name]", verdanaBold_12, White, , , , , , , DesignTypes.DesignMenuHeader, DesignTypes.DesignMenuHeader, DesignTypes.DesignMenuHeader, , , GetAddress(AddressOf RightClick_Close)
    ' Options
    CreateButton windowCount, "btnParty", 8, 26, 94, 18, "Invite to Party", verdana_12, White, , , , , , , , DesignTypes.DesignMenuHover, , , , GetAddress(AddressOf PlayerMenu_Party)
    CreateButton windowCount, "btnTrade", 8, 44, 94, 18, "Request Trade", verdana_12, White, , , , , , , , DesignTypes.DesignMenuHover, , , , GetAddress(AddressOf PlayerMenu_Trade)
    CreateButton windowCount, "btnGuild", 8, 62, 94, 18, "Invite to Guild", verdana_12, White, , , , , , , , DesignTypes.DesignMenuHover, , , , GetAddress(AddressOf PlayerMenu_Guild)
    CreateButton windowCount, "btnPM", 8, 80, 94, 18, "Private Message", verdana_12, White, , , , , , , , DesignTypes.DesignMenuHover, , , , GetAddress(AddressOf PlayerMenu_PM)

End Sub

Public Sub CreateWindow_Party()
' Create window
    CreateWindow "winParty", "", zOrder_Win, 4, 78, 252, 158, 0, , , , , , DesignTypes.designWindowDescription, DesignTypes.designWindowDescription, DesignTypes.designWindowDescription, , , , , , , , , False

    ' Name labels
    CreateLabel windowCount, "lblName1", 60, 20, 173, , "Richard - Level 10", rockwellDec_10
    CreateLabel windowCount, "lblName2", 60, 60, 173, , "Anna - Level 18", rockwellDec_10
    CreateLabel windowCount, "lblName3", 60, 100, 173, , "Doleo - Level 25", rockwellDec_10
    ' Empty Bars - HP
    CreatePictureBox windowCount, "picEmptyBar_HP1", 58, 34, 173, 9, , , , , TextureGUI(22), TextureGUI(22), TextureGUI(22)
    CreatePictureBox windowCount, "picEmptyBar_HP2", 58, 74, 173, 9, , , , , TextureGUI(22), TextureGUI(22), TextureGUI(22)
    CreatePictureBox windowCount, "picEmptyBar_HP3", 58, 114, 173, 9, , , , , TextureGUI(22), TextureGUI(22), TextureGUI(22)
    ' Empty Bars - SP
    CreatePictureBox windowCount, "picEmptyBar_SP1", 58, 44, 173, 9, , , , , TextureGUI(23), TextureGUI(23), TextureGUI(23)
    CreatePictureBox windowCount, "picEmptyBar_SP2", 58, 84, 173, 9, , , , , TextureGUI(23), TextureGUI(23), TextureGUI(23)
    CreatePictureBox windowCount, "picEmptyBar_SP3", 58, 124, 173, 9, , , , , TextureGUI(23), TextureGUI(23), TextureGUI(23)
    ' Filled bars - HP
    CreatePictureBox windowCount, "picBar_HP1", 58, 34, 173, 9, , , , , TextureGUI(24), TextureGUI(24), TextureGUI(24)
    CreatePictureBox windowCount, "picBar_HP2", 58, 74, 173, 9, , , , , TextureGUI(24), TextureGUI(24), TextureGUI(24)
    CreatePictureBox windowCount, "picBar_HP3", 58, 114, 173, 9, , , , , TextureGUI(24), TextureGUI(24), TextureGUI(24)
    ' Filled bars - SP
    CreatePictureBox windowCount, "picBar_SP1", 58, 44, 173, 9, , , , , TextureGUI(25), TextureGUI(25), TextureGUI(25)
    CreatePictureBox windowCount, "picBar_SP2", 58, 84, 173, 9, , , , , TextureGUI(25), TextureGUI(25), TextureGUI(25)
    CreatePictureBox windowCount, "picBar_SP3", 58, 124, 173, 9, , , , , TextureGUI(25), TextureGUI(25), TextureGUI(25)
    ' Shadows
    CreatePictureBox windowCount, "picShadow1", 20, 24, 32, 32, , , , , TextureShadow, TextureShadow, TextureShadow
    CreatePictureBox windowCount, "picShadow2", 20, 64, 32, 32, , , , , TextureShadow, TextureShadow, TextureShadow
    CreatePictureBox windowCount, "picShadow3", 20, 104, 32, 32, , , , , TextureShadow, TextureShadow, TextureShadow
    ' Characters
    CreatePictureBox windowCount, "picChar1", 20, 20, 32, 32, , , , , TextureChar(1), TextureChar(1), TextureChar(1)
    CreatePictureBox windowCount, "picChar2", 20, 60, 32, 32, , , , , TextureChar(1), TextureChar(1), TextureChar(1)
    CreatePictureBox windowCount, "picChar3", 20, 100, 32, 32, , , , , TextureChar(1), TextureChar(1), TextureChar(1)
End Sub

Public Sub CreateWindow_Trade()
' Create window
    CreateWindow "winTrade", "Trading with [Name]", zOrder_Win, 0, 0, 412, 386, TextureItem(112), False, Fonts.rockwellDec_15, , 2, 5, DesignTypes.DesignWindowClear, DesignTypes.DesignWindowClear, DesignTypes.DesignWindowClear, , , , , , , , , , , GetAddress(AddressOf DrawTrade)

    ' Centralise it
    CentraliseWindow windowCount

    ' Close Button
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , , , , GetAddress(AddressOf btnTrade_Close)
    ' Parchment
    CreatePictureBox windowCount, "picParchment", 10, 312, 392, 66, , , , , , , , DesignTypes.DesignParchment, DesignTypes.DesignParchment, DesignTypes.DesignParchment
    ' Labels
    CreatePictureBox windowCount, "picShadow", 36, 30, 142, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblYourTrade", 36, 27, 142, 9, "Robin's Offer", rockwellDec_15, White, Alignment.AlignCentre
    CreatePictureBox windowCount, "picShadow", 36 + 200, 30, 142, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblTheirTrade", 36 + 200, 27, 142, 9, "Richard's Offer", rockwellDec_15, White, Alignment.AlignCentre
    ' Buttons
    CreateButton windowCount, "btnAccept", 134, 340, 68, 24, "Accept", rockwellDec_15, White, , , , , , , DesignTypes.DesignGreenNormal, DesignTypes.DesignGreenHover, DesignTypes.DesignGreenClick, , , GetAddress(AddressOf btnTrade_Accept)
    CreateButton windowCount, "btnDecline", 210, 340, 68, 24, "Decline", rockwellDec_15, White, , , , , , , DesignTypes.DesignRedNormal, DesignTypes.DesignRedHover, DesignTypes.DesignRedClick, , , GetAddress(AddressOf btnTrade_Close)
    ' Labels
    CreateLabel windowCount, "lblStatus", 114, 322, 184, , "", verdanaBold_12, Red, Alignment.AlignCentre
    ' Amounts
    CreateLabel windowCount, "lblBlank", 25, 330, 100, , "Total Value", verdanaBold_12, Black, Alignment.AlignCentre
    CreateLabel windowCount, "lblBlank", 285, 330, 100, , "Total Value", verdanaBold_12, Black, Alignment.AlignCentre
    CreateLabel windowCount, "lblYourValue", 25, 344, 100, , "52,812g", verdana_12, Black, Alignment.AlignCentre
    CreateLabel windowCount, "lblTheirValue", 285, 344, 100, , "12,531g", verdana_12, Black, Alignment.AlignCentre
    ' Item Containers
    CreatePictureBox windowCount, "picYour", 14, 46, 184, 260, , , , , , , , , , , , GetAddress(AddressOf TradeMouseMove_Your), GetAddress(AddressOf TradeMouseDown_Your), GetAddress(AddressOf TradeMouseMove_Your), , GetAddress(AddressOf DrawYourTrade)
    CreatePictureBox windowCount, "picTheir", 214, 46, 184, 260, , , , , , , , , , , , GetAddress(AddressOf TradeMouseMove_Their), GetAddress(AddressOf TradeMouseMove_Their), GetAddress(AddressOf TradeMouseMove_Their), , GetAddress(AddressOf DrawTheirTrade)
End Sub

Public Sub CreateWindow_Combobox()
' background window
    CreateWindow "winComboMenuBG", "ComboMenuBG", zOrder_Win, 0, 0, 800, 600, 0, , , , , , , , , , , , , , GetAddress(AddressOf CloseComboMenu), , , False, False

    ' window
    CreateWindow "winComboMenu", "ComboMenu", zOrder_Win, 0, 0, 100, 100, 0, , Fonts.verdana_12, , , , DesignTypes.DesignComboBackground, , , , , , , , , , , False, False

    ' centralise it
    CentraliseWindow windowCount
End Sub

Public Sub CreateWindow_Guild()
' Create window

    CreateWindow "winGuild", "Guild", zOrder_Win, 0, 0, 174, 320, TextureItem(107), False, Fonts.rockwellDec_15, , 2, 6, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal, DesignTypes.DesignWindowNormal

    ' Centralise it
    CentraliseWindow windowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , , , , GetAddress(AddressOf btnMenu_Guild)
    ' Parchment
    CreatePictureBox windowCount, "picParchment", 6, 26, 162, 287, , , , , , , , DesignTypes.DesignParchment, DesignTypes.DesignParchment, DesignTypes.DesignParchment
    ' Attributes
    CreatePictureBox windowCount, "picShadow", 18, 38, 138, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblGuild", 18, 35, 138, , "Guild Name", rockwellDec_15, , Alignment.AlignCentre
    ' White boxes
    CreatePictureBox windowCount, "picWhiteBox", 13, 51, 148, 19, , , , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput
    CreatePictureBox windowCount, "picWhiteBox", 13, 71, 148, 19, , , , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput
    CreatePictureBox windowCount, "picWhiteBox", 13, 91, 148, 19, , , , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput
    CreatePictureBox windowCount, "picWhiteBox", 13, 111, 148, 19, , , , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput
    ' Labels
    CreateLabel windowCount, "lblRank", 18, 53, 147, 16, "Guild Rank: None", rockwellDec_10
    CreateLabel windowCount, "lblKills", 18, 73, 147, 16, "Enemy Kills: 0", rockwellDec_10
    CreateLabel windowCount, "lblGold", 18, 93, 147, 16, "Bank Gold: 0g", rockwellDec_10
    CreateLabel windowCount, "lblMembers", 18, 113, 147, 16, "Guild Members: 0", rockwellDec_10
End Sub

' Rendering & Initialisation
Public Sub InitGUI()

' Starter values
    zOrder_Win = 1
    zOrder_Con = 1

    ' Menu
    CreateWindow_Login
    CreateWindow_Characters
    CreateWindow_Loading
    CreateWindow_Dialogue
    CreateWindow_Classes
    CreateWindow_NewChar
    CreateWindow_Register

    ' Game
    CreateWindow_Combobox
    CreateWindow_EscMenu
    CreateWindow_Bars
    CreateWindow_Bank
    CreateWindow_Menu
    CreateWindow_Hotbar
    CreateWindow_Inventory
    CreateWindow_Character
    CreateWindow_PlayerQuest
    CreateWindow_Description
    CreateWindow_DragBox
    CreateWindow_Skills
    CreateWindow_Chat
    CreateWindow_ChatSmall
    CreateWindow_Options
    CreateWindow_Shop
    CreateWindow_NpcChat
    CreateWindow_Offer
    CreateWindow_Party
    CreateWindow_Trade
    CreateWindow_Guild

    ' Menus
    CreateWindow_RightClick
    CreateWindow_PlayerMenu
End Sub
