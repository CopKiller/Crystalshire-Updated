Attribute VB_Name = "modText"
Option Explicit

'The size of a FVF vertex
Public Const FVF_Size As Long = 28

'Point API
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type CharVA
    Vertex(0 To 3) As Vertex
End Type

Private Type VFH
    BitmapWidth As Long
    BitmapHeight As Long
    CellWidth As Long
    CellHeight As Long
    BaseCharOffset As Byte
    CharWidth(0 To 255) As Byte
    CharVA(0 To 255) As CharVA
End Type

Private Type CustomFont
    HeaderInfo As VFH
    Texture As Direct3DTexture8
    RowPitch As Integer
    RowFactor As Single
    ColFactor As Single
    CharHeight As Byte
    TextureSize As POINTAPI
    xOffset As Long
    yOffset As Long
End Type

' Fonts
Public Enum Fonts
    ' Georgia
    georgia_16 = 1
    georgiaBold_16
    georgiaDec_16
    ' Rockwell
    rockwellDec_15
    rockwell_15
    rockwellDec_10
    ' Verdana
    verdana_12
    verdanaBold_12
    verdana_13
    
    Default
    ' count value
    Fonts_Count
End Enum

' Store the fonts
Public font() As CustomFont

' Chatbox
Public Type ChatStruct
    text As String
    Color As Long
    visible As Boolean
    timer As Long
    Channel As Byte
End Type
Public Const ColourChar As String * 1 = "½"
Public Const ChatLines As Long = 200
Public Const ChatWidth As Long = 316
Public Chat(1 To ChatLines) As ChatStruct
Public chatLastRemove As Long
Public Const CHAT_DIFFERENCE_TIMER As Long = 500
Public Chat_HighIndex As Long
Public ChatScroll As Long

Sub LoadFonts()
    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
    ' re-dim the fonts
    ReDim font(1 To Fonts.Fonts_Count - 1)
    ' load the fonts
    SetFont Fonts.georgia_16, "georgia_16", 256
    SetFont Fonts.georgiaBold_16, "georgiaBold_16", 256
    SetFont Fonts.georgiaDec_16, "georgiaDec_16", 256
    SetFont Fonts.rockwellDec_15, "rockwellDec_15", 256, 2, 2
    SetFont Fonts.rockwell_15, "rockwell_15", 256, 2, 2
    SetFont Fonts.verdana_12, "verdana_12", 256
    SetFont Fonts.verdanaBold_12, "verdanaBold_12", 256
    SetFont Fonts.rockwellDec_10, "rockwellDec_10", 256, 2, 2
    SetFont Fonts.Default, "default", 256
End Sub

Sub SetFont(ByVal fontNum As Long, ByVal texName As String, ByVal Size As Long, Optional ByVal xOffset As Long, Optional ByVal yOffset As Long)
Dim Data() As Byte, f As Long, w As Long, h As Long, Path As String
    ' set the path
    Path = App.Path & PathFont & texName & GFX_EXT
    ' load the texture
    f = FreeFile
    Open Path For Binary As #f
        ReDim Data(0 To LOF(f) - 1)
        Get #f, , Data
    Close #f
    ' get size
    font(fontNum).TextureSize.X = ByteToInt(Data(18), Data(19))
    font(fontNum).TextureSize.Y = ByteToInt(Data(22), Data(23))
    ' set to struct
    Set font(fontNum).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Data(0), AryCount(Data), font(fontNum).TextureSize.X, font(fontNum).TextureSize.Y, D3DX_DEFAULT, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    font(fontNum).xOffset = xOffset
    font(fontNum).yOffset = yOffset
    LoadFontHeader font(fontNum), texName & ".dat"
End Sub

Public Function GetColourString(ByVal colourNum As Long) As String
    Select Case colourNum
        Case 0 ' Black
            GetColourString = "Black"
        Case 1 ' Blue
            GetColourString = "Blue"
        Case 2 ' Green
            GetColourString = "Green"
        Case 3 ' Cyan
            GetColourString = "Cyan"
        Case 4 ' Red
            GetColourString = "Red"
        Case 5 ' Magenta
            GetColourString = "Magenta"
        Case 6 ' Brown
            GetColourString = "Brown"
        Case 7 ' Grey
            GetColourString = "Grey"
        Case 8 ' DarkGrey
            GetColourString = "Dark Grey"
        Case 9 ' BrightBlue
            GetColourString = "Bright Blue"
        Case 10 ' BrightGreen
            GetColourString = "Bright Green"
        Case 11 ' BrightCyan
            GetColourString = "Bright Cyan"
        Case 12 ' BrightRed
            GetColourString = "Bright Red"
        Case 13 ' Pink
            GetColourString = "Pink"
        Case 14 ' Yellow
            GetColourString = "Yellow"
        Case 15 ' White
            GetColourString = "White"
        Case 16 ' dark brown
            GetColourString = "Dark Brown"
        Case 17 ' gold
            GetColourString = "Gold"
        Case 18 ' light green
            GetColourString = "Light Green"
    End Select
End Function

Public Function DX8Colour(ByVal colourNum As Long, ByVal alpha As Long) As Long
    Select Case colourNum
        Case 0 ' Black
            DX8Colour = D3DColorARGB(alpha, 0, 0, 0)
        Case 1 ' Blue
            DX8Colour = D3DColorARGB(alpha, 16, 104, 237)
        Case 2 ' Green
            DX8Colour = D3DColorARGB(alpha, 119, 188, 84)
        Case 3 ' Cyan
            DX8Colour = D3DColorARGB(alpha, 16, 224, 237)
        Case 4 ' Red
            DX8Colour = D3DColorARGB(alpha, 201, 0, 0)
        Case 5 ' Magenta
            DX8Colour = D3DColorARGB(alpha, 255, 0, 255)
        Case 6 ' Brown
            DX8Colour = D3DColorARGB(alpha, 175, 149, 92)
        Case 7 ' Grey
            DX8Colour = D3DColorARGB(alpha, 192, 192, 192)
        Case 8 ' DarkGrey
            DX8Colour = D3DColorARGB(alpha, 82, 82, 82)
        Case 9 ' BrightBlue
            DX8Colour = D3DColorARGB(alpha, 126, 182, 240)
        Case 10 ' BrightGreen
            DX8Colour = D3DColorARGB(alpha, 126, 240, 137)
        Case 11 ' BrightCyan
            DX8Colour = D3DColorARGB(alpha, 157, 242, 242)
        Case 12 ' BrightRed
            DX8Colour = D3DColorARGB(alpha, 255, 0, 0)
        Case 13 ' Pink
            DX8Colour = D3DColorARGB(alpha, 255, 118, 221)
        Case 14 ' Yellow
            DX8Colour = D3DColorARGB(alpha, 255, 255, 0)
        Case 15 ' White
            DX8Colour = D3DColorARGB(alpha, 255, 255, 255)
        Case 16 ' dark brown
            DX8Colour = D3DColorARGB(alpha, 98, 84, 52)
        Case 17 ' gold
            DX8Colour = D3DColorARGB(alpha, 255, 215, 0)
        Case 18 ' light green
            DX8Colour = D3DColorARGB(alpha, 124, 205, 80)
    End Select
End Function

Sub LoadFontHeader(ByRef theFont As CustomFont, ByVal FileName As String)
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single

    'Load the header information
    FileNum = FreeFile
    Open App.Path & PathFont & FileName For Binary As #FileNum
    Get #FileNum, , theFont.HeaderInfo
    Close #FileNum
    
    'Calculate some common values
    theFont.CharHeight = theFont.HeaderInfo.CellHeight - 4
    theFont.RowPitch = theFont.HeaderInfo.BitmapWidth \ theFont.HeaderInfo.CellWidth
    theFont.ColFactor = theFont.HeaderInfo.CellWidth / theFont.HeaderInfo.BitmapWidth
    theFont.RowFactor = theFont.HeaderInfo.CellHeight / theFont.HeaderInfo.BitmapHeight

    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - theFont.HeaderInfo.BaseCharOffset) \ theFont.RowPitch
        u = ((LoopChar - theFont.HeaderInfo.BaseCharOffset) - (Row * theFont.RowPitch)) * theFont.ColFactor
        v = Row * theFont.RowFactor

        'Set the verticies
        With theFont.HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Colour = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).RHW = 1
            .Vertex(0).tu = u
            .Vertex(0).tv = v
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).z = 0
            .Vertex(1).Colour = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).tu = u + theFont.ColFactor
            .Vertex(1).tv = v
            .Vertex(1).X = theFont.HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).z = 0
            .Vertex(2).Colour = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = v + theFont.RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = theFont.HeaderInfo.CellHeight
            .Vertex(2).z = 0
            .Vertex(3).Colour = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).tu = u + theFont.ColFactor
            .Vertex(3).tv = v + theFont.RowFactor
            .Vertex(3).X = theFont.HeaderInfo.CellWidth
            .Vertex(3).Y = theFont.HeaderInfo.CellHeight
            .Vertex(3).z = 0
        End With
    Next LoopChar
End Sub

Public Sub RenderText(ByRef UseFont As CustomFont, ByVal text As String, ByVal X As Long, ByVal Y As Long, ByVal Color As Long, Optional ByVal alpha As Long = 255, Optional Shadow As Boolean = True)
Dim TempVA(0 To 3) As Vertex, TempStr() As String, Count As Long, Ascii() As Byte, i As Long, j As Long, TempColor As Long, yOffset As Single, ignoreChar As Long, resetColor As Long
Dim tmpNum As Long

    ' set the color
    Color = DX8Colour(Color, alpha)

    'Check for valid text to render
    If LenB(text) = 0 Then Exit Sub
    'Get the text into arrays (split by vbCrLf)
    TempStr = Split(text, vbCrLf)
    'Set the temp color (or else the first character has no color)
    TempColor = Color
    resetColor = TempColor
    'Set the texture
    D3DDevice.SetTexture 0, UseFont.Texture
    CurrentTexture = -1
    ' set the position
    X = X - UseFont.xOffset
    Y = Y - UseFont.yOffset
    'Loop through each line if there are line breaks (vbCrLf)
    tmpNum = UBound(TempStr)

    For i = 0 To tmpNum
        If Len(TempStr(i)) > 0 Then
            yOffset = (i * UseFont.CharHeight) + (i * 3)
            Count = 0
            'Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(i), vbFromUnicode)
            'Loop through the characters
            tmpNum = Len(TempStr(i))
            For j = 1 To tmpNum
                ' check for colour change
                If Mid$(TempStr(i), j, 1) = ColourChar Then
                    Color = Val(Mid$(TempStr(i), j + 1, 2))
                    ' make sure the colour exists
                    If Color = -1 Then
                        TempColor = resetColor
                    Else
                        TempColor = DX8Colour(Color, alpha)
                    End If
                    ignoreChar = 3
                End If
                ' check if we're ignoring this character
                If ignoreChar > 0 Then
                    ignoreChar = ignoreChar - 1
                Else
                    'Copy from the cached vertex array to the temp vertex array
                    Call CopyMemory(TempVA(0), UseFont.HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), FVF_Size * 4)
                    'Set up the verticies
                    TempVA(0).X = X + Count
                    TempVA(0).Y = Y + yOffset
                    TempVA(1).X = TempVA(1).X + X + Count
                    TempVA(1).Y = TempVA(0).Y
                    TempVA(2).X = TempVA(0).X
                    TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
                    TempVA(3).X = TempVA(1).X
                    TempVA(3).Y = TempVA(2).Y
                    'Set the colors
                    TempVA(0).Colour = TempColor
                    TempVA(1).Colour = TempColor
                    TempVA(2).Colour = TempColor
                    TempVA(3).Colour = TempColor
                    'Draw the verticies
                    Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, TempVA(0), FVF_Size)
                    'Shift over the the position to render the next character
                    Count = Count + UseFont.HeaderInfo.CharWidth(Ascii(j - 1))
                End If
            Next j
        End If
    Next i
End Sub

Public Function TextWidth(ByRef UseFont As CustomFont, ByVal text As String) As Long
Dim LoopI As Integer, tmpNum As Long, skipCount As Long

    'Make sure we have text
    If LenB(text) = 0 Then Exit Function
    
    'Loop through the text
    tmpNum = Len(text)
    For LoopI = 1 To tmpNum
        If Mid$(text, LoopI, 1) = ColourChar Then skipCount = 3
        If skipCount > 0 Then
            skipCount = skipCount - 1
        Else
            TextWidth = TextWidth + UseFont.HeaderInfo.CharWidth(Asc(Mid$(text, LoopI, 1)))
        End If
    Next LoopI
End Function

Public Function TextHeight(ByRef UseFont As CustomFont) As Long
    TextHeight = UseFont.HeaderInfo.CellHeight
End Function

Sub DrawActionMsg(ByVal Index As Integer)
        Dim X As Long, Y As Long, i As Long, Time As Long
    Dim LenMsg As Long

    If ActionMsg(Index).message = vbNullString Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(Index).Type

        Case ACTIONMsgSTATIC
            Time = 1500
            LenMsg = TextWidth(font(Fonts.rockwell_15), Trim$(ActionMsg(Index).message))

            If ActionMsg(Index).Y > 0 Then
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - (LenMsg / 2)
                Y = ActionMsg(Index).Y + PIC_Y
            Else
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - (LenMsg / 2)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) + 18
            End If

        Case ACTIONMsgSCROLL
            Time = 1500

            If ActionMsg(Index).Y > 0 Then
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) - 2 - (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            Else
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) + 18 + (ActionMsg(Index).Scroll * 0.001)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            End If

            ActionMsg(Index).alpha = ActionMsg(Index).alpha - 5

            If ActionMsg(Index).alpha <= 0 Then ClearActionMsg Index: Exit Sub

        Case ACTIONMsgSCREEN
            Time = 3000

            ' This will kill any action screen messages that there in the system
            For i = MAX_BYTE To 1 Step -1

                If ActionMsg(i).Type = ACTIONMsgSCREEN Then
                    If i <> Index Then
                        ClearActionMsg Index
                        Index = i
                    End If
                End If

            Next

            X = (400) - ((TextWidth(font(Fonts.rockwell_15), Trim$(ActionMsg(Index).message)) \ 2))
            Y = 24
    End Select

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    If ActionMsg(Index).Created > 0 Then
        RenderText font(Fonts.rockwell_15), ActionMsg(Index).message, X, Y, ActionMsg(Index).Color, ActionMsg(Index).alpha
    End If
End Sub

Public Function DrawMapAttributes()
Dim X As Long, Y As Long, tx As Long, ty As Long, theFont As Long

    theFont = Fonts.rockwellDec_10

    If frmEditor_Map.optAttribs.Value Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    With Map.TileData.Tile(X, Y)
                        tx = ((ConvertMapX(X * PIC_X)) - 4) + (PIC_X * 0.5)
                        ty = ((ConvertMapY(Y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                RenderText font(theFont), "B", tx, ty, BrightRed
                            Case TILE_TYPE_WARP
                                RenderText font(theFont), "W", tx, ty, BrightBlue
                            Case TILE_TYPE_ITEM
                                RenderText font(theFont), "I", tx, ty, White
                            Case TILE_TYPE_NPCAVOID
                                RenderText font(theFont), "N", tx, ty, White
                            Case TILE_TYPE_KEY
                                RenderText font(theFont), "K", tx, ty, White
                            Case TILE_TYPE_KEYOPEN
                                RenderText font(theFont), "O", tx, ty, White
                            Case TILE_TYPE_RESOURCE
                                RenderText font(theFont), "R", tx, ty, Green
                            Case TILE_TYPE_DOOR
                                RenderText font(theFont), "D", tx, ty, Brown
                            Case TILE_TYPE_NPCSPAWN
                                RenderText font(theFont), "S", tx, ty, Yellow
                            Case TILE_TYPE_SHOP
                                RenderText font(theFont), "S", tx, ty, BrightBlue
                            Case TILE_TYPE_SLIDE
                                RenderText font(theFont), "S", tx, ty, Pink
                            Case TILE_TYPE_CHAT
                                RenderText font(theFont), "C", tx, ty, Blue
                        End Select
                    End With
                End If
            Next
        Next
    End If
End Function

Public Sub AddText(ByVal text As String, ByVal Color As Long, Optional ByVal alpha As Long = 255, Optional Channel As Byte = 0)
Dim i As Long

    Chat_HighIndex = 0
    ' Move the rest of it up
    For i = (ChatLines - 1) To 1 Step -1
        If Len(Chat(i).text) > 0 Then
            If i > Chat_HighIndex Then Chat_HighIndex = i + 1
        End If
        Chat(i + 1) = Chat(i)
    Next
    
    Chat(1).text = text
    Chat(1).Color = Color
    Chat(1).visible = True
    Chat(1).timer = getTime
    Chat(1).Channel = Channel
End Sub

Sub RenderChat()
Dim Xo As Long, Yo As Long, Colour As Long, yOffset As Long, rLines As Long, lineCount As Long
Dim tmpText As String, i As Long, isVisible As Boolean, topWidth As Long, tmpArray() As String, X As Long
    
    ' set the position
    Xo = 19
    Yo = ScreenHeight - 41 '545 + 14
    
    ' loop through chat
    rLines = 1
    i = 1 + ChatScroll
    Do While rLines <= 8
        If i > ChatLines Then Exit Do
        lineCount = 0
        ' exit out early if we come to a blank string
        If Len(Chat(i).text) = 0 Then Exit Do
        ' get visible state
        isVisible = True
        If inSmallChat Then
            If Not Chat(i).visible Then isVisible = False
        End If
        If Options.channelState(Chat(i).Channel) = 0 Then isVisible = False
        ' make sure it's visible
        If isVisible Then
            ' render line
            Colour = Chat(i).Color
            ' check if we need to word wrap
            If TextWidth(font(Fonts.verdana_12), Chat(i).text) > ChatWidth Then
                ' word wrap
                tmpText = WordWrap(font(Fonts.verdana_12), Chat(i).text, ChatWidth, lineCount)
                ' can't have it going offscreen.
                If rLines + lineCount > 9 Then Exit Do
                ' continue on
                yOffset = yOffset - (14 * lineCount)
                RenderText font(Fonts.verdana_12), tmpText, Xo, Yo + yOffset, Colour
                rLines = rLines + lineCount
                ' set the top width
                tmpArray = Split(tmpText, vbNewLine)
                For X = 0 To UBound(tmpArray)
                    If TextWidth(font(Fonts.verdana_12), tmpArray(X)) > topWidth Then topWidth = TextWidth(font(Fonts.verdana_12), tmpArray(X))
                Next
            Else
                ' normal
                yOffset = yOffset - 14
                RenderText font(Fonts.verdana_12), Chat(i).text, Xo, Yo + yOffset, Colour
                rLines = rLines + 1
                ' set the top width
                If TextWidth(font(Fonts.verdana_12), Chat(i).text) > topWidth Then topWidth = TextWidth(font(Fonts.verdana_12), Chat(i).text)
            End If
        End If
        ' increment chat pointer
        i = i + 1
    Loop
    
    ' get the height of the small chat box
    SetChatHeight rLines * 14
    SetChatWidth topWidth
End Sub

Public Sub WordWrap_Array(ByVal text As String, ByVal MaxLineLen As Long, ByRef theArray() As String)
    Dim lineCount As Long, i As Long, Size As Long, lastSpace As Long, B As Long, tmpNum As Long

    'Too small of text
    If Len(text) < 2 Then
        ReDim theArray(1 To 1) As String
        theArray(1) = text
        Exit Sub
    End If

    ' default values
    B = 1
    lastSpace = 1
    Size = 0
    tmpNum = Len(text)

    For i = 1 To tmpNum

        ' if it's a space, store it
        Select Case Mid$(text, i, 1)
            Case " ": lastSpace = i
        End Select

        'Add up the size
        Size = Size + font(Fonts.georgiaDec_16).HeaderInfo.CharWidth(Asc(Mid$(text, i, 1)))

        'Check for too large of a size
        If Size > MaxLineLen Then
            'Check if the last space was too far back
            If i - lastSpace > 12 Then
                'Too far away to the last space, so break at the last character
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, B, (i - 1) - B))
                B = i - 1
                Size = 0
            Else
                'Break at the last space to preserve the word
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, B, lastSpace - B))
                B = lastSpace + 1
                'Count all the words we ignored (the ones that weren't printed, but are before "i")
                Size = TextWidth(font(Fonts.georgiaDec_16), Mid$(text, lastSpace, i - lastSpace))
            End If
        End If

        ' Remainder
        If i = Len(text) Then
            If B <> i Then
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = theArray(lineCount) & Mid$(text, B, i)
            End If
        End If
    Next
End Sub

Public Function WordWrap(theFont As CustomFont, ByVal text As String, ByVal MaxLineLen As Integer, Optional ByRef lineCount As Long) As String
    Dim TempSplit() As String, TSLoop As Long, lastSpace As Long, Size As Long, i As Long, B As Long, tmpNum As Long, skipCount As Long

    'Too small of text
    If Len(text) < 2 Then
        WordWrap = text
        Exit Function
    End If

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(text, vbNewLine)
    tmpNum = UBound(TempSplit)

    For TSLoop = 0 To tmpNum
        'Clear the values for the new line
        Size = 0
        B = 1
        lastSpace = 1

        'Add back in the vbNewLines
        If TSLoop < UBound(TempSplit()) Then TempSplit(TSLoop) = TempSplit(TSLoop) & vbNewLine

        'Only check lines with a space
        If InStr(1, TempSplit(TSLoop), " ") Then
            'Loop through all the characters
            tmpNum = Len(TempSplit(TSLoop))

            For i = 1 To tmpNum
                'If it is a space, store it so we can easily break at it
                Select Case Mid$(TempSplit(TSLoop), i, 1)
                    Case " "
                        lastSpace = i
                    Case ColourChar
                        skipCount = 3
                End Select
                
                If skipCount > 0 Then
                    skipCount = skipCount - 1
                Else
                    'Add up the size
                    Size = Size + theFont.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
                    'Check for too large of a size
                    If Size > MaxLineLen Then
                        'Check if the last space was too far back
                        If i - lastSpace > 12 Then
                            'Too far away to the last space, so break at the last character
                            WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, (i - 1) - B)) & vbNewLine
                            lineCount = lineCount + 1
                            B = i - 1
                            Size = 0
                        Else
                            'Break at the last space to preserve the word
                            WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, lastSpace - B)) & vbNewLine
                            lineCount = lineCount + 1
                            B = lastSpace + 1
                            'Count all the words we ignored (the ones that weren't printed, but are before "i")
                            Size = TextWidth(theFont, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                        End If
                    End If
    
                    'This handles the remainder
                    If i = Len(TempSplit(TSLoop)) Then
                        If B <> i Then
                            WordWrap = WordWrap & Mid$(TempSplit(TSLoop), B, i)
                            lineCount = lineCount + 1
                        End If
                    End If
                End If
            Next i
        Else
            WordWrap = WordWrap & TempSplit(TSLoop)
        End If
    Next TSLoop
End Function

Public Sub DrawPlayerName(ByVal Index As Long)
    Dim textX As Long, textY As Long, text As String, textSize As Long, Colour As Long
    
    text = Trim$(GetPlayerName(Index))
    textSize = TextWidth(font(Fonts.rockwell_15), text)
    ' get the colour
    Colour = White

    If GetPlayerAccess(Index) > 0 Then Colour = Pink
    If GetPlayerPK(Index) > 0 Then Colour = BrightRed
    textX = Player(Index).X * PIC_X + Player(Index).xOffset + (PIC_X \ 2) - (textSize \ 2)
    textY = Player(Index).Y * PIC_Y + Player(Index).yOffset - 32

    If GetPlayerSprite(Index) >= 1 And GetPlayerSprite(Index) <= CountChar Then
        textY = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - (mTexture(TextureChar(GetPlayerSprite(Index))).RealHeight / 4) + 12
    End If

    Call RenderText(font(Fonts.rockwell_15), text, ConvertMapX(textX), ConvertMapY(textY), Colour)
End Sub

Public Sub DrawNpcName(ByVal Index As Long)
    Dim textX As Long, textY As Long, text As String, textSize As Long, NpcNum As Long, Colour As Long
    NpcNum = MapNpc(Index).Num
    text = Trim$(Npc(NpcNum).Name)
    textSize = TextWidth(font(Fonts.rockwell_15), text)

    If Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_ATTACKWHENATTACKED Then
        ' get the colour
        If Npc(NpcNum).Level <= GetPlayerLevel(MyIndex) - 3 Then
            Colour = Grey
        ElseIf Npc(NpcNum).Level <= GetPlayerLevel(MyIndex) - 2 Then
            Colour = Green
        ElseIf Npc(NpcNum).Level > GetPlayerLevel(MyIndex) Then
            Colour = Red
        Else
            Colour = White
        End If
    Else
        Colour = White
    End If

    textX = MapNpc(Index).X * PIC_X + MapNpc(Index).xOffset + (PIC_X \ 2) - (textSize \ 2)
    textY = MapNpc(Index).Y * PIC_Y + MapNpc(Index).yOffset - 32

    If Npc(NpcNum).sprite >= 1 And Npc(NpcNum).sprite <= CountChar Then
        textY = MapNpc(Index).Y * PIC_Y + MapNpc(Index).yOffset - (mTexture(TextureChar(Npc(NpcNum).sprite)).RealHeight / 4) + 12
    End If

    Call RenderText(font(Fonts.rockwell_15), text, ConvertMapX(textX), ConvertMapY(textY), Colour)
End Sub

Function GetColStr(Colour As Long)
    If Colour < 10 Then
        GetColStr = "0" & Colour
    Else
        GetColStr = Colour
    End If
End Function
