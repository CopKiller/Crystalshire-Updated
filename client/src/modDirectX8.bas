Attribute VB_Name = "modDirectX8"
Option Explicit

' Texture wrapper
Public Tex_Anim() As Long
Public Tex_Char() As Long
Public Tex_Face() As Long
Public Tex_Item() As Long
Public Tex_Paperdoll() As Long
Public Tex_Resource() As Long
Public Tex_Spellicon() As Long
Public Tex_Tileset() As Long
Public Tex_Fog() As Long
Public Tex_GUI() As Long
Public Tex_Design() As Long
Public Tex_Gradient() As Long
Public Tex_Surface() As Long
Public Tex_Bars As Long
Public Tex_Blood As Long
Public Tex_Direction As Long
Public Tex_Misc As Long
Public Tex_Target As Long
Public Tex_Shadow As Long
Public Tex_Fader As Long
Public Tex_Blank As Long
Public Tex_Event As Long
Public Tex_Weather As Long
Public Tex_White As Long
Public Tex_Captcha() As Long

' Texture count
Public Count_Anim As Long
Public Count_Char As Long
Public Count_Face As Long
Public Count_GUI As Long
Public Count_Design As Long
Public Count_Gradient As Long
Public Count_Item As Long
Public Count_Paperdoll As Long
Public Count_Resource As Long
Public Count_Spellicon As Long
Public Count_Tileset As Long
Public Count_Fog As Long
Public Count_Surface As Long
Public Count_Captcha As Long

' Variables
Public DX8 As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8
Public D3DDevice As Direct3DDevice8
Public DXVB As Direct3DVertexBuffer8
Public D3DWindow As D3DPRESENT_PARAMETERS
Public mhWnd As Long
Public BackBuffer As Direct3DSurface8

Public Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE 'Or D3DFVF_SPECULAR

Public Type TextureStruct
    Texture As Direct3DTexture8
    data() As Byte
    w As Long
    h As Long
End Type

Public Type TextureDataStruct
    data() As Byte
End Type

Public Type Vertex
    X As Single
    Y As Single
    z As Single
    RHW As Single
    Colour As Long
    tu As Single
    tv As Single
End Type

Public mClip As RECT
Public Box(0 To 3) As Vertex
Public mTexture() As TextureStruct
Public mTextures As Long
Public CurrentTexture As Long

Public ScreenWidth As Long, ScreenHeight As Long
Public TileWidth As Long, TileHeight As Long
Public ScreenX As Long, ScreenY As Long
Public EndX As Long, EndY As Long
Public curResolution As Byte, isFullscreen As Boolean

Public Sub InitDX8(ByVal hWnd As Long)
Dim DispMode As D3DDISPLAYMODE, Width As Long, Height As Long

    mhWnd = hWnd

    Set DX8 = New DirectX8
    Set D3D = DX8.Direct3DCreate
    Set D3DX = New D3DX8
    
    ' set size
    GetResolutionSize curResolution, Width, Height
    ScreenWidth = Width
    ScreenHeight = Height
    TileWidth = (Width / 32) - 1
    TileHeight = (Height / 32) - 1
    ScreenX = (TileWidth) * PIC_X
    ScreenY = (TileHeight) * PIC_Y
    
    ' set up window
    Call D3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DispMode)
    DispMode.Format = D3DFMT_A8R8G8B8
    
    If Options.Fullscreen = 0 Then
        isFullscreen = False
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
        D3DWindow.hDeviceWindow = hWnd
        D3DWindow.BackBufferFormat = DispMode.Format
        D3DWindow.Windowed = 1
    Else
        isFullscreen = True
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
        D3DWindow.BackBufferCount = 1
        D3DWindow.BackBufferFormat = DispMode.Format
        D3DWindow.BackBufferWidth = ScreenWidth
        D3DWindow.BackBufferHeight = ScreenHeight
    End If
    
    Select Case Options.Render
        Case 1 ' hardware
            If LoadDirectX(D3DCREATE_HARDWARE_VERTEXPROCESSING, hWnd) <> 0 Then
                Options.Fullscreen = 0
                Options.Resolution = 0
                Options.Render = 0
                SaveOptions
                Call MsgBox("Could not initialize DirectX with hardware vertex processing.", vbCritical)
                Call DestroyGame
            End If
        Case 2 ' mixed
            If LoadDirectX(D3DCREATE_MIXED_VERTEXPROCESSING, hWnd) <> 0 Then
                Options.Fullscreen = 0
                Options.Resolution = 0
                Options.Render = 0
                SaveOptions
                Call MsgBox("Could not initialize DirectX with mixed vertex processing.", vbCritical)
                Call DestroyGame
            End If
        Case 3 ' software
            If LoadDirectX(D3DCREATE_SOFTWARE_VERTEXPROCESSING, hWnd) <> 0 Then
                Options.Fullscreen = 0
                Options.Resolution = 0
                Options.Render = 0
                SaveOptions
                Call MsgBox("Could not initialize DirectX with software vertex processing.", vbCritical)
                Call DestroyGame
            End If
        Case Else ' auto
            If LoadDirectX(D3DCREATE_HARDWARE_VERTEXPROCESSING, hWnd) <> 0 Then
                If LoadDirectX(D3DCREATE_MIXED_VERTEXPROCESSING, hWnd) <> 0 Then
                    If LoadDirectX(D3DCREATE_SOFTWARE_VERTEXPROCESSING, hWnd) <> 0 Then
                        Options.Fullscreen = 0
                        Options.Resolution = 0
                        Options.Render = 0
                        SaveOptions
                        Call MsgBox("Could not initialize DirectX.  DX8VB.dll may not be registered.", vbCritical)
                        Call DestroyGame
                    End If
                End If
            End If
    End Select
    
    ' Render states
    Call D3DDevice.SetVertexShader(FVF)
    Call D3DDevice.SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
    Call D3DDevice.SetRenderState(D3DRS_LIGHTING, False)
    Call D3DDevice.SetRenderState(D3DRS_ALPHABLENDENABLE, True)
    Call D3DDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
    Call D3DDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_CURRENT)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, 2)
    Call D3DDevice.SetStreamSource(0, DXVB, Len(Box(0)))
End Sub

Public Function LoadDirectX(ByVal BehaviourFlags As CONST_D3DCREATEFLAGS, ByVal hWnd As Long)
On Error GoTo ErrorInit

    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, BehaviourFlags, D3DWindow)
    Exit Function

ErrorInit:
    LoadDirectX = 1
End Function

Sub DestroyDX8()
Dim I As Long
    'For i = 1 To mTextures
    '    mTexture(i).data
    'Next
    If Not DX8 Is Nothing Then Set DX8 = Nothing
    If Not D3D Is Nothing Then Set D3D = Nothing
    If Not D3DX Is Nothing Then Set D3DX = Nothing
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
End Sub

Public Sub LoadTextures()
Dim I As Long
    ' Arrays
    Tex_Captcha = LoadTextureFiles(Count_Captcha, App.path & Path_Captcha)
    Tex_Tileset = LoadTextureFiles(Count_Tileset, App.path & Path_Tileset)
    Tex_Anim = LoadTextureFiles(Count_Anim, App.path & Path_Anim)
    Tex_Char = LoadTextureFiles(Count_Char, App.path & Path_Char)
    Tex_Face = LoadTextureFiles(Count_Face, App.path & Path_Face)
    Tex_Item = LoadTextureFiles(Count_Item, App.path & Path_Item)
    Tex_Paperdoll = LoadTextureFiles(Count_Paperdoll, App.path & Path_Paperdoll)
    Tex_Resource = LoadTextureFiles(Count_Resource, App.path & Path_Resource)
    Tex_Spellicon = LoadTextureFiles(Count_Spellicon, App.path & Path_Spellicon)
    Tex_GUI = LoadTextureFiles(Count_GUI, App.path & Path_GUI)
    Tex_Design = LoadTextureFiles(Count_Design, App.path & Path_Design)
    Tex_Gradient = LoadTextureFiles(Count_Gradient, App.path & Path_Gradient)
    Tex_Surface = LoadTextureFiles(Count_Surface, App.path & Path_Surface)
    ' Singles
    Tex_Bars = LoadTextureFile(App.path & Path_Graphics & "bars.png")
    Tex_Blood = LoadTextureFile(App.path & Path_Graphics & "blood.png")
    Tex_Direction = LoadTextureFile(App.path & Path_Graphics & "direction.png")
    Tex_Misc = LoadTextureFile(App.path & Path_Graphics & "misc.png")
    Tex_Target = LoadTextureFile(App.path & Path_Graphics & "target.png")
    Tex_Shadow = LoadTextureFile(App.path & Path_Graphics & "shadow.png")
    Tex_Fader = LoadTextureFile(App.path & Path_Graphics & "fader.png")
    Tex_Blank = LoadTextureFile(App.path & Path_Graphics & "blank.png")
    Tex_Weather = LoadTextureFile(App.path & Path_Graphics & "weather.png")
    Tex_White = LoadTextureFile(App.path & Path_Graphics & "white.png")
    Tex_Event = LoadTextureFile(App.path & Path_Graphics & "event.png")
End Sub

Public Function LoadTextureFiles(ByRef Counter As Long, ByVal path As String) As Long()
Dim Texture() As Long
Dim I As Long

    Counter = 1
    
    Do While Dir$(path & Counter + 1 & ".png") <> vbNullString
        Counter = Counter + 1
    Loop
    
    ReDim Texture(0 To Counter)
    
    For I = 1 To Counter
        Texture(I) = LoadTextureFile(path & I & ".png")
        DoEvents
    Next
    
    LoadTextureFiles = Texture
End Function

Public Function LoadTextureFile(ByVal path As String, Optional ByVal DontReuse As Boolean) As Long
Dim data() As Byte
Dim f As Long

    If Dir$(path) = vbNullString Then
        Call MsgBox("""" & path & """ could not be found.")
        End
    End If
    
    f = FreeFile
    Open path For Binary As #f
        ReDim data(0 To LOF(f) - 1)
        Get #f, , data
    Close #f
    
    LoadTextureFile = LoadTexture(data, DontReuse)
End Function

Public Function LoadTexture(ByRef data() As Byte, Optional ByVal DontReuse As Boolean) As Long
Dim I As Long

    If AryCount(data) = 0 Then
        Exit Function
    End If
    
    mTextures = mTextures + 1
    LoadTexture = mTextures
    ReDim Preserve mTexture(1 To mTextures) As TextureStruct
    mTexture(mTextures).w = ByteToInt(data(18), data(19))
    mTexture(mTextures).h = ByteToInt(data(22), data(23))
    mTexture(mTextures).data = data
    Set mTexture(mTextures).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, data(0), AryCount(data), mTexture(mTextures).w, mTexture(mTextures).h, D3DX_DEFAULT, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
End Function

Public Sub CheckGFX()
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then
        Do While D3DDevice.TestCooperativeLevel = D3DERR_DEVICELOST
           DoEvents
        Loop
        Call ResetGFX
    End If
End Sub

Public Sub ResetGFX()
Dim Temp() As TextureDataStruct
Dim I As Long, n As Long

    n = mTextures
    ReDim Temp(1 To n)
    For I = 1 To n
        Set mTexture(I).Texture = Nothing
        Temp(I).data = mTexture(I).data
    Next
    
    Erase mTexture
    mTextures = 0
    
    Call D3DDevice.Reset(D3DWindow)
    Call D3DDevice.SetVertexShader(FVF)
    Call D3DDevice.SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
    Call D3DDevice.SetRenderState(D3DRS_LIGHTING, False)
    Call D3DDevice.SetRenderState(D3DRS_ALPHABLENDENABLE, True)
    Call D3DDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
    Call D3DDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_CURRENT)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, 2)
    
    For I = 1 To n
        Call LoadTexture(Temp(I).data)
    Next
End Sub

Public Sub SetTexture(ByVal textureNum As Long)
    If textureNum > 0 Then
        Call D3DDevice.SetTexture(0, mTexture(textureNum).Texture)
        CurrentTexture = textureNum
    Else
        Call D3DDevice.SetTexture(0, Nothing)
        CurrentTexture = 0
    End If
End Sub

Public Sub RenderTexture(Texture As Long, ByVal X As Long, ByVal Y As Long, ByVal sX As Single, ByVal sY As Single, ByVal w As Long, ByVal h As Long, ByVal sW As Single, ByVal sH As Single, Optional ByVal Colour As Long = -1, Optional ByVal offset As Boolean = False)
    SetTexture Texture
    RenderGeom X, Y, sX, sY, w, h, sW, sH, Colour, offset
End Sub

Public Sub RenderGeom(ByVal X As Long, ByVal Y As Long, ByVal sX As Single, ByVal sY As Single, ByVal w As Long, ByVal h As Long, ByVal sW As Single, ByVal sH As Single, Optional ByVal Colour As Long = -1, Optional ByVal offset As Boolean = False)
Dim I As Long

    If CurrentTexture = 0 Then Exit Sub
    If w = 0 Then Exit Sub
    If h = 0 Then Exit Sub
    If sW = 0 Then Exit Sub
    If sH = 0 Then Exit Sub
    
    If mClip.Right <> 0 Then
        If mClip.Top <> 0 Then
            If mClip.Left > X Then
                sX = sX + (mClip.Left - X) / (w / sW)
                sW = sW - (mClip.Left - X) / (w / sW)
                w = w - (mClip.Left - X)
                X = mClip.Left
            End If
            
            If mClip.Top > Y Then
                sY = sY + (mClip.Top - Y) / (h / sH)
                sH = sH - (mClip.Top - Y) / (h / sH)
                h = h - (mClip.Top - Y)
                Y = mClip.Top
            End If
            
            If mClip.Right < X + w Then
                sW = sW - (X + w - mClip.Right) / (w / sW)
                w = -X + mClip.Right
            End If
            
            If mClip.bottom < Y + h Then
                sH = sH - (Y + h - mClip.bottom) / (h / sH)
                h = -Y + mClip.bottom
            End If
            
            If w <= 0 Then Exit Sub
            If h <= 0 Then Exit Sub
            If sW <= 0 Then Exit Sub
            If sH <= 0 Then Exit Sub
        End If
    End If
    
    Call GeomCalc(Box, CurrentTexture, X, Y, w, h, sX, sY, sW, sH, Colour)
    Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), Len(Box(0)))
End Sub

Public Sub GeomCalc(ByRef Geom() As Vertex, ByVal textureNum As Long, ByVal X As Single, ByVal Y As Single, ByVal w As Integer, ByVal h As Integer, ByVal sX As Single, ByVal sY As Single, ByVal sW As Single, ByVal sH As Single, ByVal Colour As Long)
    sW = (sW + sX) / mTexture(textureNum).w + 0.000003
    sH = (sH + sY) / mTexture(textureNum).h + 0.000003
    sX = sX / mTexture(textureNum).w + 0.000003
    sY = sY / mTexture(textureNum).h + 0.000003
    Geom(0) = MakeVertex(X, Y, 0, 1, Colour, 1, sX, sY)
    Geom(1) = MakeVertex(X + w, Y, 0, 1, Colour, 0, sW, sY)
    Geom(2) = MakeVertex(X, Y + h, 0, 1, Colour, 0, sX, sH)
    Geom(3) = MakeVertex(X + w, Y + h, 0, 1, Colour, 0, sW, sH)
End Sub

Private Sub GeomSetBox(ByVal X As Single, ByVal Y As Single, ByVal w As Integer, ByVal h As Integer, ByVal Colour As Long)
    Box(0) = MakeVertex(X, Y, 0, 1, Colour, 0, 0, 0)
    Box(1) = MakeVertex(X + w, Y, 0, 1, Colour, 0, 0, 0)
    Box(2) = MakeVertex(X, Y + h, 0, 1, Colour, 0, 0, 0)
    Box(3) = MakeVertex(X + w, Y + h, 0, 1, Colour, 0, 0, 0)
End Sub

Private Function MakeVertex(X As Single, Y As Single, z As Single, RHW As Single, Colour As Long, Specular As Long, tu As Single, tv As Single) As Vertex
    MakeVertex.X = X
    MakeVertex.Y = Y
    MakeVertex.z = z
    MakeVertex.RHW = RHW
    MakeVertex.Colour = Colour
    'MakeVertex.Specular = Specular
    MakeVertex.tu = tu
    MakeVertex.tv = tv
End Function

' GDI rendering
Public Sub GDIRenderAnimation()
    Dim I As Long, Animationnum As Long, ShouldRender As Boolean, Width As Long, Height As Long, looptime As Long, FrameCount As Long
    Dim sX As Long, sY As Long, sRECT As RECT
    sRECT.Top = 0
    sRECT.bottom = 192
    sRECT.Left = 0
    sRECT.Right = 192

    For I = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(I).value

        If Animationnum <= 0 Or Animationnum > Count_Anim Then
            ' don't render lol
        Else
            looptime = frmEditor_Animation.scrlLoopTime(I)

            FrameCount = frmEditor_Animation.scrlFrameCount(I)
            ShouldRender = False

            ' check if we need to render new frame
            If AnimEditorTimer(I) + looptime <= GetTickCount Then

                ' check if out of range
                If AnimEditorFrame(I) >= FrameCount Then
                    AnimEditorFrame(I) = 1
                Else
                    AnimEditorFrame(I) = AnimEditorFrame(I) + 1
                End If

                AnimEditorTimer(I) = GetTickCount
                ShouldRender = True
            End If

            If ShouldRender Then
                If frmEditor_Animation.scrlFrameCount(I).value > 0 Then
                    ' total width divided by frame count
                    Width = 192
                    Height = 192
                    sY = (Height * ((AnimEditorFrame(I) - 1) \ AnimColumns))
                    sX = (Width * (((AnimEditorFrame(I) - 1) Mod AnimColumns)))
                    ' Start Rendering
                    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
                    Call D3DDevice.BeginScene
                    'EngineRenderRectangle Tex_Anim(Animationnum), 0, 0, sX, sY, width, height, width, height
                    RenderTexture Tex_Anim(Animationnum), 0, 0, sX, sY, Width, Height, Width, Height
                    ' Finish Rendering
                    Call D3DDevice.EndScene
                    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Animation.picSprite(I).hWnd, ByVal 0)
                End If
            End If
        End If

    Next

End Sub

Public Sub GDIRenderChar(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim Height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Char Then Exit Sub
    Height = 32
    Width = 32
    sRECT.Top = 0
    sRECT.bottom = sRECT.Top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    RenderTexture Tex_Char(sprite), 0, 0, 0, 0, Width, Height, Width, Height
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderFace(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim Height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Face Then Exit Sub
    Height = mTexture(Tex_Face(sprite)).h
    Width = mTexture(Tex_Face(sprite)).w

    If Height = 0 Or Width = 0 Then
        Height = 1
        Width = 1
    End If

    sRECT.Top = 0
    sRECT.bottom = sRECT.Top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle Tex_Face(sprite), 0, 0, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_Face(sprite), 0, 0, 0, 0, Width, Height, Width, Height
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Sub GDIRenderEventGraphic()
    Dim Height As Long, Width As Long, GraphicType As Long, graphicNum As Long, sX As Long, sY As Long, texNum As Long
    Dim sRECT As RECT, Graphic As Long

    If Not frmEditor_Events.visible Then Exit Sub
    If curPageNum = 0 Then Exit Sub
    
    GraphicType = tmpEvent.EventPage(curPageNum).GraphicType
    Graphic = tmpEvent.EventPage(curPageNum).Graphic
    sX = tmpEvent.EventPage(curPageNum).GraphicX
    sY = tmpEvent.EventPage(curPageNum).GraphicY
    
    If GraphicType = 0 Then Exit Sub
    If Graphic = 0 Then Exit Sub
    
    Height = 32
    Width = 32
    
    Select Case GraphicType
        Case 0 ' nothing
            texNum = 0
        Case 1 ' Character
            If Graphic <= Count_Char Then texNum = Tex_Char(Graphic) Else texNum = 0
        Case 2 ' Tileset
            If Graphic <= Count_Tileset Then texNum = Tex_Tileset(Graphic) Else texNum = 0
    End Select
    
    If texNum = 0 Then
        frmEditor_Events.picGraphic.Cls
        Exit Sub
    End If
    
    sRECT.Top = 0
    sRECT.bottom = sRECT.Top + frmEditor_Events.picGraphic.ScaleHeight
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + frmEditor_Events.picGraphic.ScaleWidth
    
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, DX8Colour(White, 255), 1#, 0)
    Call D3DDevice.BeginScene

    RenderTexture texNum, (frmEditor_Events.picGraphic.ScaleWidth / 2) - 16, (frmEditor_Events.picGraphic.ScaleHeight / 2) - 16, sX * 32, sY * 32, Width, Height, Width, Height

    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Events.picGraphic.hWnd, ByVal 0)
End Sub

Sub GDIRenderEventGraphicSel()
    Dim Height As Long, Width As Long, GraphicType As Long, graphicNum As Long, sX As Long, sY As Long, texNum As Long
    Dim sRECT As RECT, Graphic As Long

    If Not frmEditor_Events.visible Then Exit Sub
    If Not frmEditor_Events.fraGraphic.visible Then Exit Sub
    If curPageNum = 0 Then Exit Sub
    
    GraphicType = tmpEvent.EventPage(curPageNum).GraphicType
    Graphic = tmpEvent.EventPage(curPageNum).Graphic
    
    If GraphicType = 0 Then Exit Sub
    If Graphic = 0 Then Exit Sub
    
    Select Case GraphicType
        Case 0 ' nothing
            texNum = 0
        Case 1 ' Character
            If Graphic <= Count_Char Then texNum = Tex_Char(Graphic) Else texNum = 0
        Case 2 ' Tileset
            If Graphic <= Count_Tileset Then texNum = Tex_Tileset(Graphic) Else texNum = 0
    End Select
    
    If texNum = 0 Then
        frmEditor_Events.picGraphicSel.Cls
        Exit Sub
    End If
    
    Width = mTexture(texNum).w
    Height = mTexture(texNum).h
    
    sRECT.Top = 0
    sRECT.bottom = sRECT.Top + frmEditor_Events.picGraphicSel.ScaleHeight
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + frmEditor_Events.picGraphicSel.ScaleWidth
    
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, DX8Colour(White, 255), 1#, 0)
    Call D3DDevice.BeginScene

    RenderTexture texNum, 0, 0, 0, 0, Width, Height, Width, Height
    RenderDesign DesignTypes.desTileBox, GraphicSelX * 32, GraphicSelY * 32, 32, 32

    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Events.picGraphicSel.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderTileset()
    Dim Height As Long, Width As Long, tileSet As Byte, sRECT As RECT
    ' find tileset number
    tileSet = frmEditor_Map.scrlTileSet.value

    ' exit out if doesn't exist
    If tileSet <= 0 Or tileSet > Count_Tileset Then Exit Sub
    Height = mTexture(Tex_Tileset(tileSet)).h
    Width = mTexture(Tex_Tileset(tileSet)).w

    If Height = 0 Or Width = 0 Then
        Height = 1
        Width = 1
    End If

    frmEditor_Map.picBackSelect.Width = Width
    frmEditor_Map.picBackSelect.Height = Height
    sRECT.Top = 0
    sRECT.bottom = Height
    sRECT.Left = 0
    sRECT.Right = Width

    ' change selected shape for autotiles
    If frmEditor_Map.scrlAutotile.value > 0 Then

        Select Case frmEditor_Map.scrlAutotile.value

            Case 1 ' autotile
                shpSelectedWidth = 64
                shpSelectedHeight = 96

            Case 2 ' fake autotile
                shpSelectedWidth = 32
                shpSelectedHeight = 32

            Case 3 ' animated
                shpSelectedWidth = 192
                shpSelectedHeight = 96

            Case 4 ' cliff
                shpSelectedWidth = 64
                shpSelectedHeight = 64

            Case 5 ' waterfall
                shpSelectedWidth = 64
                shpSelectedHeight = 96
        End Select

    End If

    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, DX8Colour(White, 255), 1#, 0)
    Call D3DDevice.BeginScene

    'EngineRenderRectangle Tex_Tileset(Tileset), 0, 0, 0, 0, width, height, width, height, width, height
    If Tex_Tileset(tileSet) <= 0 Then Exit Sub
    RenderTexture Tex_Tileset(tileSet), 0, 0, 0, 0, Width, Height, Width, Height
    ' draw selection boxes
    RenderDesign DesignTypes.desTileBox, shpSelectedLeft, shpSelectedTop, shpSelectedWidth, shpSelectedHeight
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Map.picBackSelect.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderItem(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim Height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Item Then Exit Sub
    Height = mTexture(Tex_Item(sprite)).h
    Width = mTexture(Tex_Item(sprite)).w
    sRECT.Top = 0
    sRECT.bottom = 32
    sRECT.Left = 0
    sRECT.Right = 32
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle Tex_Item(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Item(sprite), 0, 0, 0, 0, 32, 32, 32, 32
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderResource(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim Height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Resource Then Exit Sub
    Height = mTexture(Tex_Resource(sprite)).h
    Width = mTexture(Tex_Resource(sprite)).w
    sRECT.Top = 0
    sRECT.bottom = 152
    sRECT.Left = 0
    sRECT.Right = 152
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    RenderTexture Tex_Resource(sprite), 0, 0, 0, 0, Width, Height, Width, Height
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub


Public Sub GDIRenderSpell(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim Height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Spellicon Then Exit Sub
    Height = mTexture(Tex_Spellicon(sprite)).h
    Width = mTexture(Tex_Spellicon(sprite)).w

    If Height = 0 Or Width = 0 Then
        Height = 1
        Width = 1
    End If

    sRECT.Top = 0
    sRECT.bottom = Height
    sRECT.Left = 0
    sRECT.Right = Width
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle Tex_Spellicon(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Spellicon(sprite), 0, 0, 0, 0, 32, 32, 32, 32
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

' Directional blocking
Public Sub DrawDirection(ByVal X As Long, ByVal Y As Long)
    Dim I As Long, Top As Long, Left As Long
    ' render grid
    Top = 24
    Left = 0
    'EngineRenderRectangle Tex_Direction, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), left, top, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Direction, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), Left, Top, 32, 32, 32, 32

    ' render dir blobs
    For I = 1 To 4
        Left = (I - 1) * 8

        ' find out whether render blocked or not
        If Not isDirBlocked(map.TileData.Tile(X, Y).DirBlock, CByte(I)) Then
            Top = 8
        Else
            Top = 16
        End If

        'render!
        'EngineRenderRectangle Tex_Direction, ConvertMapX(x * PIC_X) + DirArrowX(i), ConvertMapY(y * PIC_Y) + DirArrowY(i), left, top, 8, 8, 8, 8, 8, 8
        RenderTexture Tex_Direction, ConvertMapX(X * PIC_X) + DirArrowX(I), ConvertMapY(Y * PIC_Y) + DirArrowY(I), Left, Top, 8, 8, 8, 8
    Next

End Sub

Public Sub DrawFade()
    RenderTexture Tex_Blank, 0, 0, 0, 0, ScreenWidth, ScreenHeight, 32, 32, DX8Colour(White, fadeAlpha)
End Sub

Public Sub DrawFog()
    Dim fogNum As Long, Colour As Long, X As Long, Y As Long, RenderState As Long
    fogNum = CurrentFog

    If fogNum <= 0 Or fogNum > Count_Fog Then Exit Sub
    Colour = D3DColorRGBA(255, 255, 255, 255 - CurrentFogOpacity)
    RenderState = 0

    ' render state
    Select Case RenderState

        Case 1 ' Additive
            D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

        Case 2 ' Subtractive
            D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SUBTRACT
            D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
    End Select

    For X = 0 To ((map.MapData.MaxX * 32) / 256) + 1
        For Y = 0 To ((map.MapData.MaxY * 32) / 256) + 1
            RenderTexture Tex_Fog(fogNum), (X * 256) + fogOffsetX, (Y * 256) + fogOffsetY, 0, 0, 256, 256, 256, 256, Colour
        Next
    Next

    ' reset render state
    If RenderState > 0 Then
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    End If

End Sub

Public Sub DrawTint()
    Dim Color As Long
    Color = D3DColorRGBA(CurrentTintR, CurrentTintG, CurrentTintB, CurrentTintA)
    RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, Color
End Sub

Public Sub DrawWeather()
    Dim Color As Long, I As Long, SpriteLeft As Long
    For I = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(I).InUse Then
            If WeatherParticle(I).Type = WEATHER_TYPE_STORM Then
                SpriteLeft = 0
            Else
                SpriteLeft = WeatherParticle(I).Type - 1
            End If
            RenderTexture Tex_Weather, ConvertMapX(WeatherParticle(I).X), ConvertMapY(WeatherParticle(I).Y), SpriteLeft * 32, 0, 32, 32, 32, 32, -1
        End If
    Next
End Sub

Public Sub DrawAutoTile(ByVal layernum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim yOffset As Long, xOffset As Long

    ' calculate the offset
    Select Case map.TileData.Tile(X, Y).Autotile(layernum)

        Case AUTOTILE_WATERFALL
            yOffset = (waterfallFrame - 1) * 32

        Case AUTOTILE_ANIM
            xOffset = autoTileFrame * 64

        Case AUTOTILE_CLIFF
            yOffset = -32
    End Select

    ' Draw the quarter
    RenderTexture Tex_Tileset(map.TileData.Tile(X, Y).Layer(layernum).tileSet), destX, destY, Autotile(X, Y).Layer(layernum).srcX(quarterNum) + xOffset, Autotile(X, Y).Layer(layernum).srcY(quarterNum) + yOffset, 16, 16, 16, 16
End Sub

Sub DrawTileSelection()
    Dim tileSet As Byte
    ' find tileset number
    tileSet = frmEditor_Map.scrlTileSet.value
    
    ' exit out if doesn't exist
    If tileSet <= 0 Or tileSet > Count_Tileset Then Exit Sub
    
    If frmEditor_Map.optEvents.value Then
        RenderDesign DesignTypes.desTileBox, ConvertMapX(selTileX * PIC_X), ConvertMapY(selTileY * PIC_Y), 32, 32
    Else
'        If frmEditor_Map.scrlAutotile > 0 Then
'            RenderDesign DesignTypes.desTileBox, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), 32, 32
'        Else
'            RenderDesign DesignTypes.desTileBox, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), shpSelectedWidth, shpSelectedHeight
'        End If

       If frmEditor_Map.scrlAutotile.value > 0 Then
           RenderTexture Tex_Tileset(tileSet), ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), shpSelectedLeft, shpSelectedTop, shpSelectedWidth, shpSelectedHeight, shpSelectedWidth, shpSelectedHeight, D3DColorRGBA(255, 255, 255, 200)
       Else
           RenderTexture Tex_Tileset(tileSet), ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), shpSelectedLeft, shpSelectedTop, shpSelectedWidth, shpSelectedHeight, shpSelectedWidth, shpSelectedHeight
       End If
    End If
End Sub

' Rendering Procedures
Public Sub DrawMapTile(ByVal X As Long, ByVal Y As Long)
Dim I As Long, tileSet As Long, sX As Long, sY As Long

    With map.TileData.Tile(X, Y)
        ' draw the map
        For I = MapLayer.Ground To MapLayer.Mask2
            ' skip tile if tileset isn't set
            If Autotile(X, Y).Layer(I).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.Layer(I).tileSet), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(I).X * 32, .Layer(I).Y * 32, 32, 32, 32, 32
            ElseIf Autotile(X, Y).Layer(I).RenderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile I, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
                DrawAutoTile I, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
                DrawAutoTile I, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
                DrawAutoTile I, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
            ElseIf Autotile(X, Y).Layer(I).RenderState = RENDER_STATE_APPEAR Then
                ' check if it's fading
                If TempTile(X, Y).fadeAlpha(I) > 0 Then
                    ' render it
                    tileSet = map.TileData.Tile(X, Y).Layer(I).tileSet
                    sX = map.TileData.Tile(X, Y).Layer(I).X
                    sY = map.TileData.Tile(X, Y).Layer(I).Y
                    RenderTexture Tex_Tileset(tileSet), ConvertMapX(X * 32), ConvertMapY(Y * 32), sX * 32, sY * 32, 32, 32, 32, 32, DX8Colour(White, TempTile(X, Y).fadeAlpha(I))
                End If
            End If
        Next
    End With
End Sub

Public Sub DrawMapFringeTile(ByVal X As Long, ByVal Y As Long)
    Dim I As Long

    With map.TileData.Tile(X, Y)
        ' draw the map
        For I = MapLayer.Fringe To MapLayer.Fringe2

            ' skip tile if tileset isn't set
            If Autotile(X, Y).Layer(I).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.Layer(I).tileSet), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(I).X * 32, .Layer(I).Y * 32, 32, 32, 32, 32
            ElseIf Autotile(X, Y).Layer(I).RenderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile I, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
                DrawAutoTile I, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
                DrawAutoTile I, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
                DrawAutoTile I, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
            End If
        Next
    End With
End Sub

Public Sub DrawHotbar()
    Dim xO As Long, yO As Long, Width As Long, Height As Long, I As Long, t As Long, sS As String
    
    xO = Windows(GetWindowIndex("winHotbar")).Window.Left
    yO = Windows(GetWindowIndex("winHotbar")).Window.Top
    
    ' render start + end wood
    RenderTexture Tex_GUI(31), xO - 1, yO + 3, 0, 0, 11, 26, 11, 26
    RenderTexture Tex_GUI(31), xO + 407, yO + 3, 0, 0, 11, 26, 11, 26
    
    For I = 1 To MAX_HOTBAR
        xO = Windows(GetWindowIndex("winHotbar")).Window.Left + HotbarLeft + ((I - 1) * HotbarOffsetX)
        yO = Windows(GetWindowIndex("winHotbar")).Window.Top + HotbarTop
        Width = 36
        Height = 36
        ' don't render last one
        If I <> 10 Then
            ' render wood
            RenderTexture Tex_GUI(32), xO + 30, yO + 3, 0, 0, 13, 26, 13, 26
        End If
        ' render box
        RenderTexture Tex_GUI(30), xO - 2, yO - 2, 0, 0, Width, Height, Width, Height
        ' render icon
        If Not (DragBox.Origin = origin_Hotbar And DragBox.Slot = I) Then
            Select Case Hotbar(I).sType
                Case 1 ' inventory
                    If Len(Item(Hotbar(I).Slot).name) > 0 And Item(Hotbar(I).Slot).Pic > 0 Then
                        RenderTexture Tex_Item(Item(Hotbar(I).Slot).Pic), xO, yO, 0, 0, 32, 32, 32, 32
                    End If
                Case 2 ' spell
                    If Len(Spell(Hotbar(I).Slot).name) > 0 And Spell(Hotbar(I).Slot).icon > 0 Then
                        RenderTexture Tex_Spellicon(Spell(Hotbar(I).Slot).icon), xO, yO, 0, 0, 32, 32, 32, 32
                        For t = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(t).Spell > 0 Then
                                If PlayerSpells(t).Spell = Hotbar(I).Slot And SpellCD(t) > 0 Then
                                    RenderTexture Tex_Spellicon(Spell(Hotbar(I).Slot).icon), xO, yO, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                                End If
                            End If
                        Next
                    End If
            End Select
        End If
        ' draw the numbers
        sS = Str(I)
        If I = 10 Then sS = "0"
        RenderText font(Fonts.rockwellDec_15), sS, xO + 4, yO + 19, White
    Next
End Sub

Public Sub RenderAppearTileFade()
Dim X As Long, Y As Long, tileSet As Long, sX As Long, sY As Long, layernum As Long

    For X = 0 To map.MapData.MaxX
        For Y = 0 To map.MapData.MaxY
            For layernum = MapLayer.Ground To MapLayer.Mask
                ' check if it's fading
                If TempTile(X, Y).fadeAlpha(layernum) > 0 Then
                    ' render it
                    tileSet = map.TileData.Tile(X, Y).Layer(layernum).tileSet
                    sX = map.TileData.Tile(X, Y).Layer(layernum).X
                    sY = map.TileData.Tile(X, Y).Layer(layernum).Y
                    RenderTexture Tex_Tileset(tileSet), ConvertMapX(X * 32), ConvertMapY(Y * 32), sX * 32, sY * 32, 32, 32, 32, 32, DX8Colour(White, TempTile(X, Y).fadeAlpha(layernum))
                End If
            Next
        Next
    Next
End Sub

Public Sub DrawCharacter()
    Dim xO As Long, yO As Long, Width As Long, Height As Long, I As Long, sprite As Long, itemNum As Long, ItemPic As Long
    
    xO = Windows(GetWindowIndex("winCharacter")).Window.Left
    yO = Windows(GetWindowIndex("winCharacter")).Window.Top
    
    ' Render bottom
    RenderTexture Tex_GUI(37), xO + 4, yO + 314, 0, 0, 40, 38, 40, 38
    RenderTexture Tex_GUI(37), xO + 44, yO + 314, 0, 0, 40, 38, 40, 38
    RenderTexture Tex_GUI(37), xO + 84, yO + 314, 0, 0, 40, 38, 40, 38
    RenderTexture Tex_GUI(37), xO + 124, yO + 314, 0, 0, 46, 38, 46, 38
    
    ' render top wood
    RenderTexture Tex_GUI(1), xO + 4, yO + 23, 100, 100, 166, 291, 166, 291
    
    ' loop through equipment
    For I = 1 To Equipment.Equipment_Count - 1
        itemNum = GetPlayerEquipment(MyIndex, I)

        ' get the item sprite
        If itemNum > 0 Then
            ItemPic = Tex_Item(Item(itemNum).Pic)
        Else
            ' no item equiped - use blank image
            ItemPic = Tex_GUI(37 + I)
        End If
        
        yO = Windows(GetWindowIndex("winCharacter")).Window.Top + EqTop
        xO = Windows(GetWindowIndex("winCharacter")).Window.Left + EqLeft + ((EqOffsetX + 32) * (((I - 1) Mod EqColumns)))

        RenderTexture ItemPic, xO, yO, 0, 0, 32, 32, 32, 32
    Next
End Sub

Public Sub DrawSkills()
    Dim xO As Long, yO As Long, Width As Long, Height As Long, I As Long, Y As Long, spellnum As Long, spellPic As Long, X As Long, Top As Long, Left As Long
    
    xO = Windows(GetWindowIndex("winSkills")).Window.Left
    yO = Windows(GetWindowIndex("winSkills")).Window.Top
    
    Width = Windows(GetWindowIndex("winSkills")).Window.Width
    Height = Windows(GetWindowIndex("winSkills")).Window.Height
    
    ' render green
    RenderTexture Tex_GUI(34), xO + 4, yO + 23, 0, 0, Width - 8, Height - 27, 4, 4
    
    Width = 76
    Height = 76
    
    Y = yO + 23
    ' render grid - row
    For I = 1 To 4
        If I = 4 Then Height = 42
        RenderTexture Tex_GUI(35), xO + 4, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 80, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 156, Y, 0, 0, 42, Height, 42, Height
        Y = Y + 76
    Next
    
    ' actually draw the icons
    For I = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(I).Spell
        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            ' not dragging?
            If Not (DragBox.Origin = origin_Spells And DragBox.Slot = I) Then
                spellPic = Spell(spellnum).icon
    
                If spellPic > 0 And spellPic <= Count_Spellicon Then
                    Top = yO + InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                    Left = xO + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
    
                    RenderTexture Tex_Spellicon(spellPic), Left, Top, 0, 0, 32, 32, 32, 32
                End If
            End If
        End If
    Next
End Sub

Public Sub RenderMapName()
Dim zonetype As String, Colour As Long

    If map.MapData.Moral = 0 Then
        zonetype = "PK Zone"
        Colour = Red
    ElseIf map.MapData.Moral = 1 Then
        zonetype = "Safe Zone"
        Colour = White
    ElseIf map.MapData.Moral = 2 Then
        zonetype = "Boss Chamber"
        Colour = Grey
    End If
    
    RenderText font(Fonts.rockwellDec_10), Trim$(map.MapData.name) & " - " & zonetype, ScreenWidth - 15 - TextWidth(font(Fonts.rockwellDec_10), Trim$(map.MapData.name) & " - " & zonetype), 45, Colour, 255
End Sub

Public Sub DrawShopBackground()
    Dim xO As Long, yO As Long, Width As Long, Height As Long, I As Long, Y As Long
    
    xO = Windows(GetWindowIndex("winShop")).Window.Left
    yO = Windows(GetWindowIndex("winShop")).Window.Top
    Width = Windows(GetWindowIndex("winShop")).Window.Width
    Height = Windows(GetWindowIndex("winShop")).Window.Height
    
    ' render green
    RenderTexture Tex_GUI(34), xO + 4, yO + 23, 0, 0, Width - 8, Height - 27, 4, 4
    
    Width = 76
    Height = 76
    
    Y = yO + 23
    ' render grid - row
    For I = 1 To 3
        If I = 3 Then Height = 42
        RenderTexture Tex_GUI(35), xO + 4, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 80, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 156, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 232, Y, 0, 0, 42, Height, 42, Height
        Y = Y + 76
    Next
    ' render bottom wood
    RenderTexture Tex_GUI(1), xO + 4, Y - 34, 0, 0, 270, 72, 270, 72
End Sub

Public Sub DrawShop()
Dim xO As Long, yO As Long, ItemPic As Long, itemNum As Long, Amount As Long, I As Long, Top As Long, Left As Long, Y As Long, X As Long, Colour As Long

    If InShop = 0 Then Exit Sub
    
    xO = Windows(GetWindowIndex("winShop")).Window.Left
    yO = Windows(GetWindowIndex("winShop")).Window.Top
    
    If Not shopIsSelling Then
        ' render the shop items
        For I = 1 To MAX_TRADES
            itemNum = Shop(InShop).TradeItem(I).Item
            
            ' draw early
            Top = yO + ShopTop + ((ShopOffsetY + 32) * ((I - 1) \ ShopColumns))
            Left = xO + ShopLeft + ((ShopOffsetX + 32) * (((I - 1) Mod ShopColumns)))
            ' draw selected square
            If shopSelectedSlot = I Then RenderTexture Tex_GUI(61), Left, Top, 0, 0, 32, 32, 32, 32
            
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
                ItemPic = Item(itemNum).Pic
                If ItemPic > 0 And ItemPic <= Count_Item Then
                    ' draw item
                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                End If
            End If
        Next
    Else
        ' render the shop items
        For I = 1 To MAX_TRADES
            itemNum = GetPlayerInvItemNum(MyIndex, I)
            
            ' draw early
            Top = yO + ShopTop + ((ShopOffsetY + 32) * ((I - 1) \ ShopColumns))
            Left = xO + ShopLeft + ((ShopOffsetX + 32) * (((I - 1) Mod ShopColumns)))
            ' draw selected square
            If shopSelectedSlot = I Then RenderTexture Tex_GUI(61), Left, Top, 0, 0, 32, 32, 32, 32
            
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
                ItemPic = Item(itemNum).Pic
                If ItemPic > 0 And ItemPic <= Count_Item Then

                    ' draw item
                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                    
                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, I) > 1 Then
                        Y = Top + 21
                        X = Left + 1
                        Amount = CStr(GetPlayerInvItemValue(MyIndex, I))
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            Colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            Colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            Colour = BrightGreen
                        End If
                        
                        RenderText font(Fonts.verdana_12), ConvertCurrency(Amount), X, Y, Colour
                    End If
                End If
            End If
        Next
    End If
End Sub

Sub DrawTrade()
    Dim xO As Long, yO As Long, Width As Long, Height As Long, I As Long, Y As Long, X As Long
    
    xO = Windows(GetWindowIndex("winTrade")).Window.Left
    yO = Windows(GetWindowIndex("winTrade")).Window.Top
    Width = Windows(GetWindowIndex("winTrade")).Window.Width
    Height = Windows(GetWindowIndex("winTrade")).Window.Height
    
    ' render green
    RenderTexture Tex_GUI(34), xO + 4, yO + 23, 0, 0, Width - 8, Height - 27, 4, 4
    
    ' top wood
    RenderTexture Tex_GUI(1), xO + 4, yO + 23, 100, 100, Width - 8, 18, Width - 8, 18
    ' left wood
    RenderTexture Tex_GUI(1), xO + 4, yO + 41, 350, 0, 5, Height - 45, 5, Height - 45
    ' right wood
    RenderTexture Tex_GUI(1), xO + Width - 9, yO + 41, 350, 0, 5, Height - 45, 5, Height - 45
    ' centre wood
    RenderTexture Tex_GUI(1), xO + 203, yO + 41, 350, 0, 6, Height - 45, 6, Height - 45
    ' bottom wood
    RenderTexture Tex_GUI(1), xO + 4, yO + 307, 100, 100, Width - 8, 75, Width - 8, 75
    
    ' left
    Width = 76
    Height = 76
    Y = yO + 41
    For I = 1 To 4
        If I = 4 Then Height = 38
        RenderTexture Tex_GUI(35), xO + 4 + 5, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 80 + 5, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 156 + 5, Y, 0, 0, 42, Height, 42, Height
        Y = Y + 76
    Next
    
    ' right
    Width = 76
    Height = 76
    Y = yO + 41
    For I = 1 To 4
        If I = 4 Then Height = 38
        RenderTexture Tex_GUI(35), xO + 4 + 205, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 80 + 205, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 156 + 205, Y, 0, 0, 42, Height, 42, Height
        Y = Y + 76
    Next
End Sub

Sub DrawYourTrade()
Dim I As Long, itemNum As Long, ItemPic As Long, Top As Long, Left As Long, Colour As Long, Amount As String, X As Long, Y As Long
Dim xO As Long, yO As Long

    xO = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).Left
    yO = Windows(GetWindowIndex("winTrade")).Window.Top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).Top
    
    ' your items
    For I = 1 To MAX_INV
        itemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).num)
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
            If ItemPic > 0 And ItemPic <= Count_Item Then
                Top = yO + TradeTop + ((TradeOffsetY + 32) * ((I - 1) \ TradeColumns))
                Left = xO + TradeLeft + ((TradeOffsetX + 32) * (((I - 1) Mod TradeColumns)))

                ' draw icon
                RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                
                ' If item is a stack - draw the amount you have
                If TradeYourOffer(I).value > 1 Then
                    Y = Top + 21
                    X = Left + 1
                    Amount = CStr(TradeYourOffer(I).value)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        Colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        Colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        Colour = BrightGreen
                    End If
                    
                    RenderText font(Fonts.verdana_12), ConvertCurrency(Amount), X, Y, Colour
                End If
            End If
        End If
    Next
End Sub

Sub DrawTheirTrade()
Dim I As Long, itemNum As Long, ItemPic As Long, Top As Long, Left As Long, Colour As Long, Amount As String, X As Long, Y As Long
Dim xO As Long, yO As Long

    xO = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picTheir")).Left
    yO = Windows(GetWindowIndex("winTrade")).Window.Top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picTheir")).Top

    ' their items
    For I = 1 To MAX_INV
        itemNum = TradeTheirOffer(I).num
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
            If ItemPic > 0 And ItemPic <= Count_Item Then
                Top = yO + TradeTop + ((TradeOffsetY + 32) * ((I - 1) \ TradeColumns))
                Left = xO + TradeLeft + ((TradeOffsetX + 32) * (((I - 1) Mod TradeColumns)))

                ' draw icon
                RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                
                ' If item is a stack - draw the amount you have
                If TradeTheirOffer(I).value > 1 Then
                    Y = Top + 21
                    X = Left + 1
                    Amount = CStr(TradeTheirOffer(I).value)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        Colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        Colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        Colour = BrightGreen
                    End If
                    
                    RenderText font(Fonts.verdana_12), ConvertCurrency(Amount), X, Y, Colour
                End If
            End If
        End If
    Next
End Sub

Public Sub DrawBank()
    Dim X As Long, Y As Long, xO As Long, yO As Long, Width As Long, Height As Long
    Dim I As Long, itemNum As Long, ItemPic As Long

    Dim Left As Long, Top As Long
    Dim Colour As Long, skipItem As Boolean, Amount As Long, tmpItem As Long

    xO = Windows(GetWindowIndex("winBank")).Window.Left
    yO = Windows(GetWindowIndex("winBank")).Window.Top
    Width = Windows(GetWindowIndex("winBank")).Window.Width
    Height = Windows(GetWindowIndex("winBank")).Window.Height
    
    ' render green
    RenderTexture Tex_GUI(34), xO + 4, yO + 23, 0, 0, Width - 8, Height - 27, 4, 4

    Width = 76
    Height = 76

    Y = yO + 23
    ' render grid - row
    For I = 1 To 5
        If I = 5 Then Height = 42
        RenderTexture Tex_GUI(35), xO + 4, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 80, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 156, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 232, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 308, Y, 0, 0, 79, Height, 79, Height
        Y = Y + 76
    Next

    ' actually draw the icons
    For I = 1 To MAX_BANK
        itemNum = Bank.Item(I).num

        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ' not dragging?
            If Not (DragBox.Origin = origin_Bank And DragBox.Slot = I) Then
                ItemPic = Item(itemNum).Pic


                If ItemPic > 0 And ItemPic <= Count_Item Then
                    Top = yO + BankTop + ((BankOffsetY + 32) * ((I - 1) \ BankColumns))
                    Left = xO + BankLeft + ((BankOffsetX + 32) * (((I - 1) Mod BankColumns)))

                    ' draw icon
                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32

                    ' If item is a stack - draw the amount you have
                    If Bank.Item(I).value > 1 Then
                        Y = Top + 21
                        X = Left + 1
                        Amount = Bank.Item(I).value

                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            Colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            Colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            Colour = BrightGreen
                        End If

                        RenderText font(Fonts.rockwell_15), ConvertCurrency(Amount), X, Y, Colour
                    End If
                End If
            End If
        End If
    Next

End Sub

Public Sub DrawInventory()
    Dim xO As Long, yO As Long, Width As Long, Height As Long, I As Long, Y As Long, itemNum As Long, ItemPic As Long, X As Long, Top As Long, Left As Long, Amount As String
    Dim Colour As Long, skipItem As Boolean, amountModifier  As Long, tmpItem As Long
    
    xO = Windows(GetWindowIndex("winInventory")).Window.Left
    yO = Windows(GetWindowIndex("winInventory")).Window.Top
    Width = Windows(GetWindowIndex("winInventory")).Window.Width
    Height = Windows(GetWindowIndex("winInventory")).Window.Height
    
    ' render green
    RenderTexture Tex_GUI(34), xO + 4, yO + 23, 0, 0, Width - 8, Height - 27, 4, 4
    
    Width = 76
    Height = 76
    
    Y = yO + 23
    ' render grid - row
    For I = 1 To 4
        If I = 4 Then Height = 38
        RenderTexture Tex_GUI(35), xO + 4, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 80, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 156, Y, 0, 0, 42, Height, 42, Height
        Y = Y + 76
    Next
    ' render bottom wood
    RenderTexture Tex_GUI(1), xO + 4, yO + 289, 100, 100, 194, 26, 194, 26
    
    ' actually draw the icons
    For I = 1 To MAX_INV
        itemNum = GetPlayerInvItemNum(MyIndex, I)
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ' not dragging?
            If Not (DragBox.Origin = origin_Inventory And DragBox.Slot = I) Then
                ItemPic = Item(itemNum).Pic
                
                ' exit out if we're offering item in a trade.
                amountModifier = 0
                If InTrade > 0 Then
                    For X = 1 To MAX_INV
                        tmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(X).num)
                        If TradeYourOffer(X).num = I Then
                            ' check if currency
                            If Not Item(tmpItem).Type = ITEM_TYPE_CURRENCY Then
                                ' normal item, exit out
                                skipItem = True
                            Else
                                ' if amount = all currency, remove from inventory
                                If TradeYourOffer(X).value = GetPlayerInvItemValue(MyIndex, I) Then
                                    skipItem = True
                                Else
                                    ' not all, change modifier to show change in currency count
                                    amountModifier = TradeYourOffer(X).value
                                End If
                            End If
                        End If
                    Next
                End If
                
                If Not skipItem Then
                    If ItemPic > 0 And ItemPic <= Count_Item Then
                        Top = yO + InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                        Left = xO + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
        
                        ' draw icon
                        RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
        
                        ' If item is a stack - draw the amount you have
                        If GetPlayerInvItemValue(MyIndex, I) > 1 Then
                            Y = Top + 21
                            X = Left + 1
                            Amount = GetPlayerInvItemValue(MyIndex, I) - amountModifier
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            If CLng(Amount) < 1000000 Then
                                Colour = White
                            ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                                Colour = Yellow
                            ElseIf CLng(Amount) > 10000000 Then
                                Colour = BrightGreen
                            End If
                            
                            RenderText font(Fonts.verdana_12), ConvertCurrency(Amount), X, Y, Colour
                        End If
                    End If
                End If
                ' reset
                skipItem = False
            End If
        End If
    Next
End Sub

Public Sub DrawChatBubble(ByVal Index As Long)
    Dim theArray() As String, X As Long, Y As Long, I As Long, MaxWidth As Long, X2 As Long, Y2 As Long, Colour As Long, tmpNum As Long
    
    With chatBubble(Index)
        ' exit out early
        If .target = 0 Then Exit Sub
        ' calculate position
        Select Case .TargetType
            Case TARGET_TYPE_PLAYER
                ' it's a player
                If Not GetPlayerMap(.target) = GetPlayerMap(MyIndex) Then Exit Sub
                ' change the colour depending on access
                Colour = DarkBrown
                ' it's on our map - get co-ords
                X = ConvertMapX((Player(.target).X * 32) + Player(.target).xOffset) + 16
                Y = ConvertMapY((Player(.target).Y * 32) + Player(.target).yOffset) - 32
            Case TARGET_TYPE_EVENT
                Colour = .Colour
                X = ConvertMapX(map.TileData.Events(.target).X * 32) + 16
                Y = ConvertMapY(map.TileData.Events(.target).Y * 32) - 16
            Case Else
                Exit Sub
        End Select
        
        ' word wrap
        WordWrap_Array .Msg, ChatBubbleWidth, theArray
        ' find max width
        tmpNum = UBound(theArray)

        For I = 1 To tmpNum
            If TextWidth(font(Fonts.georgiaDec_16), theArray(I)) > MaxWidth Then MaxWidth = TextWidth(font(Fonts.georgiaDec_16), theArray(I))
        Next

        ' calculate the new position
        X2 = X - (MaxWidth \ 2)
        Y2 = Y - (UBound(theArray) * 12)
        ' render bubble - top left
        RenderTexture Tex_GUI(33), X2 - 9, Y2 - 5, 0, 0, 9, 5, 9, 5
        ' top right
        RenderTexture Tex_GUI(33), X2 + MaxWidth, Y2 - 5, 119, 0, 9, 5, 9, 5
        ' top
        RenderTexture Tex_GUI(33), X2, Y2 - 5, 9, 0, MaxWidth, 5, 5, 5
        ' bottom left
        RenderTexture Tex_GUI(33), X2 - 9, Y, 0, 19, 9, 6, 9, 6
        ' bottom right
        RenderTexture Tex_GUI(33), X2 + MaxWidth, Y, 119, 19, 9, 6, 9, 6
        ' bottom - left half
        RenderTexture Tex_GUI(33), X2, Y, 9, 19, (MaxWidth \ 2) - 5, 6, 9, 6
        ' bottom - right half
        RenderTexture Tex_GUI(33), X2 + (MaxWidth \ 2) + 6, Y, 9, 19, (MaxWidth \ 2) - 5, 6, 9, 6
        ' left
        RenderTexture Tex_GUI(33), X2 - 9, Y2, 0, 6, 9, (UBound(theArray) * 12), 9, 1
        ' right
        RenderTexture Tex_GUI(33), X2 + MaxWidth, Y2, 119, 6, 9, (UBound(theArray) * 12), 9, 1
        ' center
        RenderTexture Tex_GUI(33), X2, Y2, 9, 5, MaxWidth, (UBound(theArray) * 12), 1, 1
        ' little pointy bit
        RenderTexture Tex_GUI(33), X - 5, Y, 58, 19, 11, 11, 11, 11
        ' render each line centralised
        tmpNum = UBound(theArray)

        For I = 1 To tmpNum
            RenderText font(Fonts.georgia_16), theArray(I), X - (TextWidth(font(Fonts.georgiaDec_16), theArray(I)) / 2), Y2, Colour
            Y2 = Y2 + 12
        Next

        ' check if it's timed out - close it if so
        If .timer + 5000 < GetTickCount Then
            .active = False
        End If
    End With
End Sub

Public Function isConstAnimated(ByVal sprite As Long) As Boolean
    isConstAnimated = False

    Select Case sprite

        Case 16, 21, 22, 26, 28
            isConstAnimated = True
    End Select

End Function

Public Function hasSpriteShadow(ByVal sprite As Long) As Boolean
    hasSpriteShadow = True

    Select Case sprite

        Case 25, 26
            hasSpriteShadow = False
    End Select

End Function

Public Sub DrawPlayer(ByVal Index As Long)
    Dim Anim As Byte
    Dim X As Long
    Dim Y As Long
    Dim sprite As Long, spritetop As Long
    Dim rec As GeomRec
    Dim attackspeed As Long
    
    ' pre-load sprite for calculations
    sprite = GetPlayerSprite(Index)

    'SetTexture Tex_Char(Sprite)
    If sprite < 1 Or sprite > Count_Char Then Exit Sub

    ' speed from weapon
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(Index, Weapon)).speed
    Else
        attackspeed = 1000
    End If

    If Not isConstAnimated(GetPlayerSprite(Index)) Then
        ' Reset frame
        Anim = 1

        ' Check for attacking animation
        If Player(Index).AttackTimer + (attackspeed / 2) > GetTickCount Then
            If Player(Index).Attacking = 1 Then
                Anim = 2
            End If

        Else

            ' If not attacking, walk normally
            Select Case GetPlayerDir(Index)

                Case DIR_UP

                    If (Player(Index).yOffset > 8) Then Anim = Player(Index).Step

                Case DIR_DOWN

                    If (Player(Index).yOffset < -8) Then Anim = Player(Index).Step

                Case DIR_LEFT

                    If (Player(Index).xOffset > 8) Then Anim = Player(Index).Step

                Case DIR_RIGHT

                    If (Player(Index).xOffset < -8) Then Anim = Player(Index).Step
            End Select

        End If

    Else

        If Player(Index).AnimTimer + 100 <= GetTickCount Then
            Player(Index).Anim = Player(Index).Anim + 1

            If Player(Index).Anim >= 3 Then Player(Index).Anim = 0
            Player(Index).AnimTimer = GetTickCount
        End If

        Anim = Player(Index).Anim
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)

        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If

    End With

    ' Set the left
    Select Case GetPlayerDir(Index)

        Case DIR_UP
            spritetop = 3

        Case DIR_RIGHT
            spritetop = 2

        Case DIR_DOWN
            spritetop = 0

        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = spritetop * (mTexture(Tex_Char(sprite)).h / 4)
        .Height = (mTexture(Tex_Char(sprite)).h / 4)
        .Left = Anim * (mTexture(Tex_Char(sprite)).w / 4)
        .Width = (mTexture(Tex_Char(sprite)).w / 4)
    End With

    ' Calculate the X
    X = GetPlayerX(Index) * PIC_X + Player(Index).xOffset - ((mTexture(Tex_Char(sprite)).w / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (mTexture(Tex_Char(sprite)).h) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - ((mTexture(Tex_Char(sprite)).h / 4) - 32) - 4
    Else
        ' Proceed as normal
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - 4
    End If

    RenderTexture Tex_Char(sprite), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Width, rec.Height, rec.Width, rec.Height
End Sub

Public Sub DrawNpc(ByVal MapNpcNum As Long)
    Dim Anim As Byte
    Dim X As Long
    Dim Y As Long
    Dim sprite As Long, spritetop As Long
    Dim rec As GeomRec
    Dim attackspeed As Long

    If MapNpc(MapNpcNum).num = 0 Then Exit Sub ' no npc set
    ' pre-load texture for calculations
    sprite = Npc(MapNpc(MapNpcNum).num).sprite

    'SetTexture Tex_Char(Sprite)
    If sprite < 1 Or sprite > Count_Char Then Exit Sub
    attackspeed = 1000

    If Not isConstAnimated(Npc(MapNpc(MapNpcNum).num).sprite) Then
        ' Reset frame
        Anim = 1

        ' Check for attacking animation
        If MapNpc(MapNpcNum).AttackTimer + (attackspeed / 2) > GetTickCount Then
            If MapNpc(MapNpcNum).Attacking = 1 Then
                Anim = 2
            End If

        Else

            ' If not attacking, walk normally
            Select Case MapNpc(MapNpcNum).Dir

                Case DIR_UP

                    If (MapNpc(MapNpcNum).yOffset > 8) Then Anim = MapNpc(MapNpcNum).Step

                Case DIR_DOWN

                    If (MapNpc(MapNpcNum).yOffset < -8) Then Anim = MapNpc(MapNpcNum).Step

                Case DIR_LEFT

                    If (MapNpc(MapNpcNum).xOffset > 8) Then Anim = MapNpc(MapNpcNum).Step

                Case DIR_RIGHT

                    If (MapNpc(MapNpcNum).xOffset < -8) Then Anim = MapNpc(MapNpcNum).Step
            End Select

        End If

    Else

        With MapNpc(MapNpcNum)

            If .AnimTimer + 100 <= GetTickCount Then
                .Anim = .Anim + 1

                If .Anim >= 3 Then .Anim = 0
                .AnimTimer = GetTickCount
            End If

            Anim = .Anim
        End With

    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)

        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If

    End With

    ' Set the left
    Select Case MapNpc(MapNpcNum).Dir

        Case DIR_UP
            spritetop = 3

        Case DIR_RIGHT
            spritetop = 2

        Case DIR_DOWN
            spritetop = 0

        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = (mTexture(Tex_Char(sprite)).h / 4) * spritetop
        .Height = mTexture(Tex_Char(sprite)).h / 4
        .Left = Anim * (mTexture(Tex_Char(sprite)).w / 4)
        .Width = (mTexture(Tex_Char(sprite)).w / 4)
    End With

    ' Calculate the X
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).xOffset - ((mTexture(Tex_Char(sprite)).w / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (mTexture(Tex_Char(sprite)).h / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).yOffset - ((mTexture(Tex_Char(sprite)).h / 4) - 32) - 4
    Else
        ' Proceed as normal
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).yOffset - 4
    End If

    RenderTexture Tex_Char(sprite), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Width, rec.Height, rec.Width, rec.Height
End Sub

Sub DrawEvent(eventNum As Long, pageNum As Long)
Dim texNum As Long, X As Long, Y As Long

    ' render it
    With map.TileData.Events(eventNum).EventPage(pageNum)
        If .GraphicType > 0 Then
            If .Graphic > 0 Then
                Select Case .GraphicType
                    Case 1 ' character
                        If .Graphic < Count_Char Then
                            texNum = Tex_Char(.Graphic)
                        End If
                    Case 2 ' tileset
                        If .Graphic < Count_Tileset Then
                            texNum = Tex_Tileset(.Graphic)
                        End If
                End Select
                If texNum > 0 Then
                    X = ConvertMapX(map.TileData.Events(eventNum).X * 32)
                    Y = ConvertMapY(map.TileData.Events(eventNum).Y * 32)
                    RenderTexture texNum, X, Y, .GraphicX * 32, .GraphicY * 32, 32, 32, 32, 32
                End If
            End If
        End If
    End With
End Sub

Sub DrawLowerEvents()
Dim I As Long, X As Long

    If map.TileData.EventCount = 0 Then Exit Sub
    For I = 1 To map.TileData.EventCount
        ' find the active page
        If map.TileData.Events(I).pageCount > 0 Then
            X = ActiveEventPage(I)
            If X > 0 Then
                ' make sure it's lower
                If map.TileData.Events(I).EventPage(X).Priority <> 2 Then
                    ' render event
                    DrawEvent I, X
                End If
            End If
        End If
    Next
End Sub

Sub DrawUpperEvents()
Dim I As Long, X As Long

    If map.TileData.EventCount = 0 Then Exit Sub
    For I = 1 To map.TileData.EventCount
        ' find the active page
        If map.TileData.Events(I).pageCount > 0 Then
            X = ActiveEventPage(I)
            If X > 0 Then
                ' make sure it's lower
                If map.TileData.Events(I).EventPage(X).Priority = 2 Then
                    ' render event
                    DrawEvent I, X
                End If
            End If
        End If
    Next
End Sub

Public Sub DrawShadow(ByVal sprite As Long, ByVal X As Long, ByVal Y As Long)
    If hasSpriteShadow(sprite) Then RenderTexture Tex_Shadow, ConvertMapX(X), ConvertMapY(Y), 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawTarget(ByVal X As Long, ByVal Y As Long)
    Dim Width As Long, Height As Long
    ' calculations
    Width = mTexture(Tex_Target).w / 2
    Height = mTexture(Tex_Target).h
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2) + 16
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    'EngineRenderRectangle Tex_Target, x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_Target, X, Y, 0, 0, Width, Height, Width, Height
End Sub

Public Sub DrawTargetHover()
    Dim I As Long, X As Long, Y As Long, Width As Long, Height As Long

    If diaIndex > 0 Then Exit Sub
    Width = mTexture(Tex_Target).w / 2
    Height = mTexture(Tex_Target).h

    If Width <= 0 Then Width = 1
    If Height <= 0 Then Height = 1

    For I = 1 To MAX_PLAYERS

        If IsPlaying(I) And GetPlayerMap(MyIndex) = GetPlayerMap(I) Then
            X = (Player(I).X * 32) + Player(I).xOffset + 32
            Y = (Player(I).Y * 32) + Player(I).yOffset + 32

            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    X = ConvertMapX(X)
                    Y = ConvertMapY(Y)
                    RenderTexture Tex_Target, X - 16 - (Width / 2), Y - 16 - (Height / 2), Width, 0, Width, Height, Width, Height
                End If
            End If
        End If

    Next

    For I = 1 To MAX_MAP_NPCS

        If MapNpc(I).num > 0 Then
            X = (MapNpc(I).X * 32) + MapNpc(I).xOffset + 32
            Y = (MapNpc(I).Y * 32) + MapNpc(I).yOffset + 32

            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    X = ConvertMapX(X)
                    Y = ConvertMapY(Y)
                    RenderTexture Tex_Target, X - 16 - (Width / 2), Y - 16 - (Height / 2), Width, 0, Width, Height, Width, Height
                End If
            End If
        End If

    Next

End Sub

Public Sub DrawResource(ByVal Resource_num As Long)
    Dim Resource_master As Long
    Dim Resource_state As Long
    Dim Resource_sprite As Long
    Dim rec As RECT
    Dim X As Long, Y As Long
    Dim Width As Long, Height As Long
    X = MapResource(Resource_num).X
    Y = MapResource(Resource_num).Y

    If X < 0 Or X > map.MapData.MaxX Then Exit Sub
    If Y < 0 Or Y > map.MapData.MaxY Then Exit Sub
    ' Get the Resource type
    Resource_master = map.TileData.Tile(X, Y).Data1

    If Resource_master = 0 Then Exit Sub
    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' pre-load texture for calculations
    'SetTexture Tex_Resource(Resource_sprite)
    ' src rect
    With rec
        .Top = 0
        .bottom = mTexture(Tex_Resource(Resource_sprite)).h
        .Left = 0
        .Right = mTexture(Tex_Resource(Resource_sprite)).w
    End With

    ' Set base x + y, then the offset due to size
    X = (MapResource(Resource_num).X * PIC_X) - (mTexture(Tex_Resource(Resource_sprite)).w / 2) + 16
    Y = (MapResource(Resource_num).Y * PIC_Y) - mTexture(Tex_Resource(Resource_sprite)).h + 32
    Width = rec.Right - rec.Left
    Height = rec.bottom - rec.Top
    'EngineRenderRectangle Tex_Resource(Resource_sprite), ConvertMapX(x), ConvertMapY(y), 0, 0, width, height, width, height, width, height
    RenderTexture Tex_Resource(Resource_sprite), ConvertMapX(X), ConvertMapY(Y), 0, 0, Width, Height, Width, Height
End Sub

Public Sub DrawItem(ByVal itemNum As Long)
    Dim PicNum As Integer, dontRender As Boolean, I As Long, tmpIndex As Long
    PicNum = Item(MapItem(itemNum).num).Pic

    If PicNum < 1 Or PicNum > Count_Item Then Exit Sub

    ' if it's not us then don't render
    If MapItem(itemNum).playerName <> vbNullString Then
        If Trim$(MapItem(itemNum).playerName) <> Trim$(GetPlayerName(MyIndex)) Then

            dontRender = True
        End If

        ' make sure it's not a party drop
        If Party.Leader > 0 Then

            For I = 1 To MAX_PARTY_MEMBERS
                tmpIndex = Party.Member(I)

                If tmpIndex > 0 Then
                    If Trim$(GetPlayerName(tmpIndex)) = Trim$(MapItem(itemNum).playerName) Then
                        If MapItem(itemNum).bound = 0 Then

                            dontRender = False
                        End If
                    End If
                End If

            Next

        End If
    End If

    'If Not dontRender Then EngineRenderRectangle Tex_Item(PicNum), ConvertMapX(MapItem(itemnum).x * PIC_X), ConvertMapY(MapItem(itemnum).y * PIC_Y), 0, 0, 32, 32, 32, 32, 32, 32
    If Not dontRender Then
        RenderTexture Tex_Item(PicNum), ConvertMapX(MapItem(itemNum).X * PIC_X), ConvertMapY(MapItem(itemNum).Y * PIC_Y), 0, 0, 32, 32, 32, 32
    End If

End Sub

Public Sub DrawBars()
Dim Left As Long, Top As Long, Width As Long, Height As Long
Dim tmpX As Long, tmpY As Long, barWidth As Long, I As Long, NpcNum As Long
Dim partyIndex As Long

    ' dynamic bar calculations
    Width = mTexture(Tex_Bars).w
    Height = mTexture(Tex_Bars).h / 4
    
    ' render npc health bars
    For I = 1 To MAX_MAP_NPCS
        NpcNum = MapNpc(I).num
        ' exists?
        If NpcNum > 0 Then
            ' alive?
            If MapNpc(I).Vital(Vitals.HP) > 0 And MapNpc(I).Vital(Vitals.HP) < Npc(NpcNum).HP Then
                ' lock to npc
                tmpX = MapNpc(I).X * PIC_X + MapNpc(I).xOffset + 16 - (Width / 2)
                tmpY = MapNpc(I).Y * PIC_Y + MapNpc(I).yOffset + 35
                
                ' calculate the width to fill
                If Width > 0 Then BarWidth_NpcHP_Max(I) = ((MapNpc(I).Vital(Vitals.HP) / Width) / (Npc(NpcNum).HP / Width)) * Width
                
                ' draw bar background
                Top = Height * 1 ' HP bar background
                Left = 0
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, Width, Height, Width, Height
                
                ' draw the bar proper
                Top = 0 ' HP bar
                Left = 0
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, BarWidth_NpcHP(I), Height, BarWidth_NpcHP(I), Height
            End If
        End If
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer).Spell).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 16 - (Width / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).yOffset + 35 + Height + 1
            
            ' calculate the width to fill
            If Width > 0 Then barWidth = (GetTickCount - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer).Spell).CastTime * 1000)) * Width
            
            ' draw bar background
            Top = Height * 3 ' cooldown bar background
            Left = 0
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, Width, Height, Width, Height
             
            ' draw the bar proper
            Top = Height * 2 ' cooldown bar
            Left = 0
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, barWidth, Height, barWidth, Height
        End If
    End If
    
    ' draw own health bar
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 And GetPlayerVital(MyIndex, Vitals.HP) < GetPlayerMaxVital(MyIndex, Vitals.HP) Then
        ' lock to Player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 16 - (Width / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).yOffset + 35
       
        ' calculate the width to fill
        If Width > 0 Then BarWidth_PlayerHP_Max(MyIndex) = ((GetPlayerVital(MyIndex, Vitals.HP) / Width) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / Width)) * Width
       
        ' draw bar background
        Top = Height * 1 ' HP bar background
        Left = 0
        RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, Width, Height, Width, Height
       
        ' draw the bar proper
        Top = 0 ' HP bar
        Left = 0
        RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, BarWidth_PlayerHP(MyIndex), Height, BarWidth_PlayerHP(MyIndex), Height
    End If
End Sub

Public Sub DrawMenuBG()
    ' row 1
    RenderTexture Tex_Surface(1), ScreenWidth - 512, ScreenHeight - 512, 0, 0, 512, 512, 512, 512
    RenderTexture Tex_Surface(2), ScreenWidth - 1024, ScreenHeight - 512, 0, 0, 512, 512, 512, 512
    RenderTexture Tex_Surface(3), ScreenWidth - 1536, ScreenHeight - 512, 0, 0, 512, 512, 512, 512
    RenderTexture Tex_Surface(4), ScreenWidth - 2048, ScreenHeight - 512, 0, 0, 512, 512, 512, 512
    ' row 2
    RenderTexture Tex_Surface(5), ScreenWidth - 512, ScreenHeight - 1024, 0, 0, 512, 512, 512, 512
    RenderTexture Tex_Surface(6), ScreenWidth - 1024, ScreenHeight - 1024, 0, 0, 512, 512, 512, 512
    RenderTexture Tex_Surface(7), ScreenWidth - 1536, ScreenHeight - 1024, 0, 0, 512, 512, 512, 512
    RenderTexture Tex_Surface(8), ScreenWidth - 2048, ScreenHeight - 1024, 0, 0, 512, 512, 512, 512
    ' row 3
    RenderTexture Tex_Surface(9), ScreenWidth - 512, ScreenHeight - 1088, 0, 0, 512, 64, 512, 64
    RenderTexture Tex_Surface(10), ScreenWidth - 1024, ScreenHeight - 1088, 0, 0, 512, 64, 512, 64
    RenderTexture Tex_Surface(11), ScreenWidth - 1536, ScreenHeight - 1088, 0, 0, 512, 64, 512, 64
    RenderTexture Tex_Surface(12), ScreenWidth - 2048, ScreenHeight - 1088, 0, 0, 512, 64, 512, 64
End Sub

Public Sub DrawAnimation(ByVal Index As Long, ByVal Layer As Long)
    Dim sprite As Integer, sRECT As GeomRec, Width As Long, Height As Long, FrameCount As Long
    Dim X As Long, Y As Long, lockindex As Long

    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If

    sprite = Animation(AnimInstance(Index).Animation).sprite(Layer)

    If sprite < 1 Or sprite > Count_Anim Then Exit Sub
    ' pre-load texture for calculations
    'SetTexture Tex_Anim(Sprite)
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    ' total width divided by frame count
    Width = 192 'mTexture(Tex_Anim(Sprite)).width / frameCount
    Height = 192 'mTexture(Tex_Anim(Sprite)).height

    With sRECT
        .Top = (Height * ((AnimInstance(Index).FrameIndex(Layer) - 1) \ AnimColumns))
        .Height = Height
        .Left = (Width * (((AnimInstance(Index).FrameIndex(Layer) - 1) Mod AnimColumns)))
        .Width = Width
    End With

    ' change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' if <> none

        ' is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex

            ' check if is ingame
            If IsPlaying(lockindex) Then

                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    X = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + Player(lockindex).xOffset
                    Y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + Player(lockindex).yOffset
                End If
            End If

        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex

            ' check if NPC exists
            If MapNpc(lockindex).num > 0 Then

                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    X = (MapNpc(lockindex).X * PIC_X) + 16 - (Width / 2) + MapNpc(lockindex).xOffset
                    Y = (MapNpc(lockindex).Y * PIC_Y) + 16 - (Height / 2) + MapNpc(lockindex).yOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If

            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If

    Else
        ' no lock, default x + y
        X = (AnimInstance(Index).X * 32) + 16 - (Width / 2)
        Y = (AnimInstance(Index).Y * 32) + 16 - (Height / 2)
    End If

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    'EngineRenderRectangle Tex_Anim(sprite), x, y, sRECT.left, sRECT.top, sRECT.width, sRECT.height, sRECT.width, sRECT.height, sRECT.width, sRECT.height
    RenderTexture Tex_Anim(sprite), X, Y, sRECT.Left, sRECT.Top, sRECT.Width, sRECT.Height, sRECT.Width, sRECT.Height
End Sub

Public Sub DrawGDI()

    If frmEditor_Animation.visible Then
        GDIRenderAnimation
    ElseIf frmEditor_Item.visible Then
        GDIRenderItem frmEditor_Item.picItem, frmEditor_Item.scrlPic.value
    ElseIf frmEditor_Map.visible Then
        GDIRenderTileset
        If frmEditor_Events.visible Then
            GDIRenderEventGraphic
            GDIRenderEventGraphicSel
        End If
    ElseIf frmEditor_NPC.visible Then
        GDIRenderChar frmEditor_NPC.picSprite, frmEditor_NPC.scrlSprite.value
    ElseIf frmEditor_Resource.visible Then
        GDIRenderResource frmEditor_Resource.picNormalPic, frmEditor_Resource.scrlNormalPic.value
        GDIRenderResource frmEditor_Resource.picExhaustedPic, frmEditor_Resource.scrlExhaustedPic.value
    ElseIf frmEditor_Spell.visible Then
        GDIRenderSpell frmEditor_Spell.picSprite, frmEditor_Spell.scrlIcon.value
    End If

End Sub

' Main Loop
Public Sub Render_Graphics()
    Dim X As Long, Y As Long, I As Long, bgColour As Long
    ' fuck off if we're not doing anything
    If GettingMap Then Exit Sub
    
    ' update the camera
    UpdateCamera
    
    ' check graphics
    CheckGFX

    ' Start rendering
    If Not InMapEditor Then
        bgColour = 0
    Else
        bgColour = DX8Colour(Red, 255)
    End If
    
    ' Bg
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, bgColour, 1#, 0)
    Call D3DDevice.BeginScene
    
    ' render black if map
    If InMapEditor Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.bottom
                If IsValidMapPoint(X, Y) Then
                    RenderTexture Tex_Fader, ConvertMapX(X * 32), ConvertMapY(Y * 32), 0, 0, 32, 32, 32, 32
                End If
            Next
        Next
    End If
    
    ' Render appear tile fades
    'RenderAppearTileFade

    ' render lower tiles
    If Count_Tileset > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.bottom
                If IsValidMapPoint(X, Y) Then
                    Call DrawMapTile(X, Y)
                End If
            Next
        Next
    End If

    ' render the items
    If Count_Item > 0 Then
        For I = 1 To MAX_MAP_ITEMS
            If MapItem(I).num > 0 Then
                Call DrawItem(I)
            End If
        Next
    End If

    ' draw animations
    If Count_Anim > 0 Then
        For I = 1 To MAX_BYTE
            If AnimInstance(I).Used(0) Then
                DrawAnimation I, 0
            End If
        Next
    End If
    
    ' draw events
    DrawLowerEvents

    ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
    For Y = TileView.Top To TileView.bottom
        ' Resources
        If Count_Resource > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For I = 1 To Resource_Index
                        If MapResource(I).Y = Y Then
                            Call DrawResource(I)
                        End If
                    Next
                End If
            End If
        End If
        
        If Count_Char > 0 Then
            ' shadows - Players
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                    If Player(I).Y = Y Then
                        Call DrawShadow(Player(I).sprite, (Player(I).X * 32) + Player(I).xOffset, (Player(I).Y * 32) + Player(I).yOffset)
                    End If
                End If
            Next
    
            ' shadows - npcs
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(I).num > 0 Then
                    If MapNpc(I).Y = Y Then
                        Call DrawShadow(Npc(MapNpc(I).num).sprite, (MapNpc(I).X * 32) + MapNpc(I).xOffset, (MapNpc(I).Y * 32) + MapNpc(I).yOffset)
                    End If
                End If
            Next
    
            ' Players
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                    If Player(I).Y = Y Then
                        Call DrawPlayer(I)
                    End If
                End If
            Next
    
            ' Npcs
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(I).Y = Y Then
                    Call DrawNpc(I)
                End If
            Next
        End If
    Next Y

    ' render out upper tiles
    If Count_Tileset > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.bottom
                If IsValidMapPoint(X, Y) Then
                    Call DrawMapFringeTile(X, Y)
                End If
            Next
        Next
    End If
    
    ' draw events
    DrawUpperEvents

    ' render fog
    DrawWeather
    DrawFog
    DrawTint

    ' render animations
    If Count_Anim > 0 Then
        For I = 1 To MAX_BYTE
            If AnimInstance(I).Used(1) Then
                DrawAnimation I, 1
            End If
        Next
    End If

    ' render target
    If myTarget > 0 Then
        If myTargetType = TARGET_TYPE_PLAYER Then
            DrawTarget (Player(myTarget).X * 32) + Player(myTarget).xOffset, (Player(myTarget).Y * 32) + Player(myTarget).yOffset
        ElseIf myTargetType = TARGET_TYPE_NPC Then
            DrawTarget (MapNpc(myTarget).X * 32) + MapNpc(myTarget).xOffset, (MapNpc(myTarget).Y * 32) + MapNpc(myTarget).yOffset
        End If
    End If

    ' blt the hover icon
    DrawTargetHover
    
    ' draw the bars
    DrawBars

    ' draw attributes
    If InMapEditor Then
        DrawMapAttributes
        DrawMapEvents
    End If

    ' draw player names
    If Not screenshotMode Then
        For I = 1 To MAX_PLAYERS
            If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                Call DrawPlayerName(I)
            End If
        Next
    End If

    ' draw npc names
    If Not screenshotMode Then
        For I = 1 To MAX_MAP_NPCS
            If MapNpc(I).num > 0 Then
                Call DrawNpcName(I)
            End If
        Next
    End If

    ' draw action msg
    For I = 1 To MAX_BYTE
        DrawActionMsg I
    Next

    If InMapEditor Then
        If frmEditor_Map.optBlock.value = True Then
            For X = TileView.Left To TileView.Right
                For Y = TileView.Top To TileView.bottom
                    If IsValidMapPoint(X, Y) Then
                        Call DrawDirection(X, Y)
                    End If
                Next
            Next
        End If
    End If

    ' draw the messages
    For I = 1 To MAX_BYTE
        If chatBubble(I).active Then
            DrawChatBubble I
        End If
    Next
    
    If DrawThunder > 0 Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, 160): DrawThunder = DrawThunder - 1
    
    ' draw shadow
    If Not screenshotMode Then
        RenderTexture Tex_GUI(43), 0, 0, 0, 0, ScreenWidth, 64, 1, 64
        RenderTexture Tex_GUI(42), 0, ScreenHeight - 64, 0, 0, ScreenWidth, 64, 1, 64
    End If
    
    ' Render entities
    If Not InMapEditor And Not hideGUI And Not screenshotMode Then RenderEntities
    
    ' render the tile selection
    If InMapEditor Then DrawTileSelection
  
    ' render FPS
    If Not screenshotMode Then RenderText font(Fonts.rockwell_15), "FPS: " & GameFPS, 1, 1, White

    ' draw loc
    If BLoc Then
        RenderText font(Fonts.georgiaDec_16), Trim$("cur x: " & CurX & " y: " & CurY), 260, 6, Yellow
        RenderText font(Fonts.georgiaDec_16), Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), 260, 22, Yellow
        RenderText font(Fonts.georgiaDec_16), Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), 260, 38, Yellow
    End If
    
    ' draw map name
    RenderMapName

    ' End the rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(ByVal 0, ByVal 0, 0, ByVal 0)
    ' GDI Rendering
    DrawGDI
End Sub

Public Sub Render_Menu()
    ' check graphics
    CheckGFX
    ' Start rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, &HFFFFFF, 1#, 0)
    Call D3DDevice.BeginScene
    ' Render menu background
    DrawMenuBG
    ' Render entities
    RenderEntities
    ' render white fade
    DrawFade
    ' End the rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(ByVal 0, ByVal 0, 0, ByVal 0)
End Sub
