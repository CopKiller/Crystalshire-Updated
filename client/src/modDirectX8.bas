Attribute VB_Name = "modDirectX8"
Option Explicit

' Texture paths
Public Const Path_Anim As String = "\data files\graphics\animations\"
Public Const Path_Char As String = "\data files\graphics\characters\"
Public Const Path_Face As String = "\data files\graphics\faces\"
Public Const Path_GUI As String = "\data files\graphics\gui\"
Public Const Path_Design As String = "\data files\graphics\gui\designs\"
Public Const Path_Gradient As String = "\data files\graphics\gui\gradients\"
Public Const Path_Item As String = "\data files\graphics\items\"
Public Const Path_Paperdoll As String = "\data files\graphics\paperdolls\"
Public Const Path_Resource As String = "\data files\graphics\resources\"
Public Const Path_Spellicon As String = "\data files\graphics\spellicons\"
Public Const Path_Tileset As String = "\data files\graphics\tilesets\"
Public Const Path_Font As String = "\data files\graphics\fonts\"
Public Const Path_Graphics As String = "\data files\graphics\"
Public Const Path_Surface As String = "\data files\graphics\surfaces\"
Public Const Path_Fog As String = "\data files\graphics\fog\"
Public Const Path_Captcha As String = "\data files\graphics\captchas\"

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
    Data() As Byte
    w As Long
    h As Long
End Type

Public Type TextureDataStruct
    Data() As Byte
End Type

Public Type Vertex
    x As Single
    y As Single
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
Public curResolution As Byte, isFullscreen As Boolean

Public Sub InitDX8(ByVal hwnd As Long)
Dim DispMode As D3DDISPLAYMODE, Width As Long, Height As Long

    mhWnd = hwnd

    Set DX8 = New DirectX8
    Set D3D = DX8.Direct3DCreate
    Set D3DX = New D3DX8
    
    ' set size
    GetResolutionSize curResolution, Width, Height
    ScreenWidth = Width
    ScreenHeight = Height
    TileWidth = (ScreenWidth / 32)
    TileHeight = (ScreenHeight / 32)
    ScreenX = (TileWidth + 1) * PIC_X
    ScreenY = (TileHeight + 1) * PIC_Y
    
    ' set up window
    Call D3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DispMode)
    DispMode.Format = D3DFMT_A8R8G8B8
    
    If Options.Fullscreen = 0 Then
        isFullscreen = False
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
        D3DWindow.hDeviceWindow = hwnd
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
            If LoadDirectX(D3DCREATE_HARDWARE_VERTEXPROCESSING, hwnd) <> 0 Then
                Options.Fullscreen = 0
                Options.Resolution = 0
                Options.Render = 0
                SaveOptions
                Call MsgBox("Could not initialize DirectX with hardware vertex processing.", vbCritical)
                Call DestroyGame
            End If
        Case 2 ' mixed
            If LoadDirectX(D3DCREATE_MIXED_VERTEXPROCESSING, hwnd) <> 0 Then
                Options.Fullscreen = 0
                Options.Resolution = 0
                Options.Render = 0
                SaveOptions
                Call MsgBox("Could not initialize DirectX with mixed vertex processing.", vbCritical)
                Call DestroyGame
            End If
        Case 3 ' software
            If LoadDirectX(D3DCREATE_SOFTWARE_VERTEXPROCESSING, hwnd) <> 0 Then
                Options.Fullscreen = 0
                Options.Resolution = 0
                Options.Render = 0
                SaveOptions
                Call MsgBox("Could not initialize DirectX with software vertex processing.", vbCritical)
                Call DestroyGame
            End If
        Case Else ' auto
            If LoadDirectX(D3DCREATE_HARDWARE_VERTEXPROCESSING, hwnd) <> 0 Then
                If LoadDirectX(D3DCREATE_MIXED_VERTEXPROCESSING, hwnd) <> 0 Then
                    If LoadDirectX(D3DCREATE_SOFTWARE_VERTEXPROCESSING, hwnd) <> 0 Then
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

Public Function LoadDirectX(ByVal BehaviourFlags As CONST_D3DCREATEFLAGS, ByVal hwnd As Long)
On Error GoTo ErrorInit

    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, BehaviourFlags, D3DWindow)
    Exit Function

ErrorInit:
    LoadDirectX = 1
End Function

Sub DestroyDX8()
Dim i As Long
    'For i = 1 To mTextures
    '    mTexture(i).data
    'Next
    If Not DX8 Is Nothing Then Set DX8 = Nothing
    If Not D3D Is Nothing Then Set D3D = Nothing
    If Not D3DX Is Nothing Then Set D3DX = Nothing
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
End Sub

Public Sub LoadTextures()
Dim i As Long
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
    Tex_Fog = LoadTextureFiles(Count_Fog, App.path & Path_Fog)
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
Dim i As Long

    Counter = 1
    
    Do While Dir$(path & Counter + 1 & ".png") <> vbNullString
        Counter = Counter + 1
    Loop
    
    ReDim Texture(0 To Counter)
    
    For i = 1 To Counter
        Texture(i) = LoadTextureFile(path & i & ".png")
        DoEvents
    Next
    
    LoadTextureFiles = Texture
End Function

Public Function LoadTextureFile(ByVal path As String, Optional ByVal DontReuse As Boolean) As Long
Dim Data() As Byte
Dim f As Long

    If Dir$(path) = vbNullString Then
        Call MsgBox("""" & path & """ could not be found.")
        End
    End If
    
    f = FreeFile
    Open path For Binary As #f
        ReDim Data(0 To LOF(f) - 1)
        Get #f, , Data
    Close #f
    
    LoadTextureFile = LoadTexture(Data, DontReuse)
End Function

Public Function LoadTexture(ByRef Data() As Byte, Optional ByVal DontReuse As Boolean) As Long
Dim i As Long

    If AryCount(Data) = 0 Then
        Exit Function
    End If
    
    mTextures = mTextures + 1
    LoadTexture = mTextures
    ReDim Preserve mTexture(1 To mTextures) As TextureStruct
    mTexture(mTextures).w = ByteToInt(Data(18), Data(19))
    mTexture(mTextures).h = ByteToInt(Data(22), Data(23))
    mTexture(mTextures).Data = Data
    Set mTexture(mTextures).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Data(0), AryCount(Data), mTexture(mTextures).w, mTexture(mTextures).h, D3DX_DEFAULT, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
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
Dim i As Long, N As Long

    N = mTextures
    ReDim Temp(1 To N)
    For i = 1 To N
        Set mTexture(i).Texture = Nothing
        Temp(i).Data = mTexture(i).Data
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
    
    For i = 1 To N
        Call LoadTexture(Temp(i).Data)
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

Public Sub RenderTexture(Texture As Long, ByVal x As Long, ByVal y As Long, ByVal sX As Single, ByVal sY As Single, ByVal w As Long, ByVal h As Long, ByVal sW As Single, ByVal sH As Single, Optional ByVal Colour As Long = -1, Optional ByVal offset As Boolean = False)
    SetTexture Texture
    RenderGeom x, y, sX, sY, w, h, sW, sH, Colour, offset
End Sub

Public Sub RenderGeom(ByVal x As Long, ByVal y As Long, ByVal sX As Single, ByVal sY As Single, ByVal w As Long, ByVal h As Long, ByVal sW As Single, ByVal sH As Single, Optional ByVal Colour As Long = -1, Optional ByVal offset As Boolean = False)
Dim i As Long

    If CurrentTexture = 0 Then Exit Sub
    If w = 0 Then Exit Sub
    If h = 0 Then Exit Sub
    If sW = 0 Then Exit Sub
    If sH = 0 Then Exit Sub
    
    If mClip.Right <> 0 Then
        If mClip.Top <> 0 Then
            If mClip.Left > x Then
                sX = sX + (mClip.Left - x) / (w / sW)
                sW = sW - (mClip.Left - x) / (w / sW)
                w = w - (mClip.Left - x)
                x = mClip.Left
            End If
            
            If mClip.Top > y Then
                sY = sY + (mClip.Top - y) / (h / sH)
                sH = sH - (mClip.Top - y) / (h / sH)
                h = h - (mClip.Top - y)
                y = mClip.Top
            End If
            
            If mClip.Right < x + w Then
                sW = sW - (x + w - mClip.Right) / (w / sW)
                w = -x + mClip.Right
            End If
            
            If mClip.bottom < y + h Then
                sH = sH - (y + h - mClip.bottom) / (h / sH)
                h = -y + mClip.bottom
            End If
            
            If w <= 0 Then Exit Sub
            If h <= 0 Then Exit Sub
            If sW <= 0 Then Exit Sub
            If sH <= 0 Then Exit Sub
        End If
    End If
    
    Call GeomCalc(Box, CurrentTexture, x, y, w, h, sX, sY, sW, sH, Colour)
    Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), Len(Box(0)))
End Sub

Public Sub GeomCalc(ByRef Geom() As Vertex, ByVal textureNum As Long, ByVal x As Single, ByVal y As Single, ByVal w As Integer, ByVal h As Integer, ByVal sX As Single, ByVal sY As Single, ByVal sW As Single, ByVal sH As Single, ByVal Colour As Long)
    sW = (sW + sX) / mTexture(textureNum).w + 0.000003
    sH = (sH + sY) / mTexture(textureNum).h + 0.000003
    sX = sX / mTexture(textureNum).w + 0.000003
    sY = sY / mTexture(textureNum).h + 0.000003
    Geom(0) = MakeVertex(x, y, 0, 1, Colour, 1, sX, sY)
    Geom(1) = MakeVertex(x + w, y, 0, 1, Colour, 0, sW, sY)
    Geom(2) = MakeVertex(x, y + h, 0, 1, Colour, 0, sX, sH)
    Geom(3) = MakeVertex(x + w, y + h, 0, 1, Colour, 0, sW, sH)
End Sub

Private Sub GeomSetBox(ByVal x As Single, ByVal y As Single, ByVal w As Integer, ByVal h As Integer, ByVal Colour As Long)
    Box(0) = MakeVertex(x, y, 0, 1, Colour, 0, 0, 0)
    Box(1) = MakeVertex(x + w, y, 0, 1, Colour, 0, 0, 0)
    Box(2) = MakeVertex(x, y + h, 0, 1, Colour, 0, 0, 0)
    Box(3) = MakeVertex(x + w, y + h, 0, 1, Colour, 0, 0, 0)
End Sub

Private Function MakeVertex(x As Single, y As Single, z As Single, RHW As Single, Colour As Long, Specular As Long, tu As Single, tv As Single) As Vertex
    MakeVertex.x = x
    MakeVertex.y = y
    MakeVertex.z = z
    MakeVertex.RHW = RHW
    MakeVertex.Colour = Colour
    'MakeVertex.Specular = Specular
    MakeVertex.tu = tu
    MakeVertex.tv = tv
End Function

' GDI rendering
Public Sub GDIRenderAnimation()
    Dim i As Long, Animationnum As Long, ShouldRender As Boolean, Width As Long, Height As Long, looptime As Long, FrameCount As Long
    Dim sX As Long, sY As Long, sRECT As RECT
    sRECT.Top = 0
    sRECT.bottom = 192
    sRECT.Left = 0
    sRECT.Right = 192

    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).value

        If Animationnum <= 0 Or Animationnum > Count_Anim Then
            ' don't render lol
        Else
            looptime = frmEditor_Animation.scrlLoopTime(i)

            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            ShouldRender = False

            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= GetTickCount Then

                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If

                AnimEditorTimer(i) = GetTickCount
                ShouldRender = True
            End If

            If ShouldRender Then
                If frmEditor_Animation.scrlFrameCount(i).value > 0 Then
                    ' total width divided by frame count
                    Width = 192
                    Height = 192
                    sY = (Height * ((AnimEditorFrame(i) - 1) \ AnimColumns))
                    sX = (Width * (((AnimEditorFrame(i) - 1) Mod AnimColumns)))
                    ' Start Rendering
                    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
                    Call D3DDevice.BeginScene
                    'EngineRenderRectangle Tex_Anim(Animationnum), 0, 0, sX, sY, width, height, width, height
                    RenderTexture Tex_Anim(Animationnum), 0, 0, sX, sY, Width, Height, Width, Height
                    ' Finish Rendering
                    Call D3DDevice.EndScene
                    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Animation.picSprite(i).hwnd, ByVal 0)
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
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hwnd, ByVal 0)
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
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hwnd, ByVal 0)
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
    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Events.picGraphic.hwnd, ByVal 0)
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
    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Events.picGraphicSel.hwnd, ByVal 0)
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
    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Map.picBackSelect.hwnd, ByVal 0)
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
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hwnd, ByVal 0)
End Sub

Public Sub GDIRenderItemPaperdoll(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim Height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Paperdoll Then Exit Sub
    Height = mTexture(Tex_Paperdoll(sprite)).h
    Width = mTexture(Tex_Paperdoll(sprite)).w
    sRECT.Top = 0
    sRECT.bottom = 72
    sRECT.Left = 0
    sRECT.Right = 144
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle Tex_Item(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Paperdoll(sprite), 0, 0, 0, 0, 144, 72, 144, 72
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hwnd, ByVal 0)
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
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hwnd, ByVal 0)
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
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hwnd, ByVal 0)
End Sub

' Directional blocking
Public Sub DrawDirection(ByVal x As Long, ByVal y As Long)
    Dim i As Long, Top As Long, Left As Long
    ' render grid
    Top = 24
    Left = 0
    'EngineRenderRectangle Tex_Direction, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), left, top, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Direction, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), Left, Top, 32, 32, 32, 32

    ' render dir blobs
    For i = 1 To 4
        Left = (i - 1) * 8

        ' find out whether render blocked or not
        If Not isDirBlocked(Map.TileData.Tile(x, y).DirBlock, CByte(i)) Then
            Top = 8
        Else
            Top = 16
        End If

        'render!
        'EngineRenderRectangle Tex_Direction, ConvertMapX(x * PIC_X) + DirArrowX(i), ConvertMapY(y * PIC_Y) + DirArrowY(i), left, top, 8, 8, 8, 8, 8, 8
        RenderTexture Tex_Direction, ConvertMapX(x * PIC_X) + DirArrowX(i), ConvertMapY(y * PIC_Y) + DirArrowY(i), Left, Top, 8, 8, 8, 8
    Next

End Sub

Public Sub DrawFade()
    RenderTexture Tex_Blank, 0, 0, 0, 0, ScreenWidth, ScreenHeight, 32, 32, DX8Colour(White, fadeAlpha)
End Sub

Public Sub DrawFog()
    Dim fogNum As Long, Colour As Long, x As Long, y As Long, RenderState As Long
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

    For x = 0 To ((Map.MapData.MaxX * 32) / 256) + 1
        For y = 0 To ((Map.MapData.MaxY * 32) / 256) + 1
            RenderTexture Tex_Fog(fogNum), (x * 256) + fogOffsetX, (y * 256) + fogOffsetY, 0, 0, 256, 256, 256, 256, Colour
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
    Dim Color As Long, i As Long, SpriteLeft As Long
    For i = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(i).InUse Then
            If WeatherParticle(i).Type = WEATHER_TYPE_STORM Then
                SpriteLeft = 0
            Else
                SpriteLeft = WeatherParticle(i).Type - 1
            End If
            RenderTexture Tex_Weather, ConvertMapX(WeatherParticle(i).x), ConvertMapY(WeatherParticle(i).y), SpriteLeft * 32, 0, 32, 32, 32, 32, -1
        End If
    Next
End Sub

Public Sub DrawAutoTile(ByVal layernum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal x As Long, ByVal y As Long)
    Dim yOffset As Long, xOffset As Long

    ' calculate the offset
    Select Case Map.TileData.Tile(x, y).Autotile(layernum)

        Case AUTOTILE_WATERFALL
            yOffset = (waterfallFrame - 1) * 32

        Case AUTOTILE_ANIM
            xOffset = autoTileFrame * 64

        Case AUTOTILE_CLIFF
            yOffset = -32
    End Select

    ' Draw the quarter
    RenderTexture Tex_Tileset(Map.TileData.Tile(x, y).Layer(layernum).tileSet), destX, destY, Autotile(x, y).Layer(layernum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layernum).srcY(quarterNum) + yOffset, 16, 16, 16, 16
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
           RenderTexture Tex_Tileset(tileSet), ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), shpSelectedLeft, shpSelectedTop, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 200)
       Else
           RenderTexture Tex_Tileset(tileSet), ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), shpSelectedLeft, shpSelectedTop, shpSelectedWidth, shpSelectedHeight, shpSelectedWidth, shpSelectedHeight
       End If
    End If
End Sub

' Rendering Procedures
Public Sub DrawMapTile(ByVal x As Long, ByVal y As Long)
Dim i As Long, tileSet As Long, sX As Long, sY As Long

    With Map.TileData.Tile(x, y)
        ' draw the map
        For i = MapLayer.Ground To MapLayer.Mask2
            ' skip tile if tileset isn't set
            If Autotile(x, y).Layer(i).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.Layer(i).tileSet), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32
            ElseIf Autotile(x, y).Layer(i).RenderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), 1, x, y
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY(y * PIC_Y), 2, x, y
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY((y * PIC_Y) + 16), 3, x, y
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY((y * PIC_Y) + 16), 4, x, y
            ElseIf Autotile(x, y).Layer(i).RenderState = RENDER_STATE_APPEAR Then
                ' check if it's fading
                If TempTile(x, y).fadeAlpha(i) > 0 Then
                    ' render it
                    tileSet = Map.TileData.Tile(x, y).Layer(i).tileSet
                    sX = Map.TileData.Tile(x, y).Layer(i).x
                    sY = Map.TileData.Tile(x, y).Layer(i).y
                    RenderTexture Tex_Tileset(tileSet), ConvertMapX(x * 32), ConvertMapY(y * 32), sX * 32, sY * 32, 32, 32, 32, 32, DX8Colour(White, TempTile(x, y).fadeAlpha(i))
                End If
            End If
        Next
    End With
End Sub

Public Sub DrawMapFringeTile(ByVal x As Long, ByVal y As Long)
    Dim i As Long

    With Map.TileData.Tile(x, y)
        ' draw the map
        For i = MapLayer.Fringe To MapLayer.Fringe2

            ' skip tile if tileset isn't set
            If Autotile(x, y).Layer(i).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.Layer(i).tileSet), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32
            ElseIf Autotile(x, y).Layer(i).RenderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), 1, x, y
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY(y * PIC_Y), 2, x, y
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY((y * PIC_Y) + 16), 3, x, y
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY((y * PIC_Y) + 16), 4, x, y
            End If
        Next
    End With
End Sub

Public Sub DrawHotbar()
    Dim Xo As Long, Yo As Long, Width As Long, Height As Long, i As Long, t As Long, sS As String
    
    Xo = Windows(GetWindowIndex("winHotbar")).Window.Left
    Yo = Windows(GetWindowIndex("winHotbar")).Window.Top
    
    ' render start + end wood
    RenderTexture Tex_GUI(31), Xo - 1, Yo + 3, 0, 0, 11, 26, 11, 26
    RenderTexture Tex_GUI(31), Xo + 407, Yo + 3, 0, 0, 11, 26, 11, 26
    
    For i = 1 To MAX_HOTBAR
        Xo = Windows(GetWindowIndex("winHotbar")).Window.Left + HotbarLeft + ((i - 1) * HotbarOffsetX)
        Yo = Windows(GetWindowIndex("winHotbar")).Window.Top + HotbarTop
        Width = 36
        Height = 36
        ' don't render last one
        If i <> 10 Then
            ' render wood
            RenderTexture Tex_GUI(32), Xo + 30, Yo + 3, 0, 0, 13, 26, 13, 26
        End If
        ' render box
        RenderTexture Tex_GUI(30), Xo - 2, Yo - 2, 0, 0, Width, Height, Width, Height
        ' render icon
        If Not (DragBox.Origin = origin_Hotbar And DragBox.Slot = i) Then
            Select Case Hotbar(i).sType
                Case 1 ' inventory
                    If Len(Item(Hotbar(i).Slot).Name) > 0 And Item(Hotbar(i).Slot).Pic > 0 Then
                        RenderTexture Tex_Item(Item(Hotbar(i).Slot).Pic), Xo, Yo, 0, 0, 32, 32, 32, 32
                    End If
                Case 2 ' spell
                    If Len(Spell(Hotbar(i).Slot).Name) > 0 And Spell(Hotbar(i).Slot).icon > 0 Then
                        RenderTexture Tex_Spellicon(Spell(Hotbar(i).Slot).icon), Xo, Yo, 0, 0, 32, 32, 32, 32
                        For t = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(t).Spell > 0 Then
                                If PlayerSpells(t).Spell = Hotbar(i).Slot And SpellCD(t) > 0 Then
                                    RenderTexture Tex_Spellicon(Spell(Hotbar(i).Slot).icon), Xo, Yo, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                                End If
                            End If
                        Next
                    End If
            End Select
        End If
        ' draw the numbers
        sS = Str(i)
        If i = 10 Then sS = "0"
        RenderText font(Fonts.rockwellDec_15), sS, Xo + 4, Yo + 19, White
    Next
End Sub

Public Sub RenderAppearTileFade()
Dim x As Long, y As Long, tileSet As Long, sX As Long, sY As Long, layernum As Long

    For x = 0 To Map.MapData.MaxX
        For y = 0 To Map.MapData.MaxY
            For layernum = MapLayer.Ground To MapLayer.Mask
                ' check if it's fading
                If TempTile(x, y).fadeAlpha(layernum) > 0 Then
                    ' render it
                    tileSet = Map.TileData.Tile(x, y).Layer(layernum).tileSet
                    sX = Map.TileData.Tile(x, y).Layer(layernum).x
                    sY = Map.TileData.Tile(x, y).Layer(layernum).y
                    RenderTexture Tex_Tileset(tileSet), ConvertMapX(x * 32), ConvertMapY(y * 32), sX * 32, sY * 32, 32, 32, 32, 32, DX8Colour(White, TempTile(x, y).fadeAlpha(layernum))
                End If
            Next
        Next
    Next
End Sub

Public Sub DrawCharacter()
    Dim Xo As Long, Yo As Long, Width As Long, Height As Long, i As Long, sprite As Long, ItemNum As Long, ItemPic As Long
    
    Xo = Windows(GetWindowIndex("winCharacter")).Window.Left
    Yo = Windows(GetWindowIndex("winCharacter")).Window.Top
    
    ' Render bottom
    RenderTexture Tex_GUI(37), Xo + 4, Yo + 314, 0, 0, 40, 38, 40, 38
    RenderTexture Tex_GUI(37), Xo + 44, Yo + 314, 0, 0, 40, 38, 40, 38
    RenderTexture Tex_GUI(37), Xo + 84, Yo + 314, 0, 0, 40, 38, 40, 38
    RenderTexture Tex_GUI(37), Xo + 124, Yo + 314, 0, 0, 46, 38, 46, 38
    
    ' render top wood
    RenderTexture Tex_GUI(1), Xo + 4, Yo + 23, 100, 100, 166, 291, 166, 291
    
    ' loop through equipment
    For i = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(MyIndex, i)

        ' get the item sprite
        If ItemNum > 0 Then
            ItemPic = Tex_Item(Item(ItemNum).Pic)
        Else
            ' no item equiped - use blank image
            ItemPic = Tex_GUI(37 + i)
        End If
        
        Yo = Windows(GetWindowIndex("winCharacter")).Window.Top + EqTop
        Xo = Windows(GetWindowIndex("winCharacter")).Window.Left + EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))

        RenderTexture ItemPic, Xo, Yo, 0, 0, 32, 32, 32, 32
    Next
End Sub

Public Sub DrawSkills()
    Dim Xo As Long, Yo As Long, Width As Long, Height As Long, i As Long, y As Long, spellnum As Long, spellPic As Long, x As Long, Top As Long, Left As Long
    
    Xo = Windows(GetWindowIndex("winSkills")).Window.Left
    Yo = Windows(GetWindowIndex("winSkills")).Window.Top
    
    Width = Windows(GetWindowIndex("winSkills")).Window.Width
    Height = Windows(GetWindowIndex("winSkills")).Window.Height
    
    ' render green
    RenderTexture Tex_GUI(34), Xo + 4, Yo + 23, 0, 0, Width - 8, Height - 27, 4, 4
    
    Width = 76
    Height = 76
    
    y = Yo + 23
    ' render grid - row
    For i = 1 To 4
        If i = 4 Then Height = 42
        RenderTexture Tex_GUI(35), Xo + 4, y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 80, y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 156, y, 0, 0, 42, Height, 42, Height
        y = y + 76
    Next
    
    ' actually draw the icons
    For i = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(i).Spell
        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            ' not dragging?
            If Not (DragBox.Origin = origin_Spells And DragBox.Slot = i) Then
                spellPic = Spell(spellnum).icon
    
                If spellPic > 0 And spellPic <= Count_Spellicon Then
                    Top = Yo + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    Left = Xo + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
    
                    RenderTexture Tex_Spellicon(spellPic), Left, Top, 0, 0, 32, 32, 32, 32
                End If
            End If
        End If
    Next
End Sub

Public Sub RenderMapName()
Dim zonetype As String, Colour As Long

    If Map.MapData.Moral = 0 Then
        zonetype = "PK Zone"
        Colour = Red
    ElseIf Map.MapData.Moral = 1 Then
        zonetype = "Safe Zone"
        Colour = White
    ElseIf Map.MapData.Moral = 2 Then
        zonetype = "Boss Chamber"
        Colour = Grey
    End If
    
    RenderText font(Fonts.rockwellDec_10), Trim$(Map.MapData.Name) & " - " & zonetype, ScreenWidth - 15 - TextWidth(font(Fonts.rockwellDec_10), Trim$(Map.MapData.Name) & " - " & zonetype), 45, Colour, 255
End Sub

Public Sub DrawShopBackground()
    Dim Xo As Long, Yo As Long, Width As Long, Height As Long, i As Long, y As Long
    
    Xo = Windows(GetWindowIndex("winShop")).Window.Left
    Yo = Windows(GetWindowIndex("winShop")).Window.Top
    Width = Windows(GetWindowIndex("winShop")).Window.Width
    Height = Windows(GetWindowIndex("winShop")).Window.Height
    
    ' render green
    RenderTexture Tex_GUI(34), Xo + 4, Yo + 23, 0, 0, Width - 8, Height - 27, 4, 4
    
    Width = 76
    Height = 76
    
    y = Yo + 23
    ' render grid - row
    For i = 1 To 3
        If i = 3 Then Height = 42
        RenderTexture Tex_GUI(35), Xo + 4, y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 80, y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 156, y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 232, y, 0, 0, 42, Height, 42, Height
        y = y + 76
    Next
    ' render bottom wood
    RenderTexture Tex_GUI(1), Xo + 4, y - 34, 0, 0, 270, 72, 270, 72
End Sub

Public Sub DrawShop()
Dim Xo As Long, Yo As Long, ItemPic As Long, ItemNum As Long, Amount As Long, i As Long, Top As Long, Left As Long, y As Long, x As Long, Colour As Long

    If InShop = 0 Then Exit Sub
    
    Xo = Windows(GetWindowIndex("winShop")).Window.Left
    Yo = Windows(GetWindowIndex("winShop")).Window.Top
    
    If Not shopIsSelling Then
        ' render the shop items
        For i = 1 To MAX_TRADES
            ItemNum = Shop(InShop).TradeItem(i).Item
            
            ' draw early
            Top = Yo + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
            Left = Xo + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
            ' draw selected square
            If shopSelectedSlot = i Then RenderTexture Tex_GUI(61), Left, Top, 0, 0, 32, 32, 32, 32
            
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
                If ItemPic > 0 And ItemPic <= Count_Item Then
                    ' draw item
                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                End If
            End If
        Next
    Else
        ' render the shop items
        For i = 1 To MAX_TRADES
            ItemNum = GetPlayerInvItemNum(MyIndex, i)
            
            ' draw early
            Top = Yo + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
            Left = Xo + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
            ' draw selected square
            If shopSelectedSlot = i Then RenderTexture Tex_GUI(61), Left, Top, 0, 0, 32, 32, 32, 32
            
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
                If ItemPic > 0 And ItemPic <= Count_Item Then

                    ' draw item
                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                    
                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        y = Top + 21
                        x = Left + 1
                        Amount = CStr(GetPlayerInvItemValue(MyIndex, i))
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            Colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            Colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            Colour = BrightGreen
                        End If
                        
                        RenderText font(Fonts.verdana_12), ConvertCurrency(Amount), x, y, Colour
                    End If
                End If
            End If
        Next
    End If
End Sub

Sub DrawTrade()
    Dim Xo As Long, Yo As Long, Width As Long, Height As Long, i As Long, y As Long, x As Long
    
    Xo = Windows(GetWindowIndex("winTrade")).Window.Left
    Yo = Windows(GetWindowIndex("winTrade")).Window.Top
    Width = Windows(GetWindowIndex("winTrade")).Window.Width
    Height = Windows(GetWindowIndex("winTrade")).Window.Height
    
    ' render green
    RenderTexture Tex_GUI(34), Xo + 4, Yo + 23, 0, 0, Width - 8, Height - 27, 4, 4
    
    ' top wood
    RenderTexture Tex_GUI(1), Xo + 4, Yo + 23, 100, 100, Width - 8, 18, Width - 8, 18
    ' left wood
    RenderTexture Tex_GUI(1), Xo + 4, Yo + 41, 350, 0, 5, Height - 45, 5, Height - 45
    ' right wood
    RenderTexture Tex_GUI(1), Xo + Width - 9, Yo + 41, 350, 0, 5, Height - 45, 5, Height - 45
    ' centre wood
    RenderTexture Tex_GUI(1), Xo + 203, Yo + 41, 350, 0, 6, Height - 45, 6, Height - 45
    ' bottom wood
    RenderTexture Tex_GUI(1), Xo + 4, Yo + 307, 100, 100, Width - 8, 75, Width - 8, 75
    
    ' left
    Width = 76
    Height = 76
    y = Yo + 41
    For i = 1 To 4
        If i = 4 Then Height = 38
        RenderTexture Tex_GUI(35), Xo + 4 + 5, y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 80 + 5, y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 156 + 5, y, 0, 0, 42, Height, 42, Height
        y = y + 76
    Next
    
    ' right
    Width = 76
    Height = 76
    y = Yo + 41
    For i = 1 To 4
        If i = 4 Then Height = 38
        RenderTexture Tex_GUI(35), Xo + 4 + 205, y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 80 + 205, y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 156 + 205, y, 0, 0, 42, Height, 42, Height
        y = y + 76
    Next
End Sub

Sub DrawYourTrade()
Dim i As Long, ItemNum As Long, ItemPic As Long, Top As Long, Left As Long, Colour As Long, Amount As String, x As Long, y As Long
Dim Xo As Long, Yo As Long

    Xo = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).Left
    Yo = Windows(GetWindowIndex("winTrade")).Window.Top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).Top
    
    ' your items
    For i = 1 To MAX_INV
        ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic
            If ItemPic > 0 And ItemPic <= Count_Item Then
                Top = Yo + TradeTop + ((TradeOffsetY + 32) * ((i - 1) \ TradeColumns))
                Left = Xo + TradeLeft + ((TradeOffsetX + 32) * (((i - 1) Mod TradeColumns)))

                ' draw icon
                RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                
                ' If item is a stack - draw the amount you have
                If TradeYourOffer(i).value > 1 Then
                    y = Top + 21
                    x = Left + 1
                    Amount = CStr(TradeYourOffer(i).value)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        Colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        Colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        Colour = BrightGreen
                    End If
                    
                    RenderText font(Fonts.verdana_12), ConvertCurrency(Amount), x, y, Colour
                End If
            End If
        End If
    Next
End Sub

Sub DrawTheirTrade()
Dim i As Long, ItemNum As Long, ItemPic As Long, Top As Long, Left As Long, Colour As Long, Amount As String, x As Long, y As Long
Dim Xo As Long, Yo As Long

    Xo = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picTheir")).Left
    Yo = Windows(GetWindowIndex("winTrade")).Window.Top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picTheir")).Top

    ' their items
    For i = 1 To MAX_INV
        ItemNum = TradeTheirOffer(i).num
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic
            If ItemPic > 0 And ItemPic <= Count_Item Then
                Top = Yo + TradeTop + ((TradeOffsetY + 32) * ((i - 1) \ TradeColumns))
                Left = Xo + TradeLeft + ((TradeOffsetX + 32) * (((i - 1) Mod TradeColumns)))

                ' draw icon
                RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                
                ' If item is a stack - draw the amount you have
                If TradeTheirOffer(i).value > 1 Then
                    y = Top + 21
                    x = Left + 1
                    Amount = CStr(TradeTheirOffer(i).value)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        Colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        Colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        Colour = BrightGreen
                    End If
                    
                    RenderText font(Fonts.verdana_12), ConvertCurrency(Amount), x, y, Colour
                End If
            End If
        End If
    Next
End Sub

Public Sub DrawBank()
    Dim x As Long, y As Long, Xo As Long, Yo As Long, Width As Long, Height As Long
    Dim i As Long, ItemNum As Long, ItemPic As Long

    Dim Left As Long, Top As Long
    Dim Colour As Long, skipItem As Boolean, Amount As Long, tmpItem As Long

    Xo = Windows(GetWindowIndex("winBank")).Window.Left
    Yo = Windows(GetWindowIndex("winBank")).Window.Top
    Width = Windows(GetWindowIndex("winBank")).Window.Width
    Height = Windows(GetWindowIndex("winBank")).Window.Height
    
    ' render green
    RenderTexture Tex_GUI(34), Xo + 4, Yo + 23, 0, 0, Width - 8, Height - 27, 4, 4

    Width = 76
    Height = 76

    y = Yo + 23
    ' render grid - row
    For i = 1 To 5
        If i = 5 Then Height = 42
        RenderTexture Tex_GUI(35), Xo + 4, y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 80, y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 156, y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 232, y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 308, y, 0, 0, 79, Height, 79, Height
        y = y + 76
    Next

    ' actually draw the icons
    For i = 1 To MAX_BANK
        ItemNum = Bank.Item(i).num

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ' not dragging?
            If Not (DragBox.Origin = origin_Bank And DragBox.Slot = i) Then
                ItemPic = Item(ItemNum).Pic


                If ItemPic > 0 And ItemPic <= Count_Item Then
                    Top = Yo + BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                    Left = Xo + BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))

                    ' draw icon
                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32

                    ' If item is a stack - draw the amount you have
                    If Bank.Item(i).value > 1 Then
                        y = Top + 21
                        x = Left + 1
                        Amount = Bank.Item(i).value

                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            Colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            Colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            Colour = BrightGreen
                        End If

                        RenderText font(Fonts.rockwell_15), ConvertCurrency(Amount), x, y, Colour
                    End If
                End If
            End If
        End If
    Next

End Sub

Public Sub DrawInventory()
    Dim Xo As Long, Yo As Long, Width As Long, Height As Long, i As Long, y As Long, ItemNum As Long, ItemPic As Long, x As Long, Top As Long, Left As Long, Amount As String
    Dim Colour As Long, skipItem As Boolean, amountModifier  As Long, tmpItem As Long
    
    Xo = Windows(GetWindowIndex("winInventory")).Window.Left
    Yo = Windows(GetWindowIndex("winInventory")).Window.Top
    Width = Windows(GetWindowIndex("winInventory")).Window.Width
    Height = Windows(GetWindowIndex("winInventory")).Window.Height
    
    ' render green
    RenderTexture Tex_GUI(34), Xo + 4, Yo + 23, 0, 0, Width - 8, Height - 27, 4, 4
    
    Width = 76
    Height = 76
    
    y = Yo + 23
    ' render grid - row
    For i = 1 To 4
        If i = 4 Then Height = 38
        RenderTexture Tex_GUI(35), Xo + 4, y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 80, y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 156, y, 0, 0, 42, Height, 42, Height
        y = y + 76
    Next
    ' render bottom wood
    RenderTexture Tex_GUI(1), Xo + 4, Yo + 289, 100, 100, 194, 26, 194, 26
    
    ' actually draw the icons
    For i = 1 To MAX_INV
        ItemNum = GetPlayerInvItemNum(MyIndex, i)
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ' not dragging?
            If Not (DragBox.Origin = origin_Inventory And DragBox.Slot = i) Then
                ItemPic = Item(ItemNum).Pic
                
                ' exit out if we're offering item in a trade.
                amountModifier = 0
                If InTrade > 0 Then
                    For x = 1 To MAX_INV
                        tmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(x).num)
                        If TradeYourOffer(x).num = i Then
                            ' check if currency
                            If Not Item(tmpItem).Type = ITEM_TYPE_CURRENCY Then
                                ' normal item, exit out
                                skipItem = True
                            Else
                                ' if amount = all currency, remove from inventory
                                If TradeYourOffer(x).value = GetPlayerInvItemValue(MyIndex, i) Then
                                    skipItem = True
                                Else
                                    ' not all, change modifier to show change in currency count
                                    amountModifier = TradeYourOffer(x).value
                                End If
                            End If
                        End If
                    Next
                End If
                
                If Not skipItem Then
                    If ItemPic > 0 And ItemPic <= Count_Item Then
                        Top = Yo + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        Left = Xo + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
        
                        ' draw icon
                        RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
        
                        ' If item is a stack - draw the amount you have
                        If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                            y = Top + 21
                            x = Left + 1
                            Amount = GetPlayerInvItemValue(MyIndex, i) - amountModifier
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            If CLng(Amount) < 1000000 Then
                                Colour = White
                            ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                                Colour = Yellow
                            ElseIf CLng(Amount) > 10000000 Then
                                Colour = BrightGreen
                            End If
                            
                            RenderText font(Fonts.verdana_12), ConvertCurrency(Amount), x, y, Colour
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
    Dim theArray() As String, x As Long, y As Long, i As Long, MaxWidth As Long, x2 As Long, y2 As Long, Colour As Long, tmpNum As Long
    
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
                x = ConvertMapX((Player(.target).x * 32) + Player(.target).xOffset) + 16
                y = ConvertMapY((Player(.target).y * 32) + Player(.target).yOffset) - 32
            Case TARGET_TYPE_EVENT
                Colour = .Colour
                x = ConvertMapX(Map.TileData.Events(.target).x * 32) + 16
                y = ConvertMapY(Map.TileData.Events(.target).y * 32) - 16
            Case Else
                Exit Sub
        End Select
        
        ' word wrap
        WordWrap_Array .Msg, ChatBubbleWidth, theArray
        ' find max width
        tmpNum = UBound(theArray)

        For i = 1 To tmpNum
            If TextWidth(font(Fonts.georgiaDec_16), theArray(i)) > MaxWidth Then MaxWidth = TextWidth(font(Fonts.georgiaDec_16), theArray(i))
        Next

        ' calculate the new position
        x2 = x - (MaxWidth \ 2)
        y2 = y - (UBound(theArray) * 12)
        ' render bubble - top left
        RenderTexture Tex_GUI(33), x2 - 9, y2 - 5, 0, 0, 9, 5, 9, 5
        ' top right
        RenderTexture Tex_GUI(33), x2 + MaxWidth, y2 - 5, 119, 0, 9, 5, 9, 5
        ' top
        RenderTexture Tex_GUI(33), x2, y2 - 5, 9, 0, MaxWidth, 5, 5, 5
        ' bottom left
        RenderTexture Tex_GUI(33), x2 - 9, y, 0, 19, 9, 6, 9, 6
        ' bottom right
        RenderTexture Tex_GUI(33), x2 + MaxWidth, y, 119, 19, 9, 6, 9, 6
        ' bottom - left half
        RenderTexture Tex_GUI(33), x2, y, 9, 19, (MaxWidth \ 2) - 5, 6, 9, 6
        ' bottom - right half
        RenderTexture Tex_GUI(33), x2 + (MaxWidth \ 2) + 6, y, 9, 19, (MaxWidth \ 2) - 5, 6, 9, 6
        ' left
        RenderTexture Tex_GUI(33), x2 - 9, y2, 0, 6, 9, (UBound(theArray) * 12), 9, 1
        ' right
        RenderTexture Tex_GUI(33), x2 + MaxWidth, y2, 119, 6, 9, (UBound(theArray) * 12), 9, 1
        ' center
        RenderTexture Tex_GUI(33), x2, y2, 9, 5, MaxWidth, (UBound(theArray) * 12), 1, 1
        ' little pointy bit
        RenderTexture Tex_GUI(33), x - 5, y, 58, 19, 11, 11, 11, 11
        ' render each line centralised
        tmpNum = UBound(theArray)

        For i = 1 To tmpNum
            RenderText font(Fonts.georgia_16), theArray(i), x - (TextWidth(font(Fonts.georgiaDec_16), theArray(i)) / 2), y2, Colour
            y2 = y2 + 12
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
    Dim x As Long
    Dim y As Long
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
                Case DIR_UP_LEFT
                    If (Player(Index).yOffset > 16) Then Anim = Player(Index).Step
                    If (Player(Index).xOffset > 16) Then Anim = Player(Index).Step
                Case DIR_UP_RIGHT
                    If (Player(Index).yOffset > 16) Then Anim = Player(Index).Step
                    If (Player(Index).xOffset < -16) Then Anim = Player(Index).Step
                Case DIR_DOWN_LEFT
                    If (Player(Index).yOffset < -16) Then Anim = Player(Index).Step
                    If (Player(Index).xOffset > 16) Then Anim = Player(Index).Step
                Case DIR_DOWN_RIGHT
                    If (Player(Index).yOffset < -16) Then Anim = Player(Index).Step
                    If (Player(Index).xOffset < -16) Then Anim = Player(Index).Step
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

        Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
            spritetop = 2

        Case DIR_DOWN
            spritetop = 0

        Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = spritetop * (mTexture(Tex_Char(sprite)).h / 4)
        .Height = (mTexture(Tex_Char(sprite)).h / 4)
        .Left = Anim * (mTexture(Tex_Char(sprite)).w / 4)
        .Width = (mTexture(Tex_Char(sprite)).w / 4)
    End With

    ' Calculate the X
    x = GetPlayerX(Index) * PIC_X + Player(Index).xOffset - ((mTexture(Tex_Char(sprite)).w / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (mTexture(Tex_Char(sprite)).h) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - ((mTexture(Tex_Char(sprite)).h / 4) - 32) - 4
    Else
        ' Proceed as normal
        y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - 4
    End If

    RenderTexture Tex_Char(sprite), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.Width, rec.Height, rec.Width, rec.Height
End Sub

Public Sub DrawNpc(ByVal MapNpcNum As Long)
    Dim Anim As Byte
    Dim x As Long
    Dim y As Long
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
                Case DIR_UP_LEFT
                    If (MapNpc(MapNpcNum).yOffset > 16) Then Anim = MapNpc(MapNpcNum).Step
                    If (MapNpc(MapNpcNum).xOffset > 16) Then Anim = MapNpc(MapNpcNum).Step
                Case DIR_UP_RIGHT
                    If (MapNpc(MapNpcNum).yOffset > 16) Then Anim = MapNpc(MapNpcNum).Step
                    If (MapNpc(MapNpcNum).xOffset < -16) Then Anim = MapNpc(MapNpcNum).Step
                Case DIR_DOWN_LEFT
                    If (MapNpc(MapNpcNum).yOffset < -16) Then Anim = MapNpc(MapNpcNum).Step
                    If (MapNpc(MapNpcNum).xOffset > 16) Then Anim = MapNpc(MapNpcNum).Step
                Case DIR_DOWN_RIGHT
                    If (MapNpc(MapNpcNum).yOffset < -16) Then Anim = MapNpc(MapNpcNum).Step
                    If (MapNpc(MapNpcNum).xOffset < -16) Then Anim = MapNpc(MapNpcNum).Step
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

        Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
            spritetop = 2

        Case DIR_DOWN
            spritetop = 0

        Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = (mTexture(Tex_Char(sprite)).h / 4) * spritetop
        .Height = mTexture(Tex_Char(sprite)).h / 4
        .Left = Anim * (mTexture(Tex_Char(sprite)).w / 4)
        .Width = (mTexture(Tex_Char(sprite)).w / 4)
    End With

    ' Calculate the X
    x = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).xOffset - ((mTexture(Tex_Char(sprite)).w / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (mTexture(Tex_Char(sprite)).h / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).yOffset - ((mTexture(Tex_Char(sprite)).h / 4) - 32) - 4
    Else
        ' Proceed as normal
        y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).yOffset - 4
    End If

    RenderTexture Tex_Char(sprite), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.Top, rec.Width, rec.Height, rec.Width, rec.Height
End Sub

Sub DrawEvent(eventNum As Long, pageNum As Long)
Dim texNum As Long, x As Long, y As Long

    ' render it
    With Map.TileData.Events(eventNum).EventPage(pageNum)
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
                    x = ConvertMapX(Map.TileData.Events(eventNum).x * 32)
                    y = ConvertMapY(Map.TileData.Events(eventNum).y * 32)
                    RenderTexture texNum, x, y, .GraphicX * 32, .GraphicY * 32, 32, 32, 32, 32
                End If
            End If
        End If
    End With
End Sub

Sub DrawLowerEvents()
Dim i As Long, x As Long

    If Map.TileData.EventCount = 0 Then Exit Sub
    For i = 1 To Map.TileData.EventCount
        ' find the active page
        If Map.TileData.Events(i).pageCount > 0 Then
            x = ActiveEventPage(i)
            If x > 0 Then
                ' make sure it's lower
                If Map.TileData.Events(i).EventPage(x).Priority <> 2 Then
                    ' render event
                    DrawEvent i, x
                End If
            End If
        End If
    Next
End Sub

Sub DrawUpperEvents()
Dim i As Long, x As Long

    If Map.TileData.EventCount = 0 Then Exit Sub
    For i = 1 To Map.TileData.EventCount
        ' find the active page
        If Map.TileData.Events(i).pageCount > 0 Then
            x = ActiveEventPage(i)
            If x > 0 Then
                ' make sure it's lower
                If Map.TileData.Events(i).EventPage(x).Priority = 2 Then
                    ' render event
                    DrawEvent i, x
                End If
            End If
        End If
    Next
End Sub

Public Sub DrawShadow(ByVal sprite As Long, ByVal x As Long, ByVal y As Long)
    If hasSpriteShadow(sprite) Then RenderTexture Tex_Shadow, ConvertMapX(x), ConvertMapY(y), 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawTarget(ByVal x As Long, ByVal y As Long)
    Dim Width As Long, Height As Long
    ' calculations
    Width = mTexture(Tex_Target).w / 2
    Height = mTexture(Tex_Target).h
    x = x - ((Width - 32) / 2)
    y = y - (Height / 2) + 16
    x = ConvertMapX(x)
    y = ConvertMapY(y)
    'EngineRenderRectangle Tex_Target, x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_Target, x, y, 0, 0, Width, Height, Width, Height
End Sub

Public Sub DrawTargetHover()
    Dim i As Long, x As Long, y As Long, Width As Long, Height As Long

    If diaIndex > 0 Then Exit Sub
    Width = mTexture(Tex_Target).w / 2
    Height = mTexture(Tex_Target).h

    If Width <= 0 Then Width = 1
    If Height <= 0 Then Height = 1

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(MyIndex) = GetPlayerMap(i) Then
            x = (Player(i).x * 32) + Player(i).xOffset + 32
            y = (Player(i).y * 32) + Player(i).yOffset + 32

            If x >= GlobalX_Map And x <= GlobalX_Map + 32 Then
                If y >= GlobalY_Map And y <= GlobalY_Map + 32 Then
                    x = ConvertMapX(x)
                    y = ConvertMapY(y)
                    RenderTexture Tex_Target, x - 16 - (Width / 2), y - 16 - (Height / 2), Width, 0, Width, Height, Width, Height
                End If
            End If
        End If

    Next

    For i = 1 To MAX_MAP_NPCS

        If MapNpc(i).num > 0 Then
            x = (MapNpc(i).x * 32) + MapNpc(i).xOffset + 32
            y = (MapNpc(i).y * 32) + MapNpc(i).yOffset + 32

            If x >= GlobalX_Map And x <= GlobalX_Map + 32 Then
                If y >= GlobalY_Map And y <= GlobalY_Map + 32 Then
                    x = ConvertMapX(x)
                    y = ConvertMapY(y)
                    RenderTexture Tex_Target, x - 16 - (Width / 2), y - 16 - (Height / 2), Width, 0, Width, Height, Width, Height
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
    Dim x As Long, y As Long
    Dim Width As Long, Height As Long
    x = MapResource(Resource_num).x
    y = MapResource(Resource_num).y

    If x < 0 Or x > Map.MapData.MaxX Then Exit Sub
    If y < 0 Or y > Map.MapData.MaxY Then Exit Sub
    ' Get the Resource type
    Resource_master = Map.TileData.Tile(x, y).Data1

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
    x = (MapResource(Resource_num).x * PIC_X) - (mTexture(Tex_Resource(Resource_sprite)).w / 2) + 16
    y = (MapResource(Resource_num).y * PIC_Y) - mTexture(Tex_Resource(Resource_sprite)).h + 32
    Width = rec.Right - rec.Left
    Height = rec.bottom - rec.Top
    'EngineRenderRectangle Tex_Resource(Resource_sprite), ConvertMapX(x), ConvertMapY(y), 0, 0, width, height, width, height, width, height
    RenderTexture Tex_Resource(Resource_sprite), ConvertMapX(x), ConvertMapY(y), 0, 0, Width, Height, Width, Height
End Sub

Public Sub DrawItem(ByVal ItemNum As Long)
    Dim PicNum As Integer, dontRender As Boolean, i As Long, tmpIndex As Long
    PicNum = Item(MapItem(ItemNum).num).Pic

    If PicNum < 1 Or PicNum > Count_Item Then Exit Sub

    ' if it's not us then don't render
    If MapItem(ItemNum).playerName <> vbNullString Then
        If Trim$(MapItem(ItemNum).playerName) <> Trim$(GetPlayerName(MyIndex)) Then

            dontRender = True
        End If

        ' make sure it's not a party drop
        If Party.Leader > 0 Then

            For i = 1 To MAX_PARTY_MEMBERS
                tmpIndex = Party.Member(i)

                If tmpIndex > 0 Then
                    If Trim$(GetPlayerName(tmpIndex)) = Trim$(MapItem(ItemNum).playerName) Then
                        If MapItem(ItemNum).bound = 0 Then

                            dontRender = False
                        End If
                    End If
                End If

            Next

        End If
    End If

    'If Not dontRender Then EngineRenderRectangle Tex_Item(PicNum), ConvertMapX(MapItem(itemnum).x * PIC_X), ConvertMapY(MapItem(itemnum).y * PIC_Y), 0, 0, 32, 32, 32, 32, 32, 32
    If Not dontRender Then
        RenderTexture Tex_Item(PicNum), ConvertMapX(MapItem(ItemNum).x * PIC_X), ConvertMapY(MapItem(ItemNum).y * PIC_Y), 0, 0, 32, 32, 32, 32
    End If

End Sub

Public Sub DrawBars()
Dim Left As Long, Top As Long, Width As Long, Height As Long
Dim tmpX As Long, tmpY As Long, barWidth As Long, i As Long, NpcNum As Long
Dim partyIndex As Long

    ' dynamic bar calculations
    Width = mTexture(Tex_Bars).w
    Height = mTexture(Tex_Bars).h / 4
    
    ' render npc health bars
    For i = 1 To MAX_MAP_NPCS
        NpcNum = MapNpc(i).num
        ' exists?
        If NpcNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) < Npc(NpcNum).HP Then
                ' lock to npc
                tmpX = MapNpc(i).x * PIC_X + MapNpc(i).xOffset + 16 - (Width / 2)
                tmpY = MapNpc(i).y * PIC_Y + MapNpc(i).yOffset + 35
                
                ' calculate the width to fill
                If Width > 0 Then BarWidth_NpcHP_Max(i) = ((MapNpc(i).Vital(Vitals.HP) / Width) / (Npc(NpcNum).HP / Width)) * Width
                
                ' draw bar background
                Top = Height * 1 ' HP bar background
                Left = 0
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, Width, Height, Width, Height
                
                ' draw the bar proper
                Top = 0 ' HP bar
                Left = 0
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, BarWidth_NpcHP(i), Height, BarWidth_NpcHP(i), Height
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
    Dim x As Long, y As Long, lockindex As Long

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
                    x = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + Player(lockindex).xOffset
                    y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + Player(lockindex).yOffset
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
                    x = (MapNpc(lockindex).x * PIC_X) + 16 - (Width / 2) + MapNpc(lockindex).xOffset
                    y = (MapNpc(lockindex).y * PIC_Y) + 16 - (Height / 2) + MapNpc(lockindex).yOffset
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
        x = (AnimInstance(Index).x * 32) + 16 - (Width / 2)
        y = (AnimInstance(Index).y * 32) + 16 - (Height / 2)
    End If

    x = ConvertMapX(x)
    y = ConvertMapY(y)
    'EngineRenderRectangle Tex_Anim(sprite), x, y, sRECT.left, sRECT.top, sRECT.width, sRECT.height, sRECT.width, sRECT.height, sRECT.width, sRECT.height
    RenderTexture Tex_Anim(sprite), x, y, sRECT.Left, sRECT.Top, sRECT.Width, sRECT.Height, sRECT.Width, sRECT.Height
End Sub

Public Sub DrawGDI()

    If frmEditor_Animation.visible Then
        GDIRenderAnimation
    ElseIf frmEditor_Item.visible Then
        GDIRenderItem frmEditor_Item.picItem, frmEditor_Item.scrlPic.value
        GDIRenderItemPaperdoll frmEditor_Item.picPaperdoll, frmEditor_Item.scrlPaperdoll.value
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
    Dim x As Long, y As Long, i As Long, bgColour As Long
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
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    RenderTexture Tex_Fader, ConvertMapX(x * 32), ConvertMapY(y * 32), 0, 0, 32, 32, 32, 32
                End If
            Next
        Next
    End If
    
    ' Render appear tile fades
    'RenderAppearTileFade

    ' render lower tiles
    If Count_Tileset > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    Call DrawMapTile(x, y)
                End If
            Next
        Next
    End If

    ' render the items
    If Count_Item > 0 Then
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).num > 0 Then
                Call DrawItem(i)
            End If
        Next
    End If

    ' draw animations
    If Count_Anim > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(0) Then
                DrawAnimation i, 0
            End If
        Next
    End If
    
    ' draw events
    DrawLowerEvents

    ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
    For y = TileView.Top To TileView.bottom + 5
        ' Resources
        If Count_Resource > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).y = y Then
                            Call DrawResource(i)
                        End If
                    Next
                End If
            End If
        End If
        
        If Count_Char > 0 Then
            ' shadows - Players
            For i = 1 To Player_HighIndex
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Player(i).y = y Then
                        Call DrawShadow(Player(i).sprite, (Player(i).x * 32) + Player(i).xOffset, (Player(i).y * 32) + Player(i).yOffset)
                    End If
                End If
            Next
    
            ' shadows - npcs
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then
                    If MapNpc(i).y = y Then
                        Call DrawShadow(Npc(MapNpc(i).num).sprite, (MapNpc(i).x * 32) + MapNpc(i).xOffset, (MapNpc(i).y * 32) + MapNpc(i).yOffset)
                    End If
                End If
            Next
    
            ' Players
            For i = 1 To Player_HighIndex
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Player(i).y = y Then
                        Call DrawPlayer(i)
                    End If
                End If
            Next
    
            ' Npcs
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).y = y Then
                    Call DrawNpc(i)
                End If
            Next
        End If
    Next y

    ' render out upper tiles
    If Count_Tileset > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    Call DrawMapFringeTile(x, y)
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
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(1) Then
                DrawAnimation i, 1
            End If
        Next
    End If

    ' render target
    If myTarget > 0 Then
        If myTargetType = TARGET_TYPE_PLAYER Then
            DrawTarget (Player(myTarget).x * 32) + Player(myTarget).xOffset, (Player(myTarget).y * 32) + Player(myTarget).yOffset
        ElseIf myTargetType = TARGET_TYPE_NPC Then
            DrawTarget (MapNpc(myTarget).x * 32) + MapNpc(myTarget).xOffset, (MapNpc(myTarget).y * 32) + MapNpc(myTarget).yOffset
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
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call DrawPlayerName(i)
            End If
        Next
    End If

    ' draw npc names
    If Not screenshotMode Then
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(i).num > 0 Then
                Call DrawNpcName(i)
            End If
        Next
    End If

    ' draw action msg
    For i = 1 To MAX_BYTE
        DrawActionMsg i
    Next

    If InMapEditor Then
        If frmEditor_Map.optBlock.value = True Then
            For x = TileView.Left To TileView.Right
                For y = TileView.Top To TileView.bottom
                    If IsValidMapPoint(x, y) Then
                        Call DrawDirection(x, y)
                    End If
                Next
            Next
        End If
    End If

    ' draw the messages
    For i = 1 To MAX_BYTE
        If chatBubble(i).active Then
            DrawChatBubble i
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
