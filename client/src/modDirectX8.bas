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
Public TextureAnim() As Long
Public TextureChar() As Long
Public TextureFace() As Long
Public TextureItem() As Long
Public TexturePaperdoll() As Long
Public TextureResource() As Long
Public TextureSpellIcon() As Long
Public TextureTileset() As Long
Public TextureFog() As Long
Public TextureGUI() As Long
Public TextureDesign() As Long
Public TextureGradient() As Long
Public TextureSurface() As Long
Public TextureBars As Long
Public TextureBlood As Long
Public TextureDirection As Long
Public TextureMisc As Long
Public TextureTarget As Long
Public TextureShadow As Long
Public TextureFader As Long
Public TextureBlank As Long
Public TextureWeather As Long
Public TextureWhite As Long
Public TextureCaptcha() As Long

' Texture count
Public CountAnim As Long
Public CountChar As Long
Public CountFace As Long
Public CountGUI As Long
Public CountDesign As Long
Public CountGradient As Long
Public CountItem As Long
Public CountPaperdoll As Long
Public CountResource As Long
Public CountSpellicon As Long
Public CountTileset As Long
Public CountFog As Long
Public CountSurface As Long
Public CountCaptcha As Long

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
Dim DispMode As D3DDISPLAYMODE, width As Long, height As Long

    mhWnd = hwnd

    Set DX8 = New DirectX8
    Set D3D = DX8.Direct3DCreate
    Set D3DX = New D3DX8
    
    ' set size
    GetResolutionSize curResolution, width, height
    ScreenWidth = width
    ScreenHeight = height
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
    TextureCaptcha = LoadTextureFiles(CountCaptcha, App.path & Path_Captcha)
    TextureTileset = LoadTextureFiles(CountTileset, App.path & Path_Tileset)
    TextureAnim = LoadTextureFiles(CountAnim, App.path & Path_Anim)
    TextureChar = LoadTextureFiles(CountChar, App.path & Path_Char)
    TextureFace = LoadTextureFiles(CountFace, App.path & Path_Face)
    TextureItem = LoadTextureFiles(CountItem, App.path & Path_Item)
    TexturePaperdoll = LoadTextureFiles(CountPaperdoll, App.path & Path_Paperdoll)
    TextureResource = LoadTextureFiles(CountResource, App.path & Path_Resource)
    TextureSpellIcon = LoadTextureFiles(CountSpellicon, App.path & Path_Spellicon)
    TextureGUI = LoadTextureFiles(CountGUI, App.path & Path_GUI)
    TextureDesign = LoadTextureFiles(CountDesign, App.path & Path_Design)
    TextureGradient = LoadTextureFiles(CountGradient, App.path & Path_Gradient)
    TextureSurface = LoadTextureFiles(CountSurface, App.path & Path_Surface)
    TextureFog = LoadTextureFiles(CountFog, App.path & Path_Fog)
    ' Singles
    TextureBars = LoadTextureFile(App.path & Path_Graphics & "bars.png")
    TextureBlood = LoadTextureFile(App.path & Path_Graphics & "blood.png")
    TextureDirection = LoadTextureFile(App.path & Path_Graphics & "direction.png")
    TextureMisc = LoadTextureFile(App.path & Path_Graphics & "misc.png")
    TextureTarget = LoadTextureFile(App.path & Path_Graphics & "target.png")
    TextureShadow = LoadTextureFile(App.path & Path_Graphics & "shadow.png")
    TextureFader = LoadTextureFile(App.path & Path_Graphics & "fader.png")
    TextureBlank = LoadTextureFile(App.path & Path_Graphics & "blank.png")
    TextureWeather = LoadTextureFile(App.path & Path_Graphics & "weather.png")
    TextureWhite = LoadTextureFile(App.path & Path_Graphics & "white.png")
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
        If mClip.top <> 0 Then
            If mClip.left > x Then
                sX = sX + (mClip.left - x) / (w / sW)
                sW = sW - (mClip.left - x) / (w / sW)
                w = w - (mClip.left - x)
                x = mClip.left
            End If
            
            If mClip.top > y Then
                sY = sY + (mClip.top - y) / (h / sH)
                sH = sH - (mClip.top - y) / (h / sH)
                h = h - (mClip.top - y)
                y = mClip.top
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
    Dim i As Long, Animationnum As Long, ShouldRender As Boolean, width As Long, height As Long, looptime As Long, FrameCount As Long
    Dim sX As Long, sY As Long, sRECT As RECT
    sRECT.top = 0
    sRECT.bottom = 192
    sRECT.left = 0
    sRECT.Right = 192

    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).value

        If Animationnum <= 0 Or Animationnum > CountAnim Then
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
                    width = 192
                    height = 192
                    sY = (height * ((AnimEditorFrame(i) - 1) \ AnimColumns))
                    sX = (width * (((AnimEditorFrame(i) - 1) Mod AnimColumns)))
                    ' Start Rendering
                    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
                    Call D3DDevice.BeginScene
                    'EngineRenderRectangle TextureAnim(Animationnum), 0, 0, sX, sY, width, height, width, height
                    RenderTexture TextureAnim(Animationnum), 0, 0, sX, sY, width, height, width, height
                    ' Finish Rendering
                    Call D3DDevice.EndScene
                    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Animation.picSprite(i).hwnd, ByVal 0)
                End If
            End If
        End If

    Next

End Sub

Public Sub GDIRenderChar(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim height As Long, width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > CountChar Then Exit Sub
    height = 32
    width = 32
    sRECT.top = 0
    sRECT.bottom = sRECT.top + height
    sRECT.left = 0
    sRECT.Right = sRECT.left + width
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    RenderTexture TextureChar(sprite), 0, 0, 0, 0, width, height, width, height
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hwnd, ByVal 0)
End Sub

Public Sub GDIRenderFace(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim height As Long, width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > CountFace Then Exit Sub
    height = mTexture(TextureFace(sprite)).h
    width = mTexture(TextureFace(sprite)).w

    If height = 0 Or width = 0 Then
        height = 1
        width = 1
    End If

    sRECT.top = 0
    sRECT.bottom = sRECT.top + height
    sRECT.left = 0
    sRECT.Right = sRECT.left + width
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle TextureFace(sprite), 0, 0, 0, 0, width, height, width, height, width, height
    RenderTexture TextureFace(sprite), 0, 0, 0, 0, width, height, width, height
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hwnd, ByVal 0)
End Sub

Public Sub GDIRenderTileset()
    Dim height As Long, width As Long, tileSet As Byte, sRECT As RECT
    ' find tileset number
    tileSet = frmEditor_Map.scrlTileSet.value

    ' exit out if doesn't exist
    If tileSet <= 0 Or tileSet > CountTileset Then Exit Sub
    height = mTexture(TextureTileset(tileSet)).h
    width = mTexture(TextureTileset(tileSet)).w

    If height = 0 Or width = 0 Then
        height = 1
        width = 1
    End If

    frmEditor_Map.picBackSelect.width = width
    frmEditor_Map.picBackSelect.height = height
    sRECT.top = 0
    sRECT.bottom = height
    sRECT.left = 0
    sRECT.Right = width

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

    'EngineRenderRectangle TextureTileset(Tileset), 0, 0, 0, 0, width, height, width, height, width, height
    If TextureTileset(tileSet) <= 0 Then Exit Sub
    RenderTexture TextureTileset(tileSet), 0, 0, 0, 0, width, height, width, height
    ' draw selection boxes
    RenderDesign DesignTypes.designTilesetGrid, shpSelectedLeft, shpSelectedTop, shpSelectedWidth, shpSelectedHeight
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Map.picBackSelect.hwnd, ByVal 0)
End Sub

Public Sub GDIRenderItem(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim height As Long, width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > CountItem Then Exit Sub
    height = mTexture(TextureItem(sprite)).h
    width = mTexture(TextureItem(sprite)).w
    sRECT.top = 0
    sRECT.bottom = 32
    sRECT.left = 0
    sRECT.Right = 32
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle TextureItem(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture TextureItem(sprite), 0, 0, 0, 0, 32, 32, 32, 32
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hwnd, ByVal 0)
End Sub

Public Sub GDIRenderItemPaperdoll(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim height As Long, width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > CountPaperdoll Then Exit Sub
    height = mTexture(TexturePaperdoll(sprite)).h
    width = mTexture(TexturePaperdoll(sprite)).w
    sRECT.top = 0
    sRECT.bottom = 72
    sRECT.left = 0
    sRECT.Right = 144
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle TextureItem(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture TexturePaperdoll(sprite), 0, 0, 0, 0, 144, 72, 144, 72
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hwnd, ByVal 0)
End Sub

Public Sub GDIRenderResource(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim height As Long, width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > CountResource Then Exit Sub
    height = mTexture(TextureResource(sprite)).h
    width = mTexture(TextureResource(sprite)).w
    sRECT.top = 0
    sRECT.bottom = 152
    sRECT.left = 0
    sRECT.Right = 152
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    RenderTexture TextureResource(sprite), 0, 0, 0, 0, width, height, width, height
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hwnd, ByVal 0)
End Sub


Public Sub GDIRenderSpell(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim height As Long, width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > CountSpellicon Then Exit Sub
    height = mTexture(TextureSpellIcon(sprite)).h
    width = mTexture(TextureSpellIcon(sprite)).w

    If height = 0 Or width = 0 Then
        height = 1
        width = 1
    End If

    sRECT.top = 0
    sRECT.bottom = height
    sRECT.left = 0
    sRECT.Right = width
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle TextureSpellIcon(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture TextureSpellIcon(sprite), 0, 0, 0, 0, 32, 32, 32, 32
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hwnd, ByVal 0)
End Sub

' Directional blocking
Public Sub DrawDirection(ByVal x As Long, ByVal y As Long)
    Dim i As Long, top As Long, left As Long
    ' render grid
    top = 24
    left = 0
    'EngineRenderRectangle TextureDirection, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), left, top, 32, 32, 32, 32, 32, 32
    RenderTexture TextureDirection, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), left, top, 32, 32, 32, 32

    ' render dir blobs
    For i = 1 To 4
        left = (i - 1) * 8

        ' find out whether render blocked or not
        If Not isDirBlocked(Map.TileData.Tile(x, y).DirBlock, CByte(i)) Then
            top = 8
        Else
            top = 16
        End If

        'render!
        'EngineRenderRectangle TextureDirection, ConvertMapX(x * PIC_X) + DirArrowX(i), ConvertMapY(y * PIC_Y) + DirArrowY(i), left, top, 8, 8, 8, 8, 8, 8
        RenderTexture TextureDirection, ConvertMapX(x * PIC_X) + DirArrowX(i), ConvertMapY(y * PIC_Y) + DirArrowY(i), left, top, 8, 8, 8, 8
    Next

End Sub

Public Sub DrawFade()
    RenderTexture TextureBlank, 0, 0, 0, 0, ScreenWidth, ScreenHeight, 32, 32, DX8Colour(White, fadeAlpha)
End Sub

Public Sub DrawFog()
    Dim fogNum As Long, Colour As Long, x As Long, y As Long, RenderState As Long
    fogNum = CurrentFog

    If fogNum <= 0 Or fogNum > CountFog Then Exit Sub
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
            RenderTexture TextureFog(fogNum), ConvertMapX((x * 256) + fogOffsetX), ConvertMapY((y * 256) + fogOffsetY), 0, 0, 256, 256, 256, 256, Colour
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
    RenderTexture TextureWhite, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, Color
End Sub

Public Sub DrawWeather()
    Dim Color As Long, i As Long, SpriteLeft As Long
    For i = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(i).InUse Then
            If WeatherParticle(i).type = WEATHER_TYPE_STORM Then
                SpriteLeft = 0
            Else
                SpriteLeft = WeatherParticle(i).type - 1
            End If
            RenderTexture TextureWeather, ConvertMapX(WeatherParticle(i).x), ConvertMapY(WeatherParticle(i).y), SpriteLeft * 32, 0, 32, 32, 32, 32, -1
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
    RenderTexture TextureTileset(Map.TileData.Tile(x, y).Layer(layernum).tileSet), destX, destY, Autotile(x, y).Layer(layernum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layernum).srcY(quarterNum) + yOffset, 16, 16, 16, 16
End Sub

Sub DrawTileSelection()
    Dim tileSet As Byte
    ' find tileset number
    tileSet = frmEditor_Map.scrlTileSet.value
    
    ' exit out if doesn't exist
    If tileSet <= 0 Or tileSet > CountTileset Then Exit Sub

    If frmEditor_Map.scrlAutotile.value > 0 Then
        RenderTexture TextureTileset(tileSet), ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), shpSelectedLeft, shpSelectedTop, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 200)
    Else
        RenderTexture TextureTileset(tileSet), ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), shpSelectedLeft, shpSelectedTop, shpSelectedWidth, shpSelectedHeight, shpSelectedWidth, shpSelectedHeight
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
                RenderTexture TextureTileset(.Layer(i).tileSet), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32
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
                    RenderTexture TextureTileset(tileSet), ConvertMapX(x * 32), ConvertMapY(y * 32), sX * 32, sY * 32, 32, 32, 32, 32, DX8Colour(White, TempTile(x, y).fadeAlpha(i))
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
                RenderTexture TextureTileset(.Layer(i).tileSet), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32
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
    Dim Xo As Long, Yo As Long, width As Long, height As Long, i As Long, t As Long, sS As String
    
    Xo = Windows(GetWindowIndex("winHotbar")).Window.left
    Yo = Windows(GetWindowIndex("winHotbar")).Window.top
    
    ' render start + end wood
    RenderTexture TextureGUI(40), Xo - 1, Yo + 3, 0, 0, 11, 26, 11, 26
    RenderTexture TextureGUI(40), Xo + 407, Yo + 3, 0, 0, 11, 26, 11, 26
    
    For i = 1 To MAX_HOTBAR
        Xo = Windows(GetWindowIndex("winHotbar")).Window.left + HotbarLeft + ((i - 1) * HotbarOffsetX)
        Yo = Windows(GetWindowIndex("winHotbar")).Window.top + HotbarTop
        width = 36
        height = 36
        ' don't render last one
        If i <> 10 Then
            ' render wood
            RenderTexture TextureGUI(41), Xo + 30, Yo + 3, 0, 0, 13, 26, 13, 26
        End If
        ' render box
        RenderTexture TextureGUI(35), Xo - 2, Yo - 2, 0, 0, width, height, width, height
        ' render icon
        If Not (DragBox.origin = originHotbar And DragBox.Slot = i) Then
            Select Case Hotbar(i).sType
                Case 1 ' inventory
                    If Len(Item(Hotbar(i).Slot).name) > 0 And Item(Hotbar(i).Slot).Pic > 0 Then
                        RenderTexture TextureItem(Item(Hotbar(i).Slot).Pic), Xo, Yo, 0, 0, 32, 32, 32, 32
                    End If
                Case 2 ' spell
                    If Len(Spell(Hotbar(i).Slot).name) > 0 And Spell(Hotbar(i).Slot).icon > 0 Then
                        RenderTexture TextureSpellIcon(Spell(Hotbar(i).Slot).icon), Xo, Yo, 0, 0, 32, 32, 32, 32
                        For t = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(t).Spell > 0 Then
                                If PlayerSpells(t).Spell = Hotbar(i).Slot And SpellCD(t) > 0 Then
                                    RenderTexture TextureSpellIcon(Spell(Hotbar(i).Slot).icon), Xo, Yo, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
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
                    RenderTexture TextureTileset(tileSet), ConvertMapX(x * 32), ConvertMapY(y * 32), sX * 32, sY * 32, 32, 32, 32, 32, DX8Colour(White, TempTile(x, y).fadeAlpha(layernum))
                End If
            Next
        Next
    Next
End Sub

Public Sub DrawCharacter()
    Dim Xo As Long, Yo As Long, width As Long, height As Long, i As Long, sprite As Long, ItemNum As Long, ItemPic As Long
    Dim xEquipBar As Long, yEquipBar As Long, yOffSetEquip As Long
    
    Xo = Windows(GetWindowIndex("winCharacter")).Window.left
    Yo = Windows(GetWindowIndex("winCharacter")).Window.top
    
    xEquipBar = Xo
    yEquipBar = Yo
    yOffSetEquip = EqTop
    
    For i = 1 To Equipment.Equipment_Count - 1
        RenderTexture TextureGUI(37), xEquipBar + 170, yEquipBar + yOffSetEquip, 0, 0, 40, 38, 40, 38
        yOffSetEquip = yOffSetEquip + 38
    Next
    
    ' render top wood
    RenderTexture TextureGUI(1), Xo + 4, Yo + 23, 100, 100, 166, 291, 166, 291
    RenderTexture TextureGUI(1), Xo + 170, Yo + 23, 100, 100, 40, 63, 40, 63
    
    ' loop through equipment
    For i = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(MyIndex, i)

        ' get the item sprite
        If ItemNum > 0 Then
            ItemPic = TextureItem(Item(ItemNum).Pic)
        Else
            ' no item equiped - use blank image
            ItemPic = TextureGUI(45 + i)
        End If
        
        Yo = Windows(GetWindowIndex("winCharacter")).Window.top + EqTop + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
        Xo = Windows(GetWindowIndex("winCharacter")).Window.left + EqLeft

        RenderTexture ItemPic, Xo, Yo, 0, 0, 32, 32, 32, 32
    Next
End Sub

Public Sub DrawSkills()
    Dim Xo As Long, Yo As Long, width As Long, height As Long, i As Long, y As Long, spellnum As Long, spellPic As Long, x As Long, top As Long, left As Long
    
    Xo = Windows(GetWindowIndex("winSkills")).Window.left
    Yo = Windows(GetWindowIndex("winSkills")).Window.top
    
    width = Windows(GetWindowIndex("winSkills")).Window.width
    height = Windows(GetWindowIndex("winSkills")).Window.height
    
    ' render green
    RenderTexture TextureGUI(34), Xo + 4, Yo + 23, 0, 0, width - 8, height - 27, 4, 4
    
    width = 76
    height = 76
    
    y = Yo + 23
    ' render grid - row
    For i = 1 To 4
        If i = 4 Then height = 42
        RenderTexture TextureGUI(38), Xo + 4, y, 0, 0, width, height, width, height
        RenderTexture TextureGUI(38), Xo + 80, y, 0, 0, width, height, width, height
        RenderTexture TextureGUI(38), Xo + 156, y, 0, 0, 42, height, 42, height
        y = y + 76
    Next
    
    ' actually draw the icons
    For i = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(i).Spell
        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            ' not dragging?
            If Not (DragBox.origin = originSpells And DragBox.Slot = i) Then
                spellPic = Spell(spellnum).icon
    
                If spellPic > 0 And spellPic <= CountSpellicon Then
                    top = Yo + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    left = Xo + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
    
                    RenderTexture TextureSpellIcon(spellPic), left, top, 0, 0, 32, 32, 32, 32
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
    
    RenderText font(Fonts.rockwellDec_10), Trim$(Map.MapData.name) & " - " & zonetype, ScreenWidth - 15 - TextWidth(font(Fonts.rockwellDec_10), Trim$(Map.MapData.name) & " - " & zonetype), 45, Colour, 255
End Sub

Public Sub DrawInviteBackground()
    Dim Xo As Long, Yo As Long, width As Long, height As Long, i As Long, y As Long
    
    Xo = Windows(GetWindowIndex("winOffer")).Window.left + 475
    Yo = Windows(GetWindowIndex("winOffer")).Window.top
    
    width = 45
    height = 45
    
    y = Yo
    
    For i = 1 To 3
        If inOffer(i) > 0 Then
            RenderDesign DesignTypes.designWindowDescription, Xo, y, width, height
            RenderText font(Fonts.georgia_16), "i", Xo + 21, y + 15, Grey
            y = y + 37
        End If
    Next
End Sub

Public Sub DrawShopBackground()
    Dim Xo As Long, Yo As Long, width As Long, height As Long, i As Long, y As Long
    
    Xo = Windows(GetWindowIndex("winShop")).Window.left
    Yo = Windows(GetWindowIndex("winShop")).Window.top
    width = Windows(GetWindowIndex("winShop")).Window.width
    height = Windows(GetWindowIndex("winShop")).Window.height
    
    ' render green
    RenderTexture TextureGUI(34), Xo + 4, Yo + 23, 0, 0, width - 8, height - 27, 4, 4
    
    width = 76
    height = 76
    
    y = Yo + 23
    ' render grid - row
    For i = 1 To 3
        If i = 3 Then height = 42
        RenderTexture TextureGUI(38), Xo + 4, y, 0, 0, width, height, width, height
        RenderTexture TextureGUI(38), Xo + 80, y, 0, 0, width, height, width, height
        RenderTexture TextureGUI(38), Xo + 156, y, 0, 0, width, height, width, height
        RenderTexture TextureGUI(38), Xo + 232, y, 0, 0, 42, height, 42, height
        y = y + 76
    Next
    ' render bottom wood
    RenderTexture TextureGUI(1), Xo + 4, y - 34, 0, 0, 270, 72, 270, 72
End Sub

Public Sub DrawShop()
Dim Xo As Long, Yo As Long, ItemPic As Long, ItemNum As Long, Amount As Long, i As Long, top As Long, left As Long, y As Long, x As Long, Colour As Long

    If InShop = 0 Then Exit Sub
    
    ' AJUSTAR
    
    Xo = Windows(GetWindowIndex("winShop")).Window.left
    Yo = Windows(GetWindowIndex("winShop")).Window.top
    
    If Not shopIsSelling Then
        ' render the shop items
        For i = 1 To MAX_TRADES
            ItemNum = Shop(InShop).TradeItem(i).Item
            
            ' draw early
            top = Yo + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
            left = Xo + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
            ' draw selected square
            If shopSelectedSlot = i Then RenderTexture TextureGUI(35), left, top, 0, 0, 32, 32, 32, 32
            
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
                If ItemPic > 0 And ItemPic <= CountItem Then
                    ' draw item
                    RenderTexture TextureItem(ItemPic), left, top, 0, 0, 32, 32, 32, 32
                End If
            End If
        Next
    Else
        ' render the shop items
        For i = 1 To MAX_TRADES
            ItemNum = GetPlayerInvItemNum(MyIndex, i)
            
            ' draw early
            top = Yo + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
            left = Xo + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
            ' draw selected square
            If shopSelectedSlot = i Then RenderTexture TextureGUI(35), left, top, 0, 0, 32, 32, 32, 32
            
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
                If ItemPic > 0 And ItemPic <= CountItem Then

                    ' draw item
                    RenderTexture TextureItem(ItemPic), left, top, 0, 0, 32, 32, 32, 32
                    
                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        y = top + 21
                        x = left + 1
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
    Dim Xo As Long, Yo As Long, width As Long, height As Long, i As Long, y As Long, x As Long
    
    Xo = Windows(GetWindowIndex("winTrade")).Window.left
    Yo = Windows(GetWindowIndex("winTrade")).Window.top
    width = Windows(GetWindowIndex("winTrade")).Window.width
    height = Windows(GetWindowIndex("winTrade")).Window.height
    
    ' render green
    RenderTexture TextureGUI(34), Xo + 4, Yo + 23, 0, 0, width - 8, height - 27, 4, 4
    
    ' top wood
    RenderTexture TextureGUI(1), Xo + 4, Yo + 23, 100, 100, width - 8, 18, width - 8, 18
    ' left wood
    RenderTexture TextureGUI(1), Xo + 4, Yo + 41, 350, 0, 5, height - 45, 5, height - 45
    ' right wood
    RenderTexture TextureGUI(1), Xo + width - 9, Yo + 41, 350, 0, 5, height - 45, 5, height - 45
    ' centre wood
    RenderTexture TextureGUI(1), Xo + 203, Yo + 41, 350, 0, 6, height - 45, 6, height - 45
    ' bottom wood
    RenderTexture TextureGUI(1), Xo + 4, Yo + 307, 100, 100, width - 8, 75, width - 8, 75
    
    ' left
    width = 76
    height = 76
    y = Yo + 41
    For i = 1 To 4
        If i = 4 Then height = 38
        RenderTexture TextureGUI(38), Xo + 4 + 5, y, 0, 0, width, height, width, height
        RenderTexture TextureGUI(38), Xo + 80 + 5, y, 0, 0, width, height, width, height
        RenderTexture TextureGUI(38), Xo + 156 + 5, y, 0, 0, 42, height, 42, height
        y = y + 76
    Next
    
    ' right
    width = 76
    height = 76
    y = Yo + 41
    For i = 1 To 4
        If i = 4 Then height = 38
        RenderTexture TextureGUI(38), Xo + 4 + 205, y, 0, 0, width, height, width, height
        RenderTexture TextureGUI(38), Xo + 80 + 205, y, 0, 0, width, height, width, height
        RenderTexture TextureGUI(38), Xo + 156 + 205, y, 0, 0, 42, height, 42, height
        y = y + 76
    Next
End Sub

Sub DrawYourTrade()
Dim i As Long, ItemNum As Long, ItemPic As Long, top As Long, left As Long, Colour As Long, Amount As String, x As Long, y As Long
Dim Xo As Long, Yo As Long

    Xo = Windows(GetWindowIndex("winTrade")).Window.left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).left
    Yo = Windows(GetWindowIndex("winTrade")).Window.top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).top
    
    ' your items
    For i = 1 To MAX_INV
        ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic
            If ItemPic > 0 And ItemPic <= CountItem Then
                top = Yo + TradeTop + ((TradeOffsetY + 32) * ((i - 1) \ TradeColumns))
                left = Xo + TradeLeft + ((TradeOffsetX + 32) * (((i - 1) Mod TradeColumns)))

                ' draw icon
                RenderTexture TextureItem(ItemPic), left, top, 0, 0, 32, 32, 32, 32
                
                ' If item is a stack - draw the amount you have
                If TradeYourOffer(i).value > 1 Then
                    y = top + 21
                    x = left + 1
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
Dim i As Long, ItemNum As Long, ItemPic As Long, top As Long, left As Long, Colour As Long, Amount As String, x As Long, y As Long
Dim Xo As Long, Yo As Long

    Xo = Windows(GetWindowIndex("winTrade")).Window.left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picTheir")).left
    Yo = Windows(GetWindowIndex("winTrade")).Window.top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picTheir")).top

    ' their items
    For i = 1 To MAX_INV
        ItemNum = TradeTheirOffer(i).num
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic
            If ItemPic > 0 And ItemPic <= CountItem Then
                top = Yo + TradeTop + ((TradeOffsetY + 32) * ((i - 1) \ TradeColumns))
                left = Xo + TradeLeft + ((TradeOffsetX + 32) * (((i - 1) Mod TradeColumns)))

                ' draw icon
                RenderTexture TextureItem(ItemPic), left, top, 0, 0, 32, 32, 32, 32
                
                ' If item is a stack - draw the amount you have
                If TradeTheirOffer(i).value > 1 Then
                    y = top + 21
                    x = left + 1
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
    Dim x As Long, y As Long, Xo As Long, Yo As Long, width As Long, height As Long
    Dim i As Long, ItemNum As Long, ItemPic As Long

    Dim left As Long, top As Long
    Dim Colour As Long, skipItem As Boolean, Amount As Long, tmpItem As Long

    Xo = Windows(GetWindowIndex("winBank")).Window.left
    Yo = Windows(GetWindowIndex("winBank")).Window.top
    width = Windows(GetWindowIndex("winBank")).Window.width
    height = Windows(GetWindowIndex("winBank")).Window.height
    
    ' render green
    RenderTexture TextureGUI(34), Xo + 4, Yo + 23, 0, 0, width - 8, height - 27, 4, 4

    width = 76
    height = 76

    y = Yo + 23
    ' render grid - row
    For i = 1 To 5
        If i = 5 Then height = 42
        RenderTexture TextureGUI(38), Xo + 4, y, 0, 0, width, height, width, height
        RenderTexture TextureGUI(38), Xo + 80, y, 0, 0, width, height, width, height
        RenderTexture TextureGUI(38), Xo + 156, y, 0, 0, width, height, width, height
        RenderTexture TextureGUI(38), Xo + 232, y, 0, 0, width, height, width, height
        RenderTexture TextureGUI(38), Xo + 308, y, 0, 0, 79, height, 79, height
        y = y + 76
    Next

    ' actually draw the icons
    For i = 1 To MAX_BANK
        ItemNum = Bank.Item(i).num

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ' not dragging?
            If Not (DragBox.origin = originBank And DragBox.Slot = i) Then
                ItemPic = Item(ItemNum).Pic


                If ItemPic > 0 And ItemPic <= CountItem Then
                    top = Yo + BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                    left = Xo + BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))

                    ' draw icon
                    RenderTexture TextureItem(ItemPic), left, top, 0, 0, 32, 32, 32, 32

                    ' If item is a stack - draw the amount you have
                    If Bank.Item(i).value > 1 Then
                        y = top + 21
                        x = left + 1
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
    Dim Xo As Long, Yo As Long, width As Long, height As Long, i As Long, y As Long, ItemNum As Long, ItemPic As Long, x As Long, top As Long, left As Long, Amount As String
    Dim Colour As Long, skipItem As Boolean, amountModifier  As Long, tmpItem As Long
    
    Xo = Windows(GetWindowIndex("winInventory")).Window.left
    Yo = Windows(GetWindowIndex("winInventory")).Window.top
    width = Windows(GetWindowIndex("winInventory")).Window.width
    height = Windows(GetWindowIndex("winInventory")).Window.height
    
    ' render green
    RenderTexture TextureGUI(34), Xo + 4, Yo + 23, 0, 0, width - 8, height - 27, 4, 4
    
    width = 76
    height = 76
    
    y = Yo + 23
    ' render grid - row
    For i = 1 To 4
        If i = 4 Then height = 38
        RenderTexture TextureGUI(38), Xo + 4, y, 0, 0, width, height, width, height
        RenderTexture TextureGUI(38), Xo + 80, y, 0, 0, width, height, width, height
        RenderTexture TextureGUI(38), Xo + 156, y, 0, 0, 42, height, 42, height
        y = y + 76
    Next
    ' render bottom wood
    RenderTexture TextureGUI(1), Xo + 4, Yo + 289, 100, 100, 194, 26, 194, 26
    
    ' actually draw the icons
    For i = 1 To MAX_INV
        ItemNum = GetPlayerInvItemNum(MyIndex, i)
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ' not dragging?
            If Not (DragBox.origin = originInventory And DragBox.Slot = i) Then
                ItemPic = Item(ItemNum).Pic
                
                ' exit out if we're offering item in a trade.
                amountModifier = 0
                If InTrade > 0 Then
                    For x = 1 To MAX_INV
                        tmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(x).num)
                        If TradeYourOffer(x).num = i Then
                            ' check if currency
                            If Not Item(tmpItem).type = ITEM_TYPE_CURRENCY Then
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
                    If ItemPic > 0 And ItemPic <= CountItem Then
                        top = Yo + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        left = Xo + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
        
                        ' draw icon
                        RenderTexture TextureItem(ItemPic), left, top, 0, 0, 32, 32, 32, 32
        
                        ' If item is a stack - draw the amount you have
                        If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                            y = top + 21
                            x = left + 1
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

Public Sub DrawWinQuest()
    Dim Xo As Long, Yo As Long, width As Long, height As Long, i As Long, x As Long, y As Long, ItemNum As Long, ItemPic As Long, top As Long, left As Long, Amount As String
    Dim Colour As Long, skipItem As Boolean, amountModifier  As Long, tmpItem As Long
    
    Xo = Windows(GetWindowIndex("winPlayerQuests")).Window.left
    Yo = Windows(GetWindowIndex("winPlayerQuests")).Window.top
    width = Windows(GetWindowIndex("winPlayerQuests")).Window.width
    height = Windows(GetWindowIndex("winPlayerQuests")).Window.height
    
    ' render green
    RenderTexture TextureDesign(5), Xo + 4, Yo + 23, 3, 3, width - 8, height - 27, 4, 4
    
    width = 42
    height = 42
    
    x = Xo + 132
    
    ' render div vertical
    RenderTexture TextureGUI(1), Xo + 132, Yo + 23, 100, 100, 4, 385, 4, 385
    
    ' render bottom wood
    RenderTexture TextureGUI(1), Xo + 325, Yo + 366, 100, 100, 121, 42, 121, 42
    ' render grid - row
    For i = 1 To 5
        RenderTexture TextureGUI(38), x, Yo + 366, 0, 0, width, height, width, height
        
        If btnMissionActive <> 0 Then
            ItemNum = Mission(Player(MyIndex).Mission(btnMissionActive).id).RewardItem(i).ItemNum
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
                If ItemPic > 0 And ItemPic <= CountItem Then

                    ' draw icon
                    RenderTexture TextureItem(ItemPic), x + 4, Yo + 370, 0, 0, 32, 32, 32, 32

                    ' If item is a stack - draw the amount you have
                    If Mission(Player(MyIndex).Mission(btnMissionActive).id).RewardItem(i).ItemAmount > 1 Then
                        y = Yo + 370 + 21
                        x = x + 4 + 1
                        Amount = Mission(Player(MyIndex).Mission(btnMissionActive).id).RewardItem(i).ItemAmount

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
        End If
        
        x = x + 38
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
        RenderTexture TextureGUI(39), x2 - 9, y2 - 5, 0, 0, 9, 5, 9, 5
        ' top right
        RenderTexture TextureGUI(39), x2 + MaxWidth, y2 - 5, 119, 0, 9, 5, 9, 5
        ' top
        RenderTexture TextureGUI(39), x2, y2 - 5, 9, 0, MaxWidth, 5, 5, 5
        ' bottom left
        RenderTexture TextureGUI(39), x2 - 9, y, 0, 19, 9, 6, 9, 6
        ' bottom right
        RenderTexture TextureGUI(39), x2 + MaxWidth, y, 119, 19, 9, 6, 9, 6
        ' bottom - left half
        RenderTexture TextureGUI(39), x2, y, 9, 19, (MaxWidth \ 2) - 5, 6, 9, 6
        ' bottom - right half
        RenderTexture TextureGUI(39), x2 + (MaxWidth \ 2) + 6, y, 9, 19, (MaxWidth \ 2) - 5, 6, 9, 6
        ' left
        RenderTexture TextureGUI(39), x2 - 9, y2, 0, 6, 9, (UBound(theArray) * 12), 9, 1
        ' right
        RenderTexture TextureGUI(39), x2 + MaxWidth, y2, 119, 6, 9, (UBound(theArray) * 12), 9, 1
        ' center
        RenderTexture TextureGUI(39), x2, y2, 9, 5, MaxWidth, (UBound(theArray) * 12), 1, 1
        ' little pointy bit
        RenderTexture TextureGUI(39), x - 5, y, 58, 19, 11, 11, 11, 11
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
    Dim Anim As Byte, i As Long
    Dim x As Long
    Dim y As Long
    Dim sprite As Long, SpriteTop As Long
    Dim rec As GeomRec
    Dim attackspeed As Long
    
    ' pre-load sprite for calculations
    sprite = GetPlayerSprite(Index)

    'SetTexture TextureChar(Sprite)
    If sprite < 1 Or sprite > CountChar Then Exit Sub

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
            SpriteTop = 3

        Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
            SpriteTop = 2

        Case DIR_DOWN
            SpriteTop = 0

        Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
            SpriteTop = 1
    End Select

    With rec
        .top = SpriteTop * (mTexture(TextureChar(sprite)).h / 4)
        .height = (mTexture(TextureChar(sprite)).h / 4)
        .left = Anim * (mTexture(TextureChar(sprite)).w / 4)
        .width = (mTexture(TextureChar(sprite)).w / 4)
    End With

    ' Calculate the X
    x = GetPlayerX(Index) * PIC_X + Player(Index).xOffset - ((mTexture(TextureChar(sprite)).w / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (mTexture(TextureChar(sprite)).h) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - ((mTexture(TextureChar(sprite)).h / 4) - 32) - 4
    Else
        ' Proceed as normal
        y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - 4
    End If

    RenderTexture TextureChar(sprite), ConvertMapX(x), ConvertMapY(y), rec.left, rec.top, rec.width, rec.height, rec.width, rec.height
    
    ' check for paperdolling
    For i = 1 To UBound(PaperdollOrder)
        If GetPlayerEquipment(Index, PaperdollOrder(i)) > 0 Then
            If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll > 0 Then
                Call DrawPaperdoll(Index, Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, x, y, rec)
            End If
        End If
    Next
End Sub

Public Sub DrawPaperdoll(ByVal Index As Long, ByVal sprite As Long, ByVal x2 As Long, y2 As Long, rec As GeomRec)
    Dim x As Long, y As Long
    Dim width As Long, height As Long

    If sprite < 1 Or sprite > CountPaperdoll Then Exit Sub

    width = (rec.width - rec.left)
    height = (rec.height - rec.top)
    
    RenderTexture TexturePaperdoll(sprite), ConvertMapX(x2), ConvertMapY(y2), rec.left, rec.top, rec.width, rec.height, rec.width, rec.height, D3DColorRGBA(255, 255, 255, 255)
End Sub

Public Sub DrawNpc(ByVal MapNpcNum As Long)
    Dim Anim As Byte
    Dim x As Long
    Dim y As Long
    Dim sprite As Long, SpriteTop As Long
    Dim rec As GeomRec
    Dim attackspeed As Long

    If MapNpc(MapNpcNum).num = 0 Then Exit Sub ' no npc set
    ' pre-load texture for calculations
    sprite = Npc(MapNpc(MapNpcNum).num).sprite

    'SetTexture TextureChar(Sprite)
    If sprite < 1 Or sprite > CountChar Then Exit Sub
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
            SpriteTop = 3

        Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
            SpriteTop = 2

        Case DIR_DOWN
            SpriteTop = 0

        Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
            SpriteTop = 1
    End Select

    With rec
        .top = (mTexture(TextureChar(sprite)).h / 4) * SpriteTop
        .height = mTexture(TextureChar(sprite)).h / 4
        .left = Anim * (mTexture(TextureChar(sprite)).w / 4)
        .width = (mTexture(TextureChar(sprite)).w / 4)
    End With

    ' Calculate the X
    x = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).xOffset - ((mTexture(TextureChar(sprite)).w / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (mTexture(TextureChar(sprite)).h / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).yOffset - ((mTexture(TextureChar(sprite)).h / 4) - 32) - 4
    Else
        ' Proceed as normal
        y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).yOffset - 4
    End If

    RenderTexture TextureChar(sprite), ConvertMapX(x), ConvertMapY(y), rec.left, rec.top, rec.width, rec.height, rec.width, rec.height
End Sub

Sub DrawEvent(EventNum As Long, pageNum As Long)
Dim texNum As Long, x As Long, y As Long

    ' render it
    With Map.TileData.Events(EventNum).EventPage(pageNum)
        If .GraphicType > 0 Then
            If .Graphic > 0 Then
                Select Case .GraphicType
                    Case 1 ' character
                        If .Graphic < CountChar Then
                            texNum = TextureChar(.Graphic)
                        End If
                    Case 2 ' tileset
                        If .Graphic < CountTileset Then
                            texNum = TextureTileset(.Graphic)
                        End If
                End Select
                If texNum > 0 Then
                    x = ConvertMapX(Map.TileData.Events(EventNum).x * 32)
                    y = ConvertMapY(Map.TileData.Events(EventNum).y * 32)
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
    If hasSpriteShadow(sprite) Then RenderTexture TextureShadow, ConvertMapX(x), ConvertMapY(y), 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawTarget(ByVal x As Long, ByVal y As Long)
    Dim width As Long, height As Long
    ' calculations
    width = mTexture(TextureTarget).w / 2
    height = mTexture(TextureTarget).h
    x = x - ((width - 32) / 2)
    y = y - (height / 2) + 16
    x = ConvertMapX(x)
    y = ConvertMapY(y)
    'EngineRenderRectangle TextureTarget, x, y, 0, 0, width, height, width, height, width, height
    RenderTexture TextureTarget, x, y, 0, 0, width, height, width, height
End Sub

Public Sub DrawTargetHover()
    Dim i As Long, x As Long, y As Long, width As Long, height As Long

    If diaIndex > 0 Then Exit Sub
    width = mTexture(TextureTarget).w / 2
    height = mTexture(TextureTarget).h

    If width <= 0 Then width = 1
    If height <= 0 Then height = 1

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(MyIndex) = GetPlayerMap(i) Then
            x = (Player(i).x * 32) + Player(i).xOffset + 32
            y = (Player(i).y * 32) + Player(i).yOffset + 32

            If x >= GlobalX_Map And x <= GlobalX_Map + 32 Then
                If y >= GlobalY_Map And y <= GlobalY_Map + 32 Then
                    x = ConvertMapX(x)
                    y = ConvertMapY(y)
                    RenderTexture TextureTarget, x - 16 - (width / 2), y - 16 - (height / 2), width, 0, width, height, width, height
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
                    RenderTexture TextureTarget, x - 16 - (width / 2), y - 16 - (height / 2), width, 0, width, height, width, height
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
    Dim width As Long, height As Long
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
    'SetTexture TextureResource(Resource_sprite)
    ' src rect
    With rec
        .top = 0
        .bottom = mTexture(TextureResource(Resource_sprite)).h
        .left = 0
        .Right = mTexture(TextureResource(Resource_sprite)).w
    End With

    ' Set base x + y, then the offset due to size
    x = (MapResource(Resource_num).x * PIC_X) - (mTexture(TextureResource(Resource_sprite)).w / 2) + 16
    y = (MapResource(Resource_num).y * PIC_Y) - mTexture(TextureResource(Resource_sprite)).h + 32
    width = rec.Right - rec.left
    height = rec.bottom - rec.top
    'EngineRenderRectangle TextureResource(Resource_sprite), ConvertMapX(x), ConvertMapY(y), 0, 0, width, height, width, height, width, height
    RenderTexture TextureResource(Resource_sprite), ConvertMapX(x), ConvertMapY(y), 0, 0, width, height, width, height
End Sub

Public Sub DrawItem(ByVal ItemNum As Long)
    Dim PicNum As Integer, dontRender As Boolean, i As Long, tmpIndex As Long
    PicNum = Item(MapItem(ItemNum).num).Pic

    If PicNum < 1 Or PicNum > CountItem Then Exit Sub

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

    'If Not dontRender Then EngineRenderRectangle TextureItem(PicNum), ConvertMapX(MapItem(itemnum).x * PIC_X), ConvertMapY(MapItem(itemnum).y * PIC_Y), 0, 0, 32, 32, 32, 32, 32, 32
    If Not dontRender Then
        RenderTexture TextureItem(PicNum), ConvertMapX(MapItem(ItemNum).x * PIC_X), ConvertMapY(MapItem(ItemNum).y * PIC_Y), 0, 0, 32, 32, 32, 32
    End If

End Sub

Public Sub DrawBars()
Dim left As Long, top As Long, width As Long, height As Long
Dim tmpX As Long, tmpY As Long, barWidth As Long, i As Long, NpcNum As Long
Dim partyIndex As Long

    ' dynamic bar calculations
    width = mTexture(TextureBars).w
    height = mTexture(TextureBars).h / 4
    
    ' render npc health bars
    For i = 1 To MAX_MAP_NPCS
        NpcNum = MapNpc(i).num
        ' exists?
        If NpcNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) < Npc(NpcNum).HP Then
                ' lock to npc
                tmpX = MapNpc(i).x * PIC_X + MapNpc(i).xOffset + 16 - (width / 2)
                tmpY = MapNpc(i).y * PIC_Y + MapNpc(i).yOffset + 35
                
                ' calculate the width to fill
                If width > 0 Then BarWidth_NpcHP_Max(i) = ((MapNpc(i).Vital(Vitals.HP) / width) / (Npc(NpcNum).HP / width)) * width
                
                ' draw bar background
                top = height * 1 ' HP bar background
                left = 0
                RenderTexture TextureBars, ConvertMapX(tmpX), ConvertMapY(tmpY), left, top, width, height, width, height
                
                ' draw the bar proper
                top = 0 ' HP bar
                left = 0
                RenderTexture TextureBars, ConvertMapX(tmpX), ConvertMapY(tmpY), left, top, BarWidth_NpcHP(i), height, BarWidth_NpcHP(i), height
            End If
        End If
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer).Spell).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 16 - (width / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).yOffset + 35 + height + 1
            
            ' calculate the width to fill
            If width > 0 Then barWidth = (GetTickCount - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer).Spell).CastTime * 1000)) * width
            
            ' draw bar background
            top = height * 3 ' cooldown bar background
            left = 0
            RenderTexture TextureBars, ConvertMapX(tmpX), ConvertMapY(tmpY), left, top, width, height, width, height
             
            ' draw the bar proper
            top = height * 2 ' cooldown bar
            left = 0
            RenderTexture TextureBars, ConvertMapX(tmpX), ConvertMapY(tmpY), left, top, barWidth, height, barWidth, height
        End If
    End If
    
    ' draw own health bar
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 And GetPlayerVital(MyIndex, Vitals.HP) < GetPlayerMaxVital(MyIndex, Vitals.HP) Then
        ' lock to Player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 16 - (width / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).yOffset + 35
       
        ' calculate the width to fill
        If width > 0 Then BarWidth_PlayerHP_Max(MyIndex) = ((GetPlayerVital(MyIndex, Vitals.HP) / width) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / width)) * width
       
        ' draw bar background
        top = height * 1 ' HP bar background
        left = 0
        RenderTexture TextureBars, ConvertMapX(tmpX), ConvertMapY(tmpY), left, top, width, height, width, height
       
        ' draw the bar proper
        top = 0 ' HP bar
        left = 0
        RenderTexture TextureBars, ConvertMapX(tmpX), ConvertMapY(tmpY), left, top, BarWidth_PlayerHP(MyIndex), height, BarWidth_PlayerHP(MyIndex), height
    End If
End Sub

Public Sub DrawMenuBG()
    ' row 1
    RenderTexture TextureSurface(1), ScreenWidth - 512, ScreenHeight - 512, 0, 0, 512, 512, 512, 512
    RenderTexture TextureSurface(2), ScreenWidth - 1024, ScreenHeight - 512, 0, 0, 512, 512, 512, 512
    RenderTexture TextureSurface(3), ScreenWidth - 1536, ScreenHeight - 512, 0, 0, 512, 512, 512, 512
    RenderTexture TextureSurface(4), ScreenWidth - 2048, ScreenHeight - 512, 0, 0, 512, 512, 512, 512
    ' row 2
    RenderTexture TextureSurface(5), ScreenWidth - 512, ScreenHeight - 1024, 0, 0, 512, 512, 512, 512
    RenderTexture TextureSurface(6), ScreenWidth - 1024, ScreenHeight - 1024, 0, 0, 512, 512, 512, 512
    RenderTexture TextureSurface(7), ScreenWidth - 1536, ScreenHeight - 1024, 0, 0, 512, 512, 512, 512
    RenderTexture TextureSurface(8), ScreenWidth - 2048, ScreenHeight - 1024, 0, 0, 512, 512, 512, 512
    ' row 3
    RenderTexture TextureSurface(9), ScreenWidth - 512, ScreenHeight - 1088, 0, 0, 512, 64, 512, 64
    RenderTexture TextureSurface(10), ScreenWidth - 1024, ScreenHeight - 1088, 0, 0, 512, 64, 512, 64
    RenderTexture TextureSurface(11), ScreenWidth - 1536, ScreenHeight - 1088, 0, 0, 512, 64, 512, 64
    RenderTexture TextureSurface(12), ScreenWidth - 2048, ScreenHeight - 1088, 0, 0, 512, 64, 512, 64
End Sub

Public Sub DrawAnimation(ByVal Index As Long, ByVal Layer As Long)
    Dim sprite As Integer, sRECT As GeomRec, width As Long, height As Long, FrameCount As Long
    Dim x As Long, y As Long, lockindex As Long

    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If

    sprite = Animation(AnimInstance(Index).Animation).sprite(Layer)

    If sprite < 1 Or sprite > CountAnim Then Exit Sub
    ' pre-load texture for calculations
    'SetTexture TextureAnim(Sprite)
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    ' total width divided by frame count
    width = 192 'mTexture(TextureAnim(Sprite)).width / frameCount
    height = 192 'mTexture(TextureAnim(Sprite)).height

    With sRECT
        .top = (height * ((AnimInstance(Index).FrameIndex(Layer) - 1) \ AnimColumns))
        .height = height
        .left = (width * (((AnimInstance(Index).FrameIndex(Layer) - 1) Mod AnimColumns)))
        .width = width
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
                    x = (GetPlayerX(lockindex) * PIC_X) + 16 - (width / 2) + Player(lockindex).xOffset
                    y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (height / 2) + Player(lockindex).yOffset
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
                    x = (MapNpc(lockindex).x * PIC_X) + 16 - (width / 2) + MapNpc(lockindex).xOffset
                    y = (MapNpc(lockindex).y * PIC_Y) + 16 - (height / 2) + MapNpc(lockindex).yOffset
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
        x = (AnimInstance(Index).x * 32) + 16 - (width / 2)
        y = (AnimInstance(Index).y * 32) + 16 - (height / 2)
    End If

    x = ConvertMapX(x)
    y = ConvertMapY(y)
    'EngineRenderRectangle TextureAnim(sprite), x, y, sRECT.left, sRECT.top, sRECT.width, sRECT.height, sRECT.width, sRECT.height, sRECT.width, sRECT.height
    RenderTexture TextureAnim(sprite), x, y, sRECT.left, sRECT.top, sRECT.width, sRECT.height, sRECT.width, sRECT.height
End Sub

Public Sub DrawGDI()

    If frmEditor_Animation.visible Then
        GDIRenderAnimation
    ElseIf frmEditor_Item.visible Then
        GDIRenderItem frmEditor_Item.picItem, frmEditor_Item.scrlPic.value
        GDIRenderItemPaperdoll frmEditor_Item.picPaperdoll, frmEditor_Item.scrlPaperdoll.value
    ElseIf frmEditor_Map.visible Then
        GDIRenderTileset
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
        For x = TileView.left To TileView.Right
            For y = TileView.top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    RenderTexture TextureFader, ConvertMapX(x * 32), ConvertMapY(y * 32), 0, 0, 32, 32, 32, 32
                End If
            Next
        Next
    End If
    
    ' Render appear tile fades
    'RenderAppearTileFade

    ' render lower tiles
    If CountTileset > 0 Then
        For x = TileView.left To TileView.Right
            For y = TileView.top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    Call DrawMapTile(x, y)
                End If
            Next
        Next
    End If

    ' render the items
    If CountItem > 0 Then
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).num > 0 Then
                Call DrawItem(i)
            End If
        Next
    End If

    ' draw animations
    If CountAnim > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(0) Then
                DrawAnimation i, 0
            End If
        Next
    End If
    
    ' draw events
    DrawLowerEvents

    ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
    For y = TileView.top To TileView.bottom + 5
        ' Resources
        If CountResource > 0 Then
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
        
        If CountChar > 0 Then
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
    If CountTileset > 0 Then
        For x = TileView.left To TileView.Right
            For y = TileView.top To TileView.bottom
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
    If CountAnim > 0 Then
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
            For x = TileView.left To TileView.Right
                For y = TileView.top To TileView.bottom
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
    
    If DrawThunder > 0 Then RenderTexture TextureWhite, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, 160): DrawThunder = DrawThunder - 1
    
    ' draw shadow
    If Not screenshotMode Then
        RenderTexture TextureGUI(43), 0, 0, 0, 0, ScreenWidth, 64, 1, 64
        RenderTexture TextureGUI(42), 0, ScreenHeight - 64, 0, 0, ScreenWidth, 64, 1, 64
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
    
    RenderText font(Fonts.georgiaDec_16), Trim$("StartX: " & StartX & " CameraStarX: " & Camera.left), 260, 6, Yellow
    RenderText font(Fonts.georgiaDec_16), Trim$("EndX: " & EndX & " CameraEndX: " & Camera.Right), 260, 22, Yellow
    'RenderText font(Fonts.georgiaDec_16), Trim$("StartY: " & StartY & " CameraStartY: " & Camera.Top), 260, 38, Yellow
    'RenderText font(Fonts.georgiaDec_16), Trim$("EndY: " & EndY & " CameraEndX: " & Camera.bottom), 260, 54, Yellow
    
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
