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
Public Const Path_Projectile As String = "\data files\graphics\projectiles\"
Public Const Path_Surface As String = "\data files\graphics\surfaces\"
Public Const Path_Fog As String = "\data files\graphics\fog\"
Public Const Path_Captcha As String = "\data files\graphics\captchas\"
Public Const GFX_EXT As String = ".png"

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

Private Const COLOR_DEPTH_16_BIT As Long = D3DFMT_R5G6B5
Private Const COLOR_DEPTH_24_BIT As Long = D3DFMT_A8R8G8B8
Private Const COLOR_DEPTH_32_BIT As Long = D3DFMT_X8R8G8B8

Public Type TextureStruct
    Texture As Direct3DTexture8
    Data() As Byte
    Width As Long
    Height As Long
    RealWidth As Long
    RealHeight As Long
    UnloadTimer As Long
    Unload As Boolean
    Path As String
    Loaded As Boolean
End Type

Public Type TextureDataStruct
    Data() As Byte
    Width As Long
    Height As Long
    Unload As Boolean
    Path As String
    Loaded As Boolean
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
Private Const TEXTURE_NULL As Long = 0
Public CurrentTexture As Long

Public ScreenWidth As Long, ScreenHeight As Long
Public TileWidth As Long, TileHeight As Long
Public ScreenX As Long, ScreenY As Long
Public curResolution As Byte, isFullscreen As Boolean

Public Const DegreeToRadian As Single = 0.0174532919296
Public Const RadianToDegree As Single = 57.2958300962816

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
                Options.resolution = 0
                Options.Render = 0
                SaveOptions
                Call MsgBox("Could not initialize DirectX with hardware vertex processing.", vbCritical)
                Call DestroyGame
            End If
        Case 2 ' mixed
            If LoadDirectX(D3DCREATE_MIXED_VERTEXPROCESSING, hWnd) <> 0 Then
                Options.Fullscreen = 0
                Options.resolution = 0
                Options.Render = 0
                SaveOptions
                Call MsgBox("Could not initialize DirectX with mixed vertex processing.", vbCritical)
                Call DestroyGame
            End If
        Case 3 ' software
            If LoadDirectX(D3DCREATE_SOFTWARE_VERTEXPROCESSING, hWnd) <> 0 Then
                Options.Fullscreen = 0
                Options.resolution = 0
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
                        Options.resolution = 0
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
    'Call D3DDevice.SetStreamSource(0, DXVB, Len(Box(0)))
End Sub

Public Function LoadDirectX(ByVal BehaviourFlags As CONST_D3DCREATEFLAGS, ByVal hWnd As Long)
On Error GoTo ErrorInit

    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, BehaviourFlags, D3DWindow)
    Exit Function

ErrorInit:
    LoadDirectX = 1
End Function

Public Sub DestroyDX8()
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
    TextureBars = LoadTextureFile(App.path & Path_Graphics & "bars")
    TextureBlood = LoadTextureFile(App.path & Path_Graphics & "blood")
    TextureDirection = LoadTextureFile(App.path & Path_Graphics & "direction")
    TextureMisc = LoadTextureFile(App.path & Path_Graphics & "misc")
    TextureTarget = LoadTextureFile(App.path & Path_Graphics & "target")
    TextureShadow = LoadTextureFile(App.path & Path_Graphics & "shadow")
    TextureFader = LoadTextureFile(App.path & Path_Graphics & "fader")
    TextureBlank = LoadTextureFile(App.path & Path_Graphics & "blank")
    TextureWeather = LoadTextureFile(App.path & Path_Graphics & "weather")
    TextureWhite = LoadTextureFile(App.path & Path_Graphics & "white")
End Sub

Public Function LoadTextureFiles(ByRef Counter As Long, ByVal Path As String) As Long()
Dim Texture() As Long
Dim i As Long

    Counter = 1
    
    Do While Dir$(Path & Counter + 1 & GFX_EXT) <> vbNullString
        Counter = Counter + 1
    Loop
    
    ReDim Texture(0 To Counter)
    
    For i = 1 To Counter
        Texture(i) = LoadTextureFile(Path & i)
        DoEvents
    Next
    
    LoadTextureFiles = Texture
End Function

Public Function LoadTextureFile(ByVal Path As String, Optional ByVal Unload As Boolean = True, Optional ByVal Load As Boolean = True, Optional ByVal Ignore As Boolean = False) As Long
    Dim tempData As TextureDataStruct, Width As Long, Height As Long
    Dim Lugar As String
    

    Path = Path & GFX_EXT
    
    If Dir$(Path) = vbNullString And Not Ignore Then
        Call MsgBox("" & Path & """ could not be found.")
        End
    End If
    
    If Dir$(Path) = vbNullString Then
        Exit Function
    End If
    
    If Load Then
        tempData = LoadBitmap(Path)
        tempData.Unload = Unload
        tempData.Path = Path
        LoadTextureFile = LoadTexture(tempData)
    Else
        LoadTextureFile = PreloadTexture(Path)
    End If
End Function

Function GetNearestPOT(Value As Long) As Long
    Dim i As Long

    Do While 2 ^ i < Value
        i = i + 1
    Loop

    GetNearestPOT = 2 ^ i
End Function

Public Function LoadBitmap(ByVal Path As String) As TextureDataStruct
    Dim SourceBitmap As cGDIpImage, ConvertedBitmap As cGDIpImage, GDIGraphics As cGDIpRenderer, i As Long
    Dim newWidth As Long, newHeight As Long, ImageData() As Byte, fn As Long
    
    Set SourceBitmap = New cGDIpImage
    Call SourceBitmap.LoadPicture_FileName(Path, GDIToken)

    If SourceBitmap.Width = 0 Or SourceBitmap.Height = 0 Then
        Exit Function
    End If
    
    LoadBitmap.Height = SourceBitmap.Height
    LoadBitmap.Width = SourceBitmap.Width
    
    newWidth = GetNearestPOT(SourceBitmap.Width)
    newHeight = GetNearestPOT(SourceBitmap.Height)

    If newWidth <> SourceBitmap.Width Or newHeight <> SourceBitmap.Height Then
        Set ConvertedBitmap = New cGDIpImage
        Set GDIGraphics = New cGDIpRenderer
        i = GDIGraphics.CreateGraphicsFromImageClass(SourceBitmap)
        Call ConvertedBitmap.LoadPicture_FromNothing(newHeight, newWidth, i, GDIToken) 'This is no longer backwards and it now works.
        Call GDIGraphics.DestroyHGraphics(i)
        i = GDIGraphics.CreateGraphicsFromImageClass(ConvertedBitmap)
        Call GDIGraphics.AttachTokenClass(GDIToken)
        Call GDIGraphics.RenderImageClassToHGraphics(SourceBitmap, i)
        Call ConvertedBitmap.SaveAsPNG(ImageData)
        GDIGraphics.DestroyHGraphics (i)
        LoadBitmap.Data = ImageData
        Set ConvertedBitmap = Nothing
        Set GDIGraphics = Nothing
        Set SourceBitmap = Nothing
        'SaveFile Path & ".png", ImageData
    Else
        Call SourceBitmap.SaveAsPNG(ImageData)
        LoadBitmap.Data = ImageData
        'SaveFile Path & ".png", ImageData
        Set SourceBitmap = Nothing
    End If
End Function

Public Function LoadTexture(ByRef tempData As TextureDataStruct, Optional ByVal Path As String = "") As Long
    If AryCount(tempData.Data) = 0 Then
        Exit Function
    End If
    
    mTextures = mTextures + 1
    LoadTexture = mTextures
    ReDim Preserve mTexture(1 To mTextures) As TextureStruct
    mTexture(mTextures).RealWidth = tempData.Width
    mTexture(mTextures).RealHeight = tempData.Height
    mTexture(mTextures).Width = ByteToInt(tempData.Data(18), tempData.Data(19))
    mTexture(mTextures).Height = ByteToInt(tempData.Data(22), tempData.Data(23))

    mTexture(mTextures).Data = tempData.Data
    mTexture(mTextures).Unload = tempData.Unload
    mTexture(mTextures).Path = tempData.Path
    mTexture(mTextures).Loaded = True
    
End Function

Public Function PreloadTexture(ByVal Path As String) As Long
    mTextures = mTextures + 1
    PreloadTexture = mTextures
    ReDim Preserve mTexture(1 To mTextures) As TextureStruct
    mTexture(mTextures).Unload = True
    mTexture(mTextures).Path = Path
End Function

Public Sub UnloadGFX()
    Dim i As Long
    
    For i = 1 To mTextures
        If mTexture(i).Unload Then
            If mTexture(i).UnloadTimer > 0 And mTexture(i).UnloadTimer <= Tick Then
                Set mTexture(i).Texture = Nothing
                mTexture(i).UnloadTimer = 0
            End If
        End If
    Next
End Sub

Public Sub CheckGFX()
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then
        Do While D3DDevice.TestCooperativeLevel = D3DERR_DEVICELOST
           DoEvents
        Loop
        
        Call ResetGFX
        
        ' This is to fix chat buggy after device lost
        If InGame Then
            'ChatArrayUbound = 0
            'UpdateChatArray ChatChannel
        End If
    End If
End Sub

Public Sub ResetGFX()
Dim Temp() As TextureDataStruct
Dim i As Long, N As Long

    N = mTextures
    ReDim Temp(1 To N)
    
    Erase mTexture
    mTextures = 0
    
    InitDX8 frmMain.hWnd
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
    
    LoadTextures
End Sub

Public Function SetTexture(ByVal TextureNum As Long) As Boolean
    On Error GoTo finish:
    
    If TextureNum < 1 Or TextureNum > mTextures Then Exit Function
    
    ' Exit out early
    If mTexture(TextureNum).Texture Is Nothing Then
        If mTexture(TextureNum).Loaded = False Then
            If InGame And (Thread Or GameLooptmr <= Tick) Then Exit Function
            Dim tempData As TextureDataStruct
            tempData = LoadBitmap(mTexture(TextureNum).Path)
            mTexture(TextureNum).RealWidth = tempData.Width
            mTexture(TextureNum).RealHeight = tempData.Height
            mTexture(TextureNum).Width = ByteToInt(tempData.Data(18), tempData.Data(19))
            mTexture(TextureNum).Height = ByteToInt(tempData.Data(22), tempData.Data(23))
            mTexture(TextureNum).Data = tempData.Data
            mTexture(TextureNum).Loaded = True
        End If
        
        Set mTexture(TextureNum).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, mTexture(TextureNum).Data(0), AryCount(mTexture(TextureNum).Data), mTexture(TextureNum).Width, mTexture(TextureNum).Height, D3DX_DEFAULT, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    End If
    
    mTexture(TextureNum).UnloadTimer = Tick + 30000
    If TextureNum <> CurrentTexture Then
        Call D3DDevice.SetTexture(0, mTexture(TextureNum).Texture)
    End If
    CurrentTexture = TextureNum
    SetTexture = True
    Exit Function
    
finish:
    ' Ignore it and clear memory - this error is too much memory and nothing can be done about it with our current implementation of DX
    SetTexture = False
    Set mTexture(TextureNum).Texture = Nothing
    mTexture(TextureNum).UnloadTimer = 0
End Function

Public Sub RenderTexture(Texture As Long, ByVal X As Long, ByVal Y As Long, ByVal sX As Single, ByVal sY As Single, ByVal w As Long, ByVal h As Long, ByVal SW As Single, ByVal SH As Single, Optional ByVal Colour As Long = -1, Optional ByVal offset As Boolean = False, Optional ByVal Degrees As Single = 0)
    If SetTexture(Texture) Then
        RenderGeom X, Y, sX, sY, w, h, SW, SH, Colour, offset, Degrees
    End If
End Sub

Public Sub RenderGeom(ByVal X As Long, ByVal Y As Long, ByVal sX As Single, ByVal sY As Single, ByVal w As Long, ByVal h As Long, ByVal SW As Single, ByVal SH As Single, Optional ByVal Colour As Long = -1, Optional ByVal offset As Boolean = False, Optional ByVal Degrees As Single = 0)
Dim i As Long

    If CurrentTexture = 0 Then Exit Sub
    If w = 0 Then Exit Sub
    If h = 0 Then Exit Sub
    If SW = 0 Then Exit Sub
    If SH = 0 Then Exit Sub
    
    If mClip.Right <> 0 Then
        If mClip.top <> 0 Then
            If mClip.Left > X Then
                sX = sX + (mClip.Left - X) / (w / SW)
                SW = SW - (mClip.Left - X) / (w / SW)
                w = w - (mClip.Left - X)
                X = mClip.Left
            End If
            
            If mClip.top > Y Then
                sY = sY + (mClip.top - Y) / (h / SH)
                SH = SH - (mClip.top - Y) / (h / SH)
                h = h - (mClip.top - Y)
                Y = mClip.top
            End If
            
            If mClip.Right < X + w Then
                SW = SW - (X + w - mClip.Right) / (w / SW)
                w = -X + mClip.Right
            End If
            
            If mClip.Bottom < Y + h Then
                SH = SH - (Y + h - mClip.Bottom) / (h / SH)
                h = -Y + mClip.Bottom
            End If
            
            If w <= 0 Then Exit Sub
            If h <= 0 Then Exit Sub
            If SW <= 0 Then Exit Sub
            If SH <= 0 Then Exit Sub
        End If
    End If
    
    Call GeomCalc(Box, CurrentTexture, X, Y, w, h, sX, sY, SW, SH, Colour, Degrees)
    Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), Len(Box(0)))
End Sub

Public Sub GeomCalc(ByRef Geom() As Vertex, ByVal TextureNum As Long, ByVal X As Single, ByVal Y As Single, ByVal w As Integer, ByVal h As Integer, ByVal sX As Single, ByVal sY As Single, ByVal SW As Single, ByVal SH As Single, ByVal Colour As Long, Optional ByVal Degrees As Single = 0)
    Dim RadAngle As Single ' The angle in Radians
    Dim CenterX As Single, CenterY As Single
    Dim NewX As Single, NewY As Single
    Dim SinRad As Single, CosRad As Single, i As Byte
    
    SW = (SW + sX) / mTexture(TextureNum).Width + 0.000003
    SH = (SH + sY) / mTexture(TextureNum).Height + 0.000003
    sX = sX / mTexture(TextureNum).Width + 0.000003
    sY = sY / mTexture(TextureNum).Height + 0.000003
    Geom(0) = MakeVertex(X, Y, 0, 1, Colour, 1, sX, sY)
    Geom(1) = MakeVertex(X + w, Y, 0, 1, Colour, 0, SW, sY)
    Geom(2) = MakeVertex(X, Y + h, 0, 1, Colour, 0, sX, SH)
    Geom(3) = MakeVertex(X + w, Y + h, 0, 1, Colour, 0, SW, SH)
    
        ' Check if a rotation is required
    If Degrees <> 0 And Degrees <> 360 Then

        ' Converts the angle to rotate by into radians
        RadAngle = Degrees * DegreeToRadian

        ' Set the CenterX and CenterY values
        CenterX = X + (w * 0.5)
        CenterY = Y + (h * 0.5)

        ' Pre-calculate the cosine and sine of the radiant
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)

        ' Loops through the passed vertex buffer
        For i = 0 To 3

            ' Calculates the new X and Y co-ordinates of the vertices for the given angle around the center co-ordinates
            NewX = CenterX + (Box(i).X - CenterX) * CosRad - (Box(i).Y - CenterY) * SinRad
            NewY = CenterY + (Box(i).Y - CenterY) * CosRad + (Box(i).X - CenterX) * SinRad

            ' Applies the new co-ordinates to the buffer
            Box(i).X = NewX
            Box(i).Y = NewY
        Next

    End If
End Sub

Private Function MakeVertex(X As Single, Y As Single, z As Single, RHW As Single, Colour As Long, Specular As Long, tu As Single, tv As Single) As Vertex
    MakeVertex.X = X
    MakeVertex.Y = Y
    MakeVertex.z = z
    MakeVertex.RHW = RHW
    MakeVertex.Colour = Colour
    MakeVertex.tu = tu
    MakeVertex.tv = tv
End Function

' GDI rendering
Public Sub GDIRenderAnimation()
    Dim i As Long, Animationnum As Long, ShouldRender As Boolean, width As Long, height As Long, looptime As Long, FrameCount As Long
    Dim sX As Long, sY As Long, sRECT As RECT
    sRECT.top = 0
    sRECT.Bottom = 192
    sRECT.Left = 0
    sRECT.Right = 192

    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).Value

        If Animationnum <= 0 Or Animationnum > CountAnim Then
            ' don't render lol
        Else
            looptime = frmEditor_Animation.scrlLoopTime(i)

            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            ShouldRender = False

            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= getTime Then

                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If

                AnimEditorTimer(i) = getTime
                ShouldRender = True
            End If

            If ShouldRender Then
                If frmEditor_Animation.scrlFrameCount(i).Value > 0 Then
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
                    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Animation.picSprite(i).hWnd, ByVal 0)
                End If
            End If
        End If

    Next

End Sub

Public Sub GDIRenderChar(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim height As Long, width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Char Then Exit Sub
    Height = 32
    Width = 32
    sRECT.top = 0
    sRECT.Bottom = sRECT.top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    RenderTexture TextureChar(sprite), 0, 0, 0, 0, width, height, width, height
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderFace(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim height As Long, width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Face Then Exit Sub
    Height = mTexture(Tex_Face(sprite)).RealHeight
    Width = mTexture(Tex_Face(sprite)).RealWidth

    If height = 0 Or width = 0 Then
        height = 1
        width = 1
    End If

    sRECT.top = 0
    sRECT.Bottom = sRECT.top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle TextureFace(sprite), 0, 0, 0, 0, width, height, width, height, width, height
    RenderTexture TextureFace(sprite), 0, 0, 0, 0, width, height, width, height
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderTileset()
    Dim height As Long, width As Long, tileSet As Byte, sRECT As RECT
    ' find tileset number
    tileSet = frmEditor_Map.scrlTileSet.Value

    ' exit out if doesn't exist
    If tileSet <= 0 Or tileSet > Count_Tileset Then Exit Sub
    Height = mTexture(Tex_Tileset(tileSet)).RealHeight
    Width = mTexture(Tex_Tileset(tileSet)).RealWidth

    If height = 0 Or width = 0 Then
        height = 1
        width = 1
    End If

    frmEditor_Map.picBackSelect.Width = Width
    frmEditor_Map.picBackSelect.Height = Height
    sRECT.top = 0
    sRECT.Bottom = Height
    sRECT.Left = 0
    sRECT.Right = Width

    ' change selected shape for autotiles
    If frmEditor_Map.scrlAutotile.Value > 0 Then

        Select Case frmEditor_Map.scrlAutotile.Value

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
    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Map.picBackSelect.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderItem(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim height As Long, width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Item Then Exit Sub
    Height = mTexture(Tex_Item(sprite)).RealHeight
    Width = mTexture(Tex_Item(sprite)).RealWidth
    sRECT.top = 0
    sRECT.Bottom = 32
    sRECT.Left = 0
    sRECT.Right = 32
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle TextureItem(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture TextureItem(sprite), 0, 0, 0, 0, 32, 32, 32, 32
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderItemPaperdoll(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim height As Long, width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Paperdoll Then Exit Sub
    Height = mTexture(Tex_Paperdoll(sprite)).RealHeight
    Width = mTexture(Tex_Paperdoll(sprite)).RealWidth
    sRECT.top = 0
    sRECT.Bottom = 72
    sRECT.Left = 0
    sRECT.Right = 144
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle TextureItem(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture TexturePaperdoll(sprite), 0, 0, 0, 0, 144, 72, 144, 72
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderResource(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim height As Long, width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Resource Then Exit Sub
    Height = mTexture(Tex_Resource(sprite)).RealHeight
    Width = mTexture(Tex_Resource(sprite)).RealWidth
    sRECT.top = 0
    sRECT.Bottom = 152
    sRECT.Left = 0
    sRECT.Right = 152
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    RenderTexture TextureResource(sprite), 0, 0, 0, 0, width, height, width, height
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub


Public Sub GDIRenderSpell(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim height As Long, width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Spellicon Then Exit Sub
    Height = mTexture(Tex_Spellicon(sprite)).RealHeight
    Width = mTexture(Tex_Spellicon(sprite)).RealWidth

    If height = 0 Or width = 0 Then
        height = 1
        width = 1
    End If

    sRECT.top = 0
    sRECT.Bottom = Height
    sRECT.Left = 0
    sRECT.Right = Width
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle TextureSpellIcon(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture TextureSpellIcon(sprite), 0, 0, 0, 0, 32, 32, 32, 32
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderSpellProjectile(ByRef picBox As PictureBox, ByVal sprite As Long)
    Dim Height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Projectile Then Exit Sub
    Height = mTexture(Tex_Projectile(sprite)).RealHeight
    Width = mTexture(Tex_Projectile(sprite)).RealWidth
    sRECT.top = 0
    sRECT.Bottom = 64
    sRECT.Left = 0
    sRECT.Right = 64
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle Tex_Item(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Projectile(sprite), 0, 0, 0, 0, 64, 64, 64, 64
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

' Directional blocking
Public Sub DrawDirection(ByVal X As Long, ByVal Y As Long)
    Dim i As Long, top As Long, Left As Long
    ' render grid
    top = 24
    Left = 0
    'EngineRenderRectangle Tex_Direction, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), left, top, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Direction, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), Left, top, 32, 32, 32, 32

    ' render dir blobs
    For i = 1 To 4
        left = (i - 1) * 8

        ' find out whether render blocked or not
        If Not isDirBlocked(Map.TileData.Tile(X, Y).DirBlock, CByte(i)) Then
            top = 8
        Else
            top = 16
        End If

        'render!
        'EngineRenderRectangle Tex_Direction, ConvertMapX(x * PIC_X) + DirArrowX(i), ConvertMapY(y * PIC_Y) + DirArrowY(i), left, top, 8, 8, 8, 8, 8, 8
        RenderTexture Tex_Direction, ConvertMapX(X * PIC_X) + DirArrowX(i), ConvertMapY(Y * PIC_Y) + DirArrowY(i), Left, top, 8, 8, 8, 8
    Next

End Sub

Public Sub DrawFade()
    RenderTexture TextureBlank, 0, 0, 0, 0, ScreenWidth, ScreenHeight, 32, 32, DX8Colour(White, fadeAlpha)
End Sub

Public Sub DrawFog()
    Dim fogNum As Long, Colour As Long, X As Long, Y As Long, RenderState As Long

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

    For X = 0 To ((Map.MapData.MaxX * 32) / 256) + 1
        For Y = 0 To ((Map.MapData.MaxY * 32) / 256) + 1
            RenderTexture Tex_Fog(fogNum), ConvertMapX((X * 256) + fogOffsetX), ConvertMapY((Y * 256) + fogOffsetY), 0, 0, 256, 256, 256, 256, Colour
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
            RenderTexture Tex_Weather, ConvertMapX(WeatherParticle(i).X), ConvertMapY(WeatherParticle(i).Y), SpriteLeft * 32, 0, 32, 32, 32, 32, -1
        End If
    Next
End Sub

Public Sub DrawAutoTile(ByVal layernum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim yOffset As Long, xOffset As Long

    ' calculate the offset
    Select Case Map.TileData.Tile(X, Y).Autotile(layernum)

        Case AUTOTILE_WATERFALL
            yOffset = (waterfallFrame - 1) * 32

        Case AUTOTILE_ANIM
            xOffset = autoTileFrame * 64

        Case AUTOTILE_CLIFF
            yOffset = -32
    End Select

    ' Draw the quarter
    RenderTexture Tex_Tileset(Map.TileData.Tile(X, Y).Layer(layernum).tileSet), destX, destY, Autotile(X, Y).Layer(layernum).srcX(quarterNum) + xOffset, Autotile(X, Y).Layer(layernum).srcY(quarterNum) + yOffset, 16, 16, 16, 16
End Sub

Sub DrawTileSelection()
    Dim tileSet As Byte
    ' find tileset number
    tileSet = frmEditor_Map.scrlTileSet.Value
    
    ' exit out if doesn't exist
    If tileSet <= 0 Or tileSet > CountTileset Then Exit Sub

    If frmEditor_Map.scrlAutotile.Value > 0 Then
        RenderTexture Tex_Tileset(tileSet), ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), shpSelectedLeft, shpSelectedTop, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 200)
    Else
        RenderTexture TextureTileset(tileSet), ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), shpSelectedLeft, shpSelectedTop, shpSelectedWidth, shpSelectedHeight, shpSelectedWidth, shpSelectedHeight
    End If
End Sub

' Rendering Procedures
Public Sub DrawMapTile(ByVal X As Long, ByVal Y As Long)
Dim i As Long, tileSet As Long, sX As Long, sY As Long

    With Map.TileData.Tile(X, Y)
        ' draw the map
        For i = MapLayer.Ground To MapLayer.Mask2
            ' skip tile if tileset isn't set
            If Autotile(X, Y).Layer(i).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.Layer(i).tileSet), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(i).X * 32, .Layer(i).Y * 32, 32, 32, 32, 32
            ElseIf Autotile(X, Y).Layer(i).RenderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
            ElseIf Autotile(X, Y).Layer(i).RenderState = RENDER_STATE_APPEAR Then
                ' check if it's fading
                If TempTile(X, Y).fadeAlpha(i) > 0 Then
                    ' render it
                    tileSet = Map.TileData.Tile(X, Y).Layer(i).tileSet
                    sX = Map.TileData.Tile(X, Y).Layer(i).X
                    sY = Map.TileData.Tile(X, Y).Layer(i).Y
                    RenderTexture Tex_Tileset(tileSet), ConvertMapX(X * 32), ConvertMapY(Y * 32), sX * 32, sY * 32, 32, 32, 32, 32, DX8Colour(White, TempTile(X, Y).fadeAlpha(i))
                End If
            End If
        Next
    End With
End Sub

Public Sub DrawMapFringeTile(ByVal X As Long, ByVal Y As Long)
    Dim i As Long

    With Map.TileData.Tile(X, Y)
        ' draw the map
        For i = MapLayer.Fringe To MapLayer.Fringe2

            ' skip tile if tileset isn't set
            If Autotile(X, Y).Layer(i).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.Layer(i).tileSet), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(i).X * 32, .Layer(i).Y * 32, 32, 32, 32, 32
            ElseIf Autotile(X, Y).Layer(i).RenderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
            End If
        Next
    End With
End Sub

Public Sub DrawHotbar()
    Dim Xo As Long, Yo As Long, width As Long, height As Long, i As Long, t As Long, sS As String
    
    Xo = Windows(GetWindowIndex("winHotbar")).Window.Left
    Yo = Windows(GetWindowIndex("winHotbar")).Window.top
    
    ' render start + end wood
    RenderTexture TextureGUI(40), Xo - 1, Yo + 3, 0, 0, 11, 26, 11, 26
    RenderTexture TextureGUI(40), Xo + 407, Yo + 3, 0, 0, 11, 26, 11, 26
    
    For i = 1 To MAX_HOTBAR
        Xo = Windows(GetWindowIndex("winHotbar")).Window.Left + HotbarLeft + ((i - 1) * HotbarOffsetX)
        Yo = Windows(GetWindowIndex("winHotbar")).Window.top + HotbarTop
        Width = 36
        Height = 36
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
                    If Len(Item(Hotbar(i).Slot).Name) > 0 And Item(Hotbar(i).Slot).pic > 0 Then
                        RenderTexture Tex_Item(Item(Hotbar(i).Slot).pic), Xo, Yo, 0, 0, 32, 32, 32, 32
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
Dim X As Long, Y As Long, tileSet As Long, sX As Long, sY As Long, layernum As Long

    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            For layernum = MapLayer.Ground To MapLayer.Mask
                ' check if it's fading
                If TempTile(X, Y).fadeAlpha(layernum) > 0 Then
                    ' render it
                    tileSet = Map.TileData.Tile(X, Y).Layer(layernum).tileSet
                    sX = Map.TileData.Tile(X, Y).Layer(layernum).X
                    sY = Map.TileData.Tile(X, Y).Layer(layernum).Y
                    RenderTexture Tex_Tileset(tileSet), ConvertMapX(X * 32), ConvertMapY(Y * 32), sX * 32, sY * 32, 32, 32, 32, 32, DX8Colour(White, TempTile(X, Y).fadeAlpha(layernum))
                End If
            Next
        Next
    Next
End Sub

Public Sub DrawCharacter()
    Dim Xo As Long, Yo As Long, width As Long, height As Long, i As Long, sprite As Long, ItemNum As Long, ItemPic As Long
    Dim xEquipBar As Long, yEquipBar As Long, yOffSetEquip As Long
    
    Xo = Windows(GetWindowIndex("winCharacter")).Window.Left
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
            ItemPic = Tex_Item(Item(ItemNum).pic)
        Else
            ' no item equiped - use blank image
            ItemPic = TextureGUI(45 + i)
        End If
        
        Yo = Windows(GetWindowIndex("winCharacter")).Window.top + EqTop + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
        Xo = Windows(GetWindowIndex("winCharacter")).Window.Left + EqLeft

        RenderTexture ItemPic, Xo, Yo, 0, 0, 32, 32, 32, 32
    Next
End Sub

Public Sub DrawSkills()
    Dim Xo As Long, Yo As Long, Width As Long, Height As Long, i As Long, Y As Long, spellnum As Long, spellPic As Long, X As Long, top As Long, Left As Long
    
    Xo = Windows(GetWindowIndex("winSkills")).Window.Left
    Yo = Windows(GetWindowIndex("winSkills")).Window.top
    
    width = Windows(GetWindowIndex("winSkills")).Window.width
    height = Windows(GetWindowIndex("winSkills")).Window.height
    
    ' render green
    RenderTexture TextureGUI(34), Xo + 4, Yo + 23, 0, 0, width - 8, height - 27, 4, 4
    
    width = 76
    height = 76
    
    Y = Yo + 23
    ' render grid - row
    For i = 1 To 4
        If i = 4 Then Height = 42
        RenderTexture Tex_GUI(35), Xo + 4, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 80, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 156, Y, 0, 0, 42, Height, 42, Height
        Y = Y + 76
    Next
    
    ' actually draw the icons
    For i = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(i).Spell
        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            ' not dragging?
            If Not (DragBox.origin = originSpells And DragBox.Slot = i) Then
                spellPic = Spell(spellnum).icon
    
                If spellPic > 0 And spellPic <= Count_Spellicon Then
                    top = Yo + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    Left = Xo + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
    
                    RenderTexture Tex_Spellicon(spellPic), Left, top, 0, 0, 32, 32, 32, 32
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
    Dim Xo As Long, Yo As Long, Width As Long, Height As Long, i As Long, Y As Long
    
    Xo = Windows(GetWindowIndex("winOffer")).Window.Left + 475
    Yo = Windows(GetWindowIndex("winOffer")).Window.top
    
    width = 45
    height = 45
    
    Y = Yo
    
    For i = 1 To 3
        If inOffer(i) > 0 Then
            RenderDesign DesignTypes.desWin_Desc, Xo, Y, Width, Height
            RenderText font(Fonts.georgia_16), "i", Xo + 21, Y + 15, Grey
            Y = Y + 37
        End If
    Next
End Sub

Public Sub DrawShopBackground()
    Dim Xo As Long, Yo As Long, Width As Long, Height As Long, i As Long, Y As Long
    
    Xo = Windows(GetWindowIndex("winShop")).Window.Left
    Yo = Windows(GetWindowIndex("winShop")).Window.top
    Width = Windows(GetWindowIndex("winShop")).Window.Width
    Height = Windows(GetWindowIndex("winShop")).Window.Height
    
    ' render green
    RenderTexture TextureGUI(34), Xo + 4, Yo + 23, 0, 0, width - 8, height - 27, 4, 4
    
    width = 76
    height = 76
    
    Y = Yo + 23
    ' render grid - row
    For i = 1 To 3
        If i = 3 Then Height = 42
        RenderTexture Tex_GUI(35), Xo + 4, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 80, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 156, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 232, Y, 0, 0, 42, Height, 42, Height
        Y = Y + 76
    Next
    ' render bottom wood
    RenderTexture Tex_GUI(1), Xo + 4, Y - 34, 0, 0, 270, 72, 270, 72
End Sub

Public Sub DrawShop()
Dim Xo As Long, Yo As Long, ItemPic As Long, ItemNum As Long, Amount As Long, i As Long, top As Long, Left As Long, Y As Long, X As Long, Colour As Long

    If InShop = 0 Then Exit Sub
    
    Xo = Windows(GetWindowIndex("winShop")).Window.Left
    Yo = Windows(GetWindowIndex("winShop")).Window.top
    
    If Not shopIsSelling Then
        ' render the shop items
        For i = 1 To MAX_TRADES
            ItemNum = Shop(InShop).TradeItem(i).Item
            
            ' draw early
            top = Yo + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
            Left = Xo + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
            ' draw selected square
            If shopSelectedSlot = i Then RenderTexture Tex_GUI(61), Left, top, 0, 0, 32, 32, 32, 32
            
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).pic
                If ItemPic > 0 And ItemPic <= Count_Item Then
                    ' draw item
                    RenderTexture Tex_Item(ItemPic), Left, top, 0, 0, 32, 32, 32, 32
                End If
            End If
        Next
    Else
        ' render the shop items
        For i = 1 To MAX_TRADES
            ItemNum = GetPlayerInvItemNum(MyIndex, i)
            
            ' draw early
            top = Yo + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
            Left = Xo + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
            ' draw selected square
            If shopSelectedSlot = i Then RenderTexture Tex_GUI(61), Left, top, 0, 0, 32, 32, 32, 32
            
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).pic
                If ItemPic > 0 And ItemPic <= Count_Item Then

                    ' draw item
                    RenderTexture Tex_Item(ItemPic), Left, top, 0, 0, 32, 32, 32, 32
                    
                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        Y = top + 21
                        X = Left + 1
                        Amount = CStr(GetPlayerInvItemValue(MyIndex, i))
                        
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
    Dim Xo As Long, Yo As Long, Width As Long, Height As Long, i As Long, Y As Long, X As Long
    
    Xo = Windows(GetWindowIndex("winTrade")).Window.Left
    Yo = Windows(GetWindowIndex("winTrade")).Window.top
    Width = Windows(GetWindowIndex("winTrade")).Window.Width
    Height = Windows(GetWindowIndex("winTrade")).Window.Height
    
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
    Width = 76
    Height = 76
    Y = Yo + 41
    For i = 1 To 4
        If i = 4 Then Height = 38
        RenderTexture Tex_GUI(35), Xo + 4 + 5, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 80 + 5, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 156 + 5, Y, 0, 0, 42, Height, 42, Height
        Y = Y + 76
    Next
    
    ' right
    Width = 76
    Height = 76
    Y = Yo + 41
    For i = 1 To 4
        If i = 4 Then Height = 38
        RenderTexture Tex_GUI(35), Xo + 4 + 205, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 80 + 205, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), Xo + 156 + 205, Y, 0, 0, 42, Height, 42, Height
        Y = Y + 76
    Next
End Sub

Sub DrawYourTrade()
Dim i As Long, ItemNum As Long, ItemPic As Long, top As Long, Left As Long, Colour As Long, Amount As String, X As Long, Y As Long
Dim Xo As Long, Yo As Long

    Xo = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).Left
    Yo = Windows(GetWindowIndex("winTrade")).Window.top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).top
    
    ' your items
    For i = 1 To MAX_INV
        ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).pic
            If ItemPic > 0 And ItemPic <= Count_Item Then
                top = Yo + TradeTop + ((TradeOffsetY + 32) * ((i - 1) \ TradeColumns))
                Left = Xo + TradeLeft + ((TradeOffsetX + 32) * (((i - 1) Mod TradeColumns)))

                ' draw icon
                RenderTexture Tex_Item(ItemPic), Left, top, 0, 0, 32, 32, 32, 32
                
                ' If item is a stack - draw the amount you have
                If TradeYourOffer(i).Value > 1 Then
                    Y = top + 21
                    X = Left + 1
                    Amount = CStr(TradeYourOffer(i).Value)
                    
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
Dim i As Long, ItemNum As Long, ItemPic As Long, top As Long, Left As Long, Colour As Long, Amount As String, X As Long, Y As Long
Dim Xo As Long, Yo As Long

    Xo = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picTheir")).Left
    Yo = Windows(GetWindowIndex("winTrade")).Window.top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picTheir")).top

    ' their items
    For i = 1 To MAX_INV
        ItemNum = TradeTheirOffer(i).num
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).pic
            If ItemPic > 0 And ItemPic <= Count_Item Then
                top = Yo + TradeTop + ((TradeOffsetY + 32) * ((i - 1) \ TradeColumns))
                Left = Xo + TradeLeft + ((TradeOffsetX + 32) * (((i - 1) Mod TradeColumns)))

                ' draw icon
                RenderTexture Tex_Item(ItemPic), Left, top, 0, 0, 32, 32, 32, 32
                
                ' If item is a stack - draw the amount you have
                If TradeTheirOffer(i).Value > 1 Then
                    Y = top + 21
                    X = Left + 1
                    Amount = CStr(TradeTheirOffer(i).Value)
                    
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
    Dim X As Long, Y As Long, Xo As Long, Yo As Long, Width As Long, Height As Long
    Dim i As Long, ItemNum As Long, ItemPic As Long

    Dim Left As Long, top As Long
    Dim Colour As Long, skipItem As Boolean, Amount As Long, tmpItem As Long

    Xo = Windows(GetWindowIndex("winBank")).Window.Left
    Yo = Windows(GetWindowIndex("winBank")).Window.top
    Width = Windows(GetWindowIndex("winBank")).Window.Width
    Height = Windows(GetWindowIndex("winBank")).Window.Height
    
    ' render green
    RenderTexture TextureGUI(34), Xo + 4, Yo + 23, 0, 0, width - 8, height - 27, 4, 4

    width = 76
    height = 76

    Y = Yo + 23
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
            If Not (DragBox.Origin = origin_Bank And DragBox.Slot = i) Then
                ItemPic = Item(ItemNum).pic


                If ItemPic > 0 And ItemPic <= Count_Item Then
                    top = Yo + BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                    Left = Xo + BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))

                    ' draw icon
                    RenderTexture Tex_Item(ItemPic), Left, top, 0, 0, 32, 32, 32, 32

                    ' If item is a stack - draw the amount you have
                    If Bank.Item(i).Value > 1 Then
                        Y = top + 21
                        X = Left + 1
                        Amount = Bank.Item(i).Value

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
    Dim Xo As Long, Yo As Long, Width As Long, Height As Long, i As Long, Y As Long, ItemNum As Long, ItemPic As Long, X As Long, top As Long, Left As Long, Amount As String
    Dim Colour As Long, skipItem As Boolean, amountModifier  As Long, tmpItem As Long
    
    Xo = Windows(GetWindowIndex("winInventory")).Window.Left
    Yo = Windows(GetWindowIndex("winInventory")).Window.top
    Width = Windows(GetWindowIndex("winInventory")).Window.Width
    Height = Windows(GetWindowIndex("winInventory")).Window.Height
    
    ' render green
    RenderTexture TextureGUI(34), Xo + 4, Yo + 23, 0, 0, width - 8, height - 27, 4, 4
    
    width = 76
    height = 76
    
    Y = Yo + 23
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
    Dim Xo As Long, Yo As Long, Width As Long, Height As Long, i As Long, X As Long, Y As Long, ItemNum As Long, ItemPic As Long, top As Long, Left As Long, Amount As String
    Dim Colour As Long, skipItem As Boolean, amountModifier  As Long, tmpItem As Long
    
    Xo = Windows(GetWindowIndex("winPlayerQuests")).Window.Left
    Yo = Windows(GetWindowIndex("winPlayerQuests")).Window.top
    Width = Windows(GetWindowIndex("winPlayerQuests")).Window.Width
    Height = Windows(GetWindowIndex("winPlayerQuests")).Window.Height
    
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
        
        If Button_MissionActive <> 0 Then
            If Player(MyIndex).Mission(Button_MissionActive).ID <= 0 Then Exit Sub
            ItemNum = Mission(Player(MyIndex).Mission(Button_MissionActive).ID).RewardItem(i).ItemNum
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
                If ItemPic > 0 And ItemPic <= CountItem Then

                    ' draw icon
                    RenderTexture TextureItem(ItemPic), x + 4, Yo + 370, 0, 0, 32, 32, 32, 32

                    ' If item is a stack - draw the amount you have
                    If Mission(Player(MyIndex).Mission(Button_MissionActive).ID).RewardItem(i).ItemAmount > 1 Then
                        Y = Yo + 370 + 21
                        X = X + 4 + 1
                        Amount = Mission(Player(MyIndex).Mission(Button_MissionActive).ID).RewardItem(i).ItemAmount

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
        If .timer + 5000 < getTime Then
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
        If Player(Index).AttackTimer + (attackspeed / 2) > getTime Then
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

        If Player(Index).AnimTimer + 100 <= getTime Then
            Player(Index).Anim = Player(Index).Anim + 1

            If Player(Index).Anim >= 3 Then Player(Index).Anim = 0
            Player(Index).AnimTimer = getTime
        End If

        Anim = Player(Index).Anim
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)

        If .AttackTimer + attackspeed < getTime Then
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
        .top = SpriteTop * (mTexture(Tex_Char(sprite)).RealHeight / 4)
        .Height = (mTexture(Tex_Char(sprite)).RealHeight / 4)
        .Left = Anim * (mTexture(Tex_Char(sprite)).RealWidth / 4)
        .Width = (mTexture(Tex_Char(sprite)).RealWidth / 4)
    End With

    ' Calculate the X
    x = GetPlayerX(Index) * PIC_X + Player(Index).xOffset - ((mTexture(TextureChar(sprite)).RealWidth / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (mTexture(TextureChar(sprite)).RealHeight) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - ((mTexture(TextureChar(sprite)).RealHeight / 4) - 32) - 4
    Else
        ' Proceed as normal
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - 4
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

Public Sub DrawPaperdoll(ByVal Index As Long, ByVal sprite As Long, ByVal X2 As Long, Y2 As Long, rec As GeomRec)
    Dim X As Long, Y As Long
    Dim Width As Long, Height As Long

    If sprite < 1 Or sprite > CountPaperdoll Then Exit Sub

    Width = (rec.Width - rec.Left)
    Height = (rec.Height - rec.top)
    
    RenderTexture TexturePaperdoll(sprite), ConvertMapX(X2), ConvertMapY(Y2), rec.Left, rec.top, rec.Width, rec.Height, rec.Width, rec.Height, D3DColorRGBA(255, 255, 255, 255)
End Sub

Public Sub DrawProjectile(ByVal i As Long)
    Dim Angle As Long, X As Long, Y As Long, N As Long, SpriteTop As Byte
    Dim Xo As Long, Yo As Long
    Dim sRECT As RECT, Anim As Long

    If i < 0 Or i > MAX_PROJECTILE_MAP Then Exit Sub
    
    ' ****** Create Particle ******
    With MapProjectile(i)
        If .Graphic > 0 Then
            If .speed < 5000 Then
        
                ' ****** Update Position ******
                Angle = DegreeToRadian * Engine_GetAngle(.X, .Y, .tx, .ty)
                .X = .X + (Sin(Angle) * ElapsedTime * (.speed / 1000))
                .Y = .Y - (Cos(Angle) * ElapsedTime * (.speed / 1000))
                
                ' ****** Update Rotation ******
                If .RotateSpeed > 0 Then
                    .Rotate = .Rotate + (.RotateSpeed * ElapsedTime * 0.01)
                    Do While .Rotate > 360
                        .Rotate = .Rotate - 360
                    Loop
                End If
                
                If Abs(.X - .tx) < 60 Then
                    If Abs(.Y - .ty) < 60 Then
                        If .curAnim < 9 Then
                            .curAnim = 9
                        End If
                    End If
                End If
            Else
                If Tick + 120 >= .Duration Then
                    .curAnim = 9
                End If
            End If
            
            
            
            Anim = .curAnim
            Xo = .X
            Yo = .Y
            
            With sRECT
                If MapProjectile(i).IsDirectional Then
                    Select Case MapProjectile(i).Direction
                        Case DIR_UP
                            SpriteTop = 0
                        Case DIR_DOWN
                            SpriteTop = 1
                        Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                            SpriteTop = 2
                        Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                            SpriteTop = 3
                    End Select
                    
                    .top = SpriteTop * (mTexture(Tex_Projectile(MapProjectile(i).Graphic)).RealHeight / 4)
                    .Bottom = .top + (mTexture(Tex_Projectile(MapProjectile(i).Graphic)).RealHeight / 4)
                Else
                    .top = 0
                    .Bottom = .top + (mTexture(Tex_Projectile(MapProjectile(i).Graphic)).RealHeight)
                End If
                .Left = Anim * (mTexture(Tex_Projectile(MapProjectile(i).Graphic)).RealWidth / 12)
                .Right = .Left + (mTexture(Tex_Projectile(MapProjectile(i).Graphic)).RealWidth / 12)
            End With
            
            Select Case .Direction
                ' Up
                Case DIR_UP
                    Xo = Xo + .ProjectileOffset(DIR_UP + 1).X
                    Yo = Yo + .ProjectileOffset(DIR_UP + 1).Y

                ' Down
                Case DIR_DOWN
                    Xo = Xo + .ProjectileOffset(DIR_DOWN + 1).X
                    Yo = Yo + .ProjectileOffset(DIR_DOWN + 1).Y

                ' Left
                Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                    Xo = Xo + .ProjectileOffset(DIR_LEFT + 1).X
                    Yo = Yo + .ProjectileOffset(DIR_LEFT + 1).Y

                ' Right
                Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                    Xo = Xo + .ProjectileOffset(DIR_RIGHT + 1).X
                    Yo = Yo + .ProjectileOffset(DIR_RIGHT + 1).Y

            End Select

            ' ****** Render Projectile ******
            If .Rotate = 0 Then
                Call RenderTexture(Tex_Projectile(.Graphic), ConvertMapX(Xo), ConvertMapY(Yo), sRECT.Left, sRECT.top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.top)
            Else
                Call RenderTexture(Tex_Projectile(.Graphic), ConvertMapX(Xo), ConvertMapY(Yo), sRECT.Left, sRECT.top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.top, , , .Rotate)
            End If
            
        End If
    End With
    
    ' ****** Erase Projectile ******    Seperate Loop For Erasing
    If MapProjectile(i).Graphic Then
        ' VERIFICA TYLE_BLOCK e TYLE_RESOURCE
        For X = 0 To Map.MapData.MaxX
            For Y = 0 To Map.MapData.MaxY
                If Map.TileData.Tile(X, Y).Type = TILE_TYPE_BLOCKED Or Map.TileData.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                    If Abs(MapProjectile(i).X - (X * PIC_X)) < 20 Then
                        If Abs(MapProjectile(i).Y - (Y * PIC_Y)) < 20 Then
                            Call ClearProjectile(i)
                            Exit Sub
                        End If
                    End If
                End If
            Next Y
        Next X
        
        ' VERIFICA PLAYER NO CAMINHO
        For N = 1 To Player_HighIndex
            If N <> MyIndex Then
                If IsPlaying(N) Then
                    If Abs(MapProjectile(i).X - (GetPlayerX(N) * PIC_X)) < 20 Then
                        If Abs(MapProjectile(i).Y - (GetPlayerY(N) * PIC_Y)) < 20 Then
                            If MapProjectile(i).speed <> 6000 Then
                                Call SendTarget(N, TARGET_TYPE_PLAYER)
                                If Not MapProjectile(i).IsAoE Then
                                    Call ClearProjectile(i)
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
    
        ' VERIFICA NPC NO CAMINHO
        For N = 1 To Npc_HighIndex
            If MapNpc(N).num <> 0 Then
                If Abs(MapProjectile(i).X - (MapNpc(N).X * PIC_X)) < 20 Then
                    If Abs(MapProjectile(i).Y - (MapNpc(N).Y * PIC_Y)) < 20 Then
                        If MapProjectile(i).speed <> 6000 Then
                            Call SendTarget(N, TARGET_TYPE_NPC)
                            If Not MapProjectile(i).IsAoE Then
                                Call ClearProjectile(i)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Next
        
        ' VERIFICA SE CHEGOU AO ALVO
        If Abs(MapProjectile(i).X - MapProjectile(i).tx) < 20 Then
            If Abs(MapProjectile(i).Y - MapProjectile(i).ty) < 20 Then
                If MapProjectile(i).speed <> 6000 Then
                    Call ClearProjectile(i)
                    Exit Sub
                End If
            End If
        End If
        
        ' VERIFICAR SE É UMA TRAP E O TEMPO DE SPAWN ACABOU
        If MapProjectile(i).speed >= 5000 Then
            If Tick >= MapProjectile(i).Duration Then
                Call ClearProjectile(i)
                Exit Sub
            End If
        End If
    End If
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
        If MapNpc(MapNpcNum).AttackTimer + (attackspeed / 2) > getTime Then
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

            If .AnimTimer + 100 <= getTime Then
                .Anim = .Anim + 1

                If .Anim >= 3 Then .Anim = 0
                .AnimTimer = getTime
            End If

            Anim = .Anim
        End With

    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)

        If .AttackTimer + attackspeed < getTime Then
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
        .top = (mTexture(TextureChar(sprite)).RealHeight / 4) * SpriteTop
        .Height = mTexture(TextureChar(sprite)).RealHeight / 4
        .Left = Anim * (TextureChar(Tex_Char(sprite)).RealWidth / 4)
        .Width = (mTexture(TextureChar(sprite)).RealWidth / 4)
    End With

    ' Calculate the X
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).xOffset - ((mTexture(TextureChar(sprite)).RealWidth / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (mTexture(TextureChar(sprite)).RealHeight / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).yOffset - ((mTexture(TextureChar(sprite)).RealHeight / 4) - 32) - 4
    Else
        ' Proceed as normal
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).yOffset - 4
    End If
    
    RenderTexture Tex_Char(sprite), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.top, rec.Width, rec.Height, rec.Width, rec.Height
End Sub

Public Sub DrawShadow(ByVal sprite As Long, ByVal X As Long, ByVal Y As Long)
    If hasSpriteShadow(sprite) Then RenderTexture Tex_Shadow, ConvertMapX(X), ConvertMapY(Y), 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawTarget(ByVal X As Long, ByVal Y As Long)
    Dim Width As Long, Height As Long
    ' calculations
    Width = mTexture(Tex_Target).RealWidth / 2
    Height = mTexture(Tex_Target).RealHeight
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2) + 16
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    'EngineRenderRectangle Tex_Target, x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_Target, X, Y, 0, 0, Width, Height, Width, Height
End Sub

Public Sub DrawTargetHover()
    Dim i As Long, X As Long, Y As Long, Width As Long, Height As Long

    If diaIndex > 0 Then Exit Sub
    Width = mTexture(Tex_Target).RealWidth / 2
    Height = mTexture(Tex_Target).RealHeight

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
        .Bottom = mTexture(TextureResource(Resource_sprite)).RealHeight
        .Left = 0
        .Right = mTexture(TextureResource(Resource_sprite)).RealWidth
    End With

    ' Set base x + y, then the offset due to size
    X = (MapResource(Resource_num).X * PIC_X) - (mTexture(TextureResource(Resource_sprite)).RealWidth / 2) + 16
    Y = (MapResource(Resource_num).Y * PIC_Y) - mTexture(TextureResource(Resource_sprite)).RealHeight + 32
    Width = rec.Right - rec.Left
    Height = rec.Bottom - rec.top
    'EngineRenderRectangle Tex_Resource(Resource_sprite), ConvertMapX(x), ConvertMapY(y), 0, 0, width, height, width, height, width, height
    RenderTexture TextureResource(Resource_sprite), ConvertMapX(X), ConvertMapY(Y), 0, 0, Width, Height, Width, Height

End Sub

Public Sub DrawItem(ByVal ItemNum As Long)
    Dim Picnum As Integer, dontRender As Boolean, i As Long, tmpIndex As Long
    Picnum = Item(MapItem(ItemNum).num).pic

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
    Width = mTexture(TextureBars).RealWidth
    Height = mTexture(TextureBars).RealHeight / 4
    
    ' render npc health bars
    For i = 1 To MAX_MAP_NPCS
        NpcNum = MapNpc(i).num
        ' exists?
        If NpcNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) < Npc(NpcNum).HP Then
                ' lock to npc
                tmpX = MapNpc(i).X * PIC_X + MapNpc(i).xOffset + 16 - (Width / 2)
                tmpY = MapNpc(i).Y * PIC_Y + MapNpc(i).yOffset + 35
                
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
        GDIRenderItem frmEditor_Item.picItem, frmEditor_Item.scrlPic.Value
        GDIRenderItemPaperdoll frmEditor_Item.picPaperdoll, frmEditor_Item.scrlPaperdoll.Value
    ElseIf frmEditor_Map.visible Then
        GDIRenderTileset
    ElseIf frmEditor_NPC.visible Then
        GDIRenderChar frmEditor_NPC.picSprite, frmEditor_NPC.scrlSprite.Value
    ElseIf frmEditor_Resource.visible Then
        GDIRenderResource frmEditor_Resource.picNormalPic, frmEditor_Resource.scrlNormalPic.Value
        GDIRenderResource frmEditor_Resource.picExhaustedPic, frmEditor_Resource.scrlExhaustedPic.Value
    ElseIf frmEditor_Spell.visible Then
        GDIRenderSpell frmEditor_Spell.picSprite, frmEditor_Spell.scrlIcon.Value
        GDIRenderSpellProjectile frmEditor_Spell.picProjectile, frmEditor_Spell.scrlProjectilePic.Value
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

    ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
    For y = TileView.top To TileView.bottom + 5

        ' Resources
        If CountResource > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).Y = Y Then
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
                    If Player(i).Y = Y Then
                        Call DrawPlayer(i)
                    End If
                End If
            Next

            ' Npcs
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).Y = Y Then
                    Call DrawNpc(i)
                End If
            Next
        End If
        
        ' draw projectiles on map
        For i = 1 To LastProjectile
            If MapProjectile(i).Owner > 0 Then
                If MapProjectile(i).speed < 5000 Then
                    If Int(MapProjectile(i).Y / PIC_Y) = Y Then
                        Call DrawProjectile(i)
                    End If
                End If
            End If
        Next
        
        ' draw traps on map
        For i = 1 To LastProjectile
            If MapProjectile(i).Owner > 0 Then
                If MapProjectile(i).speed >= 5000 Then
                    If Int(MapProjectile(i).Y / PIC_Y) = Y Then
                        Call DrawProjectile(i)
                    End If
                End If
            End If
        Next
    Next Y
    
    

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
