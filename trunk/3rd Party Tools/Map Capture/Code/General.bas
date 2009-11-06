Attribute VB_Name = "General"
Option Explicit

'Holds data about where a png can be found,
'How big it is and animation info
Public Type GrhData
    SX As Integer
    SY As Integer
    FileNum As Long
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Byte
    Frames() As Long
    Speed As Single
End Type

'Holds a position on a 2d grid
Public Type Position
    x As Long
    y As Long
End Type

'Holds a position on a 2d grid in floating variables (singles)
Public Type FloatPos
    x As Single
    y As Single
End Type

'Holds a world position
Private Type WorldPos
    x As Byte
    y As Byte
End Type

'Points to a grhData and keeps animation info
Public Type Grh
    GrhIndex As Long
    LastCount As Long
    FrameCounter As Single
    Started As Byte
End Type

'Texture information
Public Type TexInfo
    x As Long
    y As Long
End Type

'Holds info about each tile position
Public Type MapBlock
    Graphic(1 To 6) As Grh
    Light(1 To 24) As Long
    Shadow(1 To 6) As Byte
End Type

'Hold info about each map
Public Type MapInfo
    Name As String
    Width As Byte
    Height As Byte
End Type

Public MapNum2 As Long

'Describes the return from a texture init
Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

Public Enum LogType
    General = 0
    CodeTracker = 1
    PacketIn = 2
    PacketOut = 3
    CriticalError = 4
    InvalidPacketData = 5
End Enum

'Describes a transformable lit vertex
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
Public Type TLVERTEX
    x As Single
    y As Single
    Z As Single
    Rhw As Single
    Color As Long
    tU As Single
    tV As Single
End Type

'The size of a FVF vertex
Public Const FVF_Size As Long = 28

'DirectX 8 Objects
Public DX As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8
Public D3DDevice As Direct3DDevice8

'Everything else
Public NumGrhFiles As Long
Public NumGrhs As Long
Public SurfaceDB() As Direct3DTexture8          'The list of all the textures
Public SurfaceLoaded() As Boolean
Public LastTexture As Long                      'The last texture used
Public D3DWindow As D3DPRESENT_PARAMETERS       'Describes the viewport and used to restore when in fullscreen
Public UsedCreateFlags As CONST_D3DCREATEFLAGS  'The flags we used to create the device when it first succeeded
Public DispMode As D3DDISPLAYMODE               'Describes the display mode
Public GrhData() As GrhData             'Holds data for the graphic structure
Public SurfaceSize() As TexInfo         'Holds the size of the surfaces for SurfaceDB()
Public MapData() As MapBlock            'Holds map data for current map
Public MapInfo As MapInfo               'Holds map info for current map

Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Sub SaveBackBuffer(ByVal FileName As String)
Dim RECT As RECT
Dim PAL As PALETTEENTRY
    PAL.blue = 255
    PAL.green = 255
    PAL.red = 255
    RECT.Right = frmMain.ScreenPic.Width / Screen.TwipsPerPixelY
    RECT.bottom = frmMain.ScreenPic.Height / Screen.TwipsPerPixelY
    D3DX.SaveSurfaceToFile FileName, D3DXIFF_BMP, D3DDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO), PAL, RECT
End Sub

Sub Main()
    InitManifest
    NumGrhFiles = CLng(Var_Get(GetRootPath & "Data\Grh.ini", "INIT", "NumGrhFiles"))
    ReDim SurfaceDB(1 To NumGrhFiles)
    ReDim SurfaceLoaded(1 To NumGrhFiles)
    ReDim SurfaceSize(1 To NumGrhFiles)
    NumGrhs = CLng(Var_Get(GetRootPath & "Data\Grh.ini", "INIT", "NumGrhs"))
    ReDim GrhData(1 To NumGrhs) As GrhData
    Game_Map_Switch 1
    Engine_Init_GrhData
    InitTileEngine
    frmMain.Show
End Sub

Sub InitTileEngine()
    Load frmMain
    frmMain.Show
    DoEvents
    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate()
    Set D3DX = New D3DX8
    If Not Engine_Init_D3DDevice(D3DCREATE_PUREDEVICE) Then
        If Not Engine_Init_D3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
            If Not Engine_Init_D3DDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                If Not Engine_Init_D3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                    MsgBox "Could not init D3DDevice. Exiting..."
                    UnloadProject
                End If
            End If
        End If
    End If
    Engine_Init_RenderStates
End Sub

Public Sub UnloadProject()
Dim i As Long
    If Not DX Is Nothing Then Set DX = Nothing
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    If Not D3DX Is Nothing Then Set D3DX = Nothing
    For i = 1 To LastTexture
        If Not SurfaceDB(i) Is Nothing Then Set SurfaceDB(i) = Nothing
    Next i
    Unload frmMain
    End
End Sub

Private Function Engine_Init_D3DDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS) As Boolean
    On Error GoTo ErrOut
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3DWindow.Windowed = 1
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
    D3DWindow.BackBufferFormat = DispMode.Format
    D3DWindow.BackBufferWidth = 800 ' MapInfo.Width * 32
    D3DWindow.BackBufferHeight = 600 ' MapInfo.Height * 32
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.ScreenPic.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
    UsedCreateFlags = D3DCREATEFLAGS
    Engine_Init_D3DDevice = True
    frmMain.Show
    frmMain.Refresh
    DoEvents
Exit Function
ErrOut:
    Set D3DDevice = Nothing
    Engine_Init_D3DDevice = False
End Function

Private Sub Engine_Init_RenderStates()
    With D3DDevice
        D3DDevice.SetVertexShader FVF
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    End With
End Sub

Public Sub DrawScreen(ByVal StartX As Long, ByVal StartY As Long)
Dim i As Long
Dim x As Long
Dim y As Long

    D3DDevice.BeginScene
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    
        'Draw
        For i = 1 To 6
            For y = 1 To MapInfo.Height
                For x = 1 To MapInfo.Width
                    With MapData(x, y)
                        If .Graphic(i).GrhIndex > 0 Then
                            Engine_Render_Grh .Graphic(i), ((x - StartX) - 1) * 32, ((y - StartY) - 1) * 32, .Light(((i - 1) * 4) + 1), .Light(((i - 1) * 4) + 2), .Light(((i - 1) * 4) + 3), .Light(((i - 1) * 4) + 4)
                        End If
                    End With
                Next x
            Next y
        Next i
        
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Sub Engine_Render_Grh(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, Optional ByVal Light1 As Long = -1, Optional ByVal Light2 As Long = -1, Optional ByVal Light3 As Long = -1, Optional ByVal Light4 As Long = -1, Optional ByVal Shadow As Byte = 0, Optional ByVal Angle As Single = 0)
Dim CurrGrhIndex As Long
Dim FileNum As Integer
    If Grh.GrhIndex < 1 Then Exit Sub
    If Grh.GrhIndex > NumGrhs Then Exit Sub
    If GrhData(Grh.GrhIndex).NumFrames < 1 Then Exit Sub
    CurrGrhIndex = GrhData(Grh.GrhIndex).Frames(1)
    FileNum = GrhData(GrhData(Grh.GrhIndex).Frames(1)).FileNum
    Engine_Render_Rectangle x, y, GrhData(CurrGrhIndex).pixelWidth, GrhData(CurrGrhIndex).pixelHeight, GrhData(CurrGrhIndex).SX, GrhData(CurrGrhIndex).SY, GrhData(CurrGrhIndex).pixelWidth, GrhData(CurrGrhIndex).pixelHeight, , , Angle, FileNum, Light1, Light2, Light3, Light4, Shadow
End Sub

Public Sub Log(one, two)
'o.O
End Sub

Sub Game_Map_Switch(Map As Integer)

'*****************************************************************
'Loads and switches to a new map
'*****************************************************************
Dim LargestTileSize As Long
Dim MapBuf As DataBuffer
Dim GetParticleCount As Integer
Dim GetEffectNum As Byte
Dim GetDirection As Integer
Dim GetGfx As Byte
Dim GetX As Integer
Dim GetY As Integer
Dim ByFlags As Long
Dim MapNum As Byte
Dim i As Integer
Dim y As Byte
Dim x As Byte
Dim b() As Byte
Dim TempInt As Integer
    MapNum2 = Map
    frmMain.Caption = "Map Capture Tool - Map " & Map
    MapNum = FreeFile
    Open GetRootPath & "Maps\" & Map & ".map" For Binary As #MapNum
        Seek #MapNum, 1
        ReDim b(0 To LOF(MapNum) - 1)
        Get #MapNum, , b()
    Close #MapNum
    Set MapBuf = New DataBuffer
    MapBuf.Set_Buffer b()
    Erase b()
    TempInt = MapBuf.Get_Integer
    MapInfo.Width = MapBuf.Get_Byte
    MapInfo.Height = MapBuf.Get_Byte
    ReDim MapData(1 To MapInfo.Width, 1 To MapInfo.Height) As MapBlock
    For y = 1 To MapInfo.Height
        For x = 1 To MapInfo.Width
            For i = 1 To 6
                MapData(x, y).Graphic(i).GrhIndex = 0
            Next i
            ByFlags = MapBuf.Get_Long
            If ByFlags And 1 Then i = MapBuf.Get_Byte
            If ByFlags And 2 Then MapData(x, y).Graphic(1).GrhIndex = MapBuf.Get_Long
            If ByFlags And 4 Then MapData(x, y).Graphic(2).GrhIndex = MapBuf.Get_Long
            If ByFlags And 8 Then MapData(x, y).Graphic(3).GrhIndex = MapBuf.Get_Long
            If ByFlags And 16 Then MapData(x, y).Graphic(4).GrhIndex = MapBuf.Get_Long
            If ByFlags And 32 Then MapData(x, y).Graphic(5).GrhIndex = MapBuf.Get_Long
            If ByFlags And 64 Then MapData(x, y).Graphic(6).GrhIndex = MapBuf.Get_Long
            For i = 1 To 24
                MapData(x, y).Light(i) = -1
            Next i
            If ByFlags And 128 Then
                For i = 1 To 4
                    MapData(x, y).Light(i) = MapBuf.Get_Long
                Next i
            End If
            If ByFlags And 256 Then
                For i = 5 To 8
                    MapData(x, y).Light(i) = MapBuf.Get_Long
                Next i
            End If
            If ByFlags And 512 Then
                For i = 9 To 12
                    MapData(x, y).Light(i) = MapBuf.Get_Long
                Next i
            End If
            If ByFlags And 1024 Then
                For i = 13 To 16
                    MapData(x, y).Light(i) = MapBuf.Get_Long
                Next i
            End If
            If ByFlags And 2048 Then
                For i = 17 To 20
                    MapData(x, y).Light(i) = MapBuf.Get_Long
                Next i
            End If
            If ByFlags And 4096 Then
                For i = 21 To 24
                    MapData(x, y).Light(i) = MapBuf.Get_Long
                Next i
            End If
            If ByFlags And 16384 Then MapData(x, y).Shadow(1) = 1 Else MapData(x, y).Shadow(1) = 0
            If ByFlags And 32768 Then MapData(x, y).Shadow(2) = 1 Else MapData(x, y).Shadow(2) = 0
            If ByFlags And 65536 Then MapData(x, y).Shadow(3) = 1 Else MapData(x, y).Shadow(3) = 0
            If ByFlags And 131072 Then MapData(x, y).Shadow(4) = 1 Else MapData(x, y).Shadow(4) = 0
            If ByFlags And 262144 Then MapData(x, y).Shadow(5) = 1 Else MapData(x, y).Shadow(5) = 0
            If ByFlags And 524288 Then MapData(x, y).Shadow(6) = 1 Else MapData(x, y).Shadow(6) = 0
            If ByFlags And 1048576 Then i = MapBuf.Get_Integer
            If ByFlags And 4194304 Then i = MapBuf.Get_Integer
        Next x
    Next y
    Set MapBuf = Nothing
End Sub

Sub Engine_Render_Rectangle(ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal SrcX As Single, ByVal SrcY As Single, ByVal SrcWidth As Single, ByVal SrcHeight As Single, Optional ByVal SrcBitmapWidth As Long = -1, Optional ByVal SrcBitmapHeight As Long = -1, Optional ByVal Degrees As Single = 0, Optional ByVal TextureNum As Long, Optional ByVal Color0 As Long = -1, Optional ByVal Color1 As Long = -1, Optional ByVal Color2 As Long = -1, Optional ByVal Color3 As Long = -1, Optional ByVal Shadow As Byte = 0)
Dim VertexArray(0 To 3) As TLVERTEX
Dim RadAngle As Single
Dim CenterX As Single
Dim CenterY As Single
Dim Index As Integer
Dim NewX As Single
Dim NewY As Single
Dim SinRad As Single
Dim CosRad As Single
Dim ShadowAdd As Single
Dim l As Single
    Engine_ReadyTexture TextureNum
    If SrcBitmapWidth = -1 Then SrcBitmapWidth = SurfaceSize(TextureNum).x
    If SrcBitmapHeight = -1 Then SrcBitmapHeight = SurfaceSize(TextureNum).y
    VertexArray(0).Rhw = 1
    VertexArray(1).Rhw = 1
    VertexArray(2).Rhw = 1
    VertexArray(3).Rhw = 1
    VertexArray(0).Color = Color0
    VertexArray(1).Color = Color1
    VertexArray(2).Color = Color2
    VertexArray(3).Color = Color3
    If Shadow Then
        VertexArray(0).x = x + (Width * 0.5)
        VertexArray(0).y = y - (Height * 0.5)
        VertexArray(0).tU = (SrcX / SrcBitmapWidth)
        VertexArray(0).tV = (SrcY / SrcBitmapHeight)
        VertexArray(1).x = VertexArray(0).x + Width
        VertexArray(1).tU = ((SrcX + Width) / SrcBitmapWidth)
        VertexArray(2).x = x
        VertexArray(2).tU = (SrcX / SrcBitmapWidth)
        VertexArray(3).x = x + Width
        VertexArray(3).tU = (SrcX + SrcWidth + ShadowAdd) / SrcBitmapWidth
    Else
        ShadowAdd = 1
        VertexArray(0).x = x
        VertexArray(0).tU = (SrcX / SrcBitmapWidth)
        VertexArray(0).y = y
        VertexArray(0).tV = (SrcY / SrcBitmapHeight)
        VertexArray(1).x = x + Width
        VertexArray(1).tU = (SrcX + SrcWidth + ShadowAdd) / SrcBitmapWidth
        VertexArray(2).x = VertexArray(0).x
        VertexArray(3).x = VertexArray(1).x
    End If
    VertexArray(2).y = y + Height
    VertexArray(2).tV = (SrcY + SrcHeight + ShadowAdd) / SrcBitmapHeight
    VertexArray(1).y = VertexArray(0).y
    VertexArray(1).tV = VertexArray(0).tV
    VertexArray(2).tU = VertexArray(0).tU
    VertexArray(3).y = VertexArray(2).y
    VertexArray(3).tU = VertexArray(1).tU
    VertexArray(3).tV = VertexArray(2).tV
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
End Sub

Private Sub Engine_ReadyTexture(ByVal TextureNum As Long)
    If TextureNum <= 0 Then Exit Sub
    If Not SurfaceLoaded(TextureNum) Then
        Engine_Init_Texture TextureNum
    End If
    If LastTexture <> TextureNum Then
        D3DDevice.SetTexture 0, SurfaceDB(TextureNum)
        LastTexture = TextureNum
    End If
End Sub

Sub Engine_Init_Texture(ByVal TextureNum As Integer)
Dim UseTextureFormat As CONST_D3DFORMAT
Dim TexInfo As D3DXIMAGE_INFO_A
Dim FilePath As String
    If TextureNum < 1 Then Exit Sub
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
    SurfaceLoaded(TextureNum) = True
    FilePath = GetRootPath & "Grh\" & TextureNum & ".png"
    Set SurfaceDB(TextureNum) = D3DX.CreateTextureFromFileEx(D3DDevice, FilePath, D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, TexInfo, ByVal 0)
    SurfaceSize(TextureNum).x = TexInfo.Width
    SurfaceSize(TextureNum).y = TexInfo.Height
End Sub

Function GetRootPath() As String
Dim s() As String
Dim i As Long
    s = Split(App.Path, "\")
    For i = 0 To UBound(s) - 2
        GetRootPath = GetRootPath & s(i) & "\"
    Next i
End Function

Sub Engine_Init_GrhData()
Dim FileNum As Byte
Dim Grh As Long
Dim Frame As Long
    NumGrhs = CLng(Var_Get(GetRootPath & "Data\Grh.ini", "INIT", "NumGrhs"))
    ReDim GrhData(1 To NumGrhs) As GrhData
    FileNum = FreeFile
    Open GetRootPath & "Data\Grh.dat" For Binary As #FileNum
    Seek #FileNum, 1
    Get #FileNum, , Grh
    Do Until Grh <= 0
        Get #FileNum, , GrhData(Grh).NumFrames
        If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
        If GrhData(Grh).NumFrames > 1 Then
            ReDim GrhData(Grh).Frames(1 To GrhData(Grh).NumFrames)
            For Frame = 1 To GrhData(Grh).NumFrames
                Get #FileNum, , GrhData(Grh).Frames(Frame)
                If GrhData(Grh).Frames(Frame) <= 0 Then
                    GoTo ErrorHandler
                End If
            Next Frame
            Get #FileNum, , GrhData(Grh).Speed
            GrhData(Grh).Speed = 1
            If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
            GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
            GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
            GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
            If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
            GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
            If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
        Else
            ReDim GrhData(Grh).Frames(1 To 1)
            Get #FileNum, , GrhData(Grh).FileNum
            If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
            Get #FileNum, , GrhData(Grh).SX
            If GrhData(Grh).SX < 0 Then GoTo ErrorHandler
            Get #FileNum, , GrhData(Grh).SY
            If GrhData(Grh).SY < 0 Then GoTo ErrorHandler
            Get #FileNum, , GrhData(Grh).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
            Get #FileNum, , GrhData(Grh).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
            GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / 32
            GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / 32
            GrhData(Grh).Frames(1) = Grh
        End If
        Get #FileNum, , Grh
    Loop
    Close #FileNum
Exit Sub
ErrorHandler:
    Close #FileNum
    MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh
    UnloadProject
End Sub

Function Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
    Var_Get = Space$(1000)
    getprivateprofilestring Main, Var, vbNullString, Var_Get, 1000, File
    Var_Get = RTrim$(Var_Get)
    If LenB(Var_Get) <> 0 Then Var_Get = Left$(Var_Get, Len(Var_Get) - 1)
End Function
