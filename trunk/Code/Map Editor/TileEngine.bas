Attribute VB_Name = "TileEngine"
Option Explicit

Public Const ShadowColor As Long = 1677721600  'ARGB 100/0/0/0

Public ParticleOffsetX As Long
Public ParticleOffsetY As Long
Public LastOffsetX As Integer           'The last offset values stored, used to get the offset difference
Public LastOffsetY As Integer           ' so the particle engine can adjust weather particles accordingly

Public lngTextHeight As Long

'********** CONSTANTS ***********
'Heading constants
Public Const NORTH = 1
Public Const EAST = 2
Public Const SOUTH = 3
Public Const WEST = 4
Public Const NORTHEAST = 5
Public Const SOUTHEAST = 6
Public Const SOUTHWEST = 7
Public Const NORTHWEST = 8

'********** WEATHER ***********
Public Type LightType
    Light(1 To 24) As Long
End Type
Public SaveLightBuffer() As LightType
Public WeatherEffectIndex As Integer    'Index returned by the weather effect initialization
Public WeatherDoLightning As Byte   'Are we using lightning? >1 = Yes, 0 = No
Public WeatherFogX1 As Single       'Fog 1 position
Public WeatherFogY1 As Single       'Fog 1 position
Public WeatherFogX2 As Single       'Fog 2 position
Public WeatherFogY2 As Single       'Fog 2 position
Public WeatherDoFog As Byte         'Are we using fog? >1 = Yes, 0 = No
Public WeatherFogCount As Byte      'How many fog effects there are
Public LightningTimer As Single     'How long until our next lightning bolt strikes
Public FlashTimer As Single         'How long until the flash goes away (being > 0 states flash is happening)
Public LightningX As Integer        'Position of the lightning (top-left corner)
Public LightningY As Integer
Public LastWeather As Byte

'********** TYPES ***********

'Holds a position on a 2d grid
Public Type Position
    X As Long
    Y As Long
End Type

'Holds a position on a 2d grid in floating variables (singles)
Public Type FloatPos
    X As Single
    Y As Single
End Type

'Holds a world position
Public Type WorldPos
    X As Byte
    Y As Byte
End Type

'Holds a world position
Private Type WorldPosEX
    Map As Integer
    X As Byte
    Y As Byte
End Type

'Holds data about where a png can be found,
'How big it is and animation info
Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Long
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Byte
    Frames() As Long
    Speed As Single
End Type

'Points to a grhData and keeps animation info
Public Type Grh
    GrhIndex As Long
    LastCount As Long
    FrameCounter As Single
    SpeedCounter As Byte
    Started As Byte
End Type

'Bodies list
Public Type BodyData
    Walk(1 To 8) As Grh
    Attack(1 To 8) As Grh
    HeadOffset As Position
End Type

'Wings list
Public Type WingData
    Walk(1 To 8) As Grh
    Attack(1 To 8) As Grh
End Type

'Weapons list
Public Type WeaponData
    Walk(1 To 8) As Grh
    Attack(1 To 8) As Grh
End Type

'Heads list
Public Type HeadData
    Head(1 To 8) As Grh
    Blink(1 To 8) As Grh
    AgrHead(1 To 8) As Grh
    AgrBlink(1 To 8) As Grh
End Type

'Hair list
Public Type HairData
    Hair(1 To 8) As Grh
End Type

'Hold info about the character's status
Public Type CharStatus
    Cursed As Byte
    WarCursed As Byte
    Blessed As Byte
    Protected As Byte
    Strengthened As Byte
    IronSkinned As Byte
    Exhausted As Byte
End Type

'Hold info about a character
Public Type Char
    Active As Byte
    Heading As Byte
    HeadHeading As Byte
    Pos As Position             'Tile position on the map
    RealPos As Position         'Position on the game screen
    Body As BodyData
    Head As HeadData
    Weapon As WeaponData
    Hair As HairData
    Wings As WingData
    Moving As Byte
    Aggressive As Byte
    MoveOffset As FloatPos
    BlinkTimer As Single
    ScrollDirectionX As Integer
    ScrollDirectionY As Integer
    Name As String
    ActionIndex As Byte
    HealthPercent As Byte
    CharStatus As CharStatus
    NPCNumber As Integer
    Emoticon As Grh
    EmoFade As Single
    EmoDir As Byte      'Direction the fading is going - 0 = Stopped, 1 = Up, 2 = Down
End Type

'Holds data for a character - used for saving/loading NPCs (not displaying the char)
Type CharShort
    CharIndex As Integer    'Character's index
    Hair As Integer         'Hair index
    Head As Integer         'Head index
    Body As Integer         'Body index
    Weapon As Integer       'Weapon index
    Heading As Byte         'Current direction facing
    HeadHeading As Byte     'Direction char's head is facing
    Desc As String          'Description
End Type

'Holds info about a object
Public Type OBJ
    ObjIndex As Integer     'Index of the object
    Amount As Integer       'Amount of the object
End Type

'Holds info about each tile position
Type MapBlock
    Blocked As Byte             'If the tile is blocked
    BlockedAttack As Byte       'If you can not attack through the tile
    Graphic(1 To 6) As Grh      'Index of the 4 graphic layers
    Light(1 To 24) As Long      'Holds the light values - retrieve with Index = Light * Layer
    UserIndex As Integer        'Index of the user on the tile
    NPCIndex As Integer         'Index of the NPC on the tile
    ObjInfo As OBJ              'Information of the object on the tile
    TileExit As WorldPosEX      'Warp location when user touches the tile
    Mailbox As Byte             'If there is a mailbox on the tile
    Sign As Integer             'The sign value, if any
    Shadow(1 To 6) As Byte      'If the surface shows a shadow
    Sfx As Integer              'Index of the .wav file to be looped
End Type

'Hold info about each map
Public Type MapInfo
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    Weather As Byte
    Music As Byte
    Width As Byte
    Height As Byte
End Type

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

'Describes a layer bound to tile position but not in the map array (to save memory)
Private Type FloatSurface
    Pos As WorldPos
    Grh As Grh
End Type

'********** Public VARS ***********

'Where the map borders are.. Set during load
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'User status vars
Public CurMap As Integer            'Current map loaded
Public UserMoving As Boolean
Public UserPos As Position          'Holds current user pos
Public AddtoUserPos As Position     'For moving user
Public EngineRun As Boolean
Public FPS As Long
Private FramesPerSecCounter As Long
Private FPS_Last_Check As Long

'Main view size size in tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

'How many tiles the engine "looks ahead" when drawing the screen
Public TileBufferSize As Integer

'Handle to where all the drawing is going to take place
Private DisplayFormhWnd As Long

'Tile size in pixels
Private TilePixelHeight As Integer
Private TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Private ScrollPixelsPerFrameX As Integer
Private ScrollPixelsPerFrameY As Integer

'Totals
Private NumBodies As Integer    'Number of bodies
Private NumWings As Integer     'Number of wings
Private NumHeads As Integer     'Number of heads
Private NumHairs As Integer     'Number of hairs
Private NumWeapons As Integer   'Number of weapons
Private NumGrhs As Long         'Number of grhs
Public NumMaps As Integer       'Number of maps
Public NumGrhFiles As Integer   'Number of pngs
Public LastChar As Integer      'Last character
Public LastObj As Integer       'Last object

'Screen positioning
Public minY As Integer          'Start Y pos on current screen + tilebuffer
Public maxY As Integer          'End Y pos on current screen
Public minX As Integer          'Start X pos on current screen
Public maxX As Integer          'End X pos on current screen
Public minXOffset As Integer
Public minYOffset As Integer
Public ScreenMinY As Integer    'Start Y pos on current screen
Public ScreenMaxY As Integer    'End Y pos on current screen
Public ScreenMinX As Integer    'Start X pos on current screen
Public ScreenMaxX As Integer    'End X pos on current screen
Public LastTileX As Integer
Public LastTileY As Integer

'********** Direct X ***********
Public Const SurfaceTimerMax As Long = 600000    'How long a texture stays in memory unused (miliseconds)
Public SurfaceDB() As Direct3DTexture8          'The list of all the textures
Public SurfaceTimer() As Long                   'How long until the surface unloads
Public LastTexture As Long                      'The last texture used
Public D3DWindow As D3DPRESENT_PARAMETERS       'Describes the viewport and used to restore when in fullscreen
Public UsedCreateFlags As CONST_D3DCREATEFLAGS  'The flags we used to create the device when it first succeeded

'Texture for particle effects - this is handled differently then the rest of the graphics
Public ParticleTexture(1 To 12) As Direct3DTexture8

'DirectX 8 Objects
Private DX As DirectX8
Private D3D As Direct3D8
Public D3DX As D3DX8
Public D3DDevice As Direct3DDevice8
Private MainFont As D3DXFont
Private MainFontDesc As IFont

'Describes a transformable lit vertex
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE

Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    Rhw As Single
    Color As Long
    Tu As Single
    Tv As Single
End Type

'The size of a FVF vertex
Public Const FVF_Size As Long = 28

'Used to hold the graphic layers in a quick-to-draw format
Public Type Tile
    TileX As Byte
    TileY As Byte
    PixelPosX As Integer
    PixelPosY As Integer
End Type
Public Type TileLayer
    Tile() As Tile
    NumTiles As Integer
End Type
Public TileLayer(1 To 6) As TileLayer

'Holds the information on the "info layer" (map editor only)
Public Type InfoTileGrh
    PixelPosX As Integer
    PixelPosY As Integer
    Grh As Grh
End Type
Public Type InfoTile
    NumGrhs As Byte
    Grh() As InfoTileGrh
End Type
Public Type InfoLayer
    Tile() As InfoTile
    NumTiles As Integer
End Type
Public InfoLayer As InfoLayer

'********** Public ARRAYS ***********
Public GrhData() As GrhData         'Holds data for the graphic structure
Public SurfaceSize() As Point       'Holds the size of the surfaces for SurfaceDB()
Public BodyData() As BodyData       'Holds data about body structure
Public HeadData() As HeadData       'Holds data about head structure
Public HairData() As HairData       'Holds data about hair structure
Public WeaponData() As WeaponData   'Holds data about weapon structure
Public WingData() As WingData       'Holds data about wing structure
Public MapData() As MapBlock        'Holds map data for current map
Public MapInfo As MapInfo           'Holds map info for current map
Public CharList() As Char           'Holds info about all characters on the map
Public OBJList() As FloatSurface    'Holds info about all objects on the map
Public PreviewGrhList() As Grh      'Holds the list of Grhs for the Grh Preview screen

'Preview of the graphic to be set
Public PreviewMapGrh As Grh

'Maximum amount of 64x64 tiles that can fit into the screen
Public tsWidth As Long
Public tsHeight As Long

'The first number that we are starting from for rendering
Public tsStart As Integer

'If this is the first time drawing the list - is reset when the list is updated or screen was set from invisible to visible
Public tsDrawAll As Byte

'Size of the tiles to be displayed (anything larger is cut off)
Public tsTileWidth As Integer
Public tsTileHeight As Integer

'FPS
Public EndTime As Long
Public ElapsedTime As Single
Public TickPerFrame As Single
Public TimerMultiplier As Single
Public EngineBaseSpeed As Single
Public OffsetCounterX As Single
Public OffsetCounterY As Single

'Point API
Public Type POINTAPI
    X As Long
    Y As Long
End Type

'Mini-map tiles
Public NumMiniMapTiles As Integer   'UBound of the MiniMapTile array
Public Type MiniMapTile
    X As Byte
    Y As Byte
    Color As Long
    Caption As String
End Type
Public MiniMapTile() As MiniMapTile 'Color of each tile and their position
Public ShowMiniMap As Byte

Private LastThingy As Long   'Yes, I named a variable "thingy", wanna fight about it!? >:|

'********** OUTSIDE FUNCTIONS ***********
Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

Private Sub Engine_AddItem2Array1D(ByRef VarArray As Variant, ByVal VarValue As Variant)

'************************************************************
'Adds a variant to one-dimensional array
'************************************************************

Dim i  As Long
Dim iVarType As Integer

    iVarType = VarType(VarArray) - 8192

    i = UBound(VarArray)
    Select Case iVarType
    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
        If VarArray(0) = 0 Then
            i = 0
        Else
            i = i + 1
        End If
    Case vbDate
        If VarArray(0) = "00:00:00" Then
            i = 0
        Else
            i = i + 1
        End If
    Case vbString
        If VarArray(0) = vbNullString Then
            i = 0
        Else
            i = i + 1
        End If
    Case vbBoolean
        If VarArray(0) = False Then
            i = 0
        Else
            i = i + 1
        End If
    Case Else
    End Select
    ReDim Preserve VarArray(i)
    VarArray(i) = VarValue

End Sub

Sub Engine_Char_Erase(ByVal CharIndex As Integer)

'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************

    'Check for valid position
    If CharList(CharIndex).Pos.X < 1 Then Exit Sub
    If CharList(CharIndex).Pos.X > MapInfo.Width Then Exit Sub
    If CharList(CharIndex).Pos.Y < 1 Then Exit Sub
    If CharList(CharIndex).Pos.Y > MapInfo.Height Then Exit Sub

    'Make inactive
    CharList(CharIndex).Active = 0
    
    'Erase from map
    MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).NPCIndex = 0

    'Update LastChar
    If CharIndex = LastChar Then
        Do Until CharList(LastChar).Active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then
                Exit Do
            Else
                ReDim Preserve CharList(1 To LastChar)
            End If
        Loop
    End If

End Sub

Sub Engine_Char_Make(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Name As String, ByVal Weapon As Integer, ByVal Hair As Integer, ByVal NPCNumber As Integer)

'*****************************************************************
'Makes a new character and puts it on the map
'*****************************************************************

Dim EmptyChar As Char

'Update LastChar

    If CharIndex > LastChar Then
        LastChar = CharIndex
        ReDim Preserve CharList(1 To LastChar)
    End If

    'Clear the character
    CharList(CharIndex) = EmptyChar

    'Set the apperances
    CharList(CharIndex).Body = BodyData(Body)
    CharList(CharIndex).Head = HeadData(Head)
    CharList(CharIndex).Hair = HairData(Hair)
    CharList(CharIndex).Weapon = WeaponData(Weapon)
    CharList(CharIndex).Heading = Heading
    CharList(CharIndex).HeadHeading = Heading
    CharList(CharIndex).HealthPercent = 100

    'Reset moving stats
    CharList(CharIndex).Moving = 0
    CharList(CharIndex).MoveOffset.X = 0
    CharList(CharIndex).MoveOffset.Y = 0

    'Update position
    CharList(CharIndex).Pos.X = X
    CharList(CharIndex).Pos.Y = Y
    MapData(X, Y).NPCIndex = CharIndex

    'Make active
    CharList(CharIndex).Active = 1
    CharList(CharIndex).Name = Name
    CharList(CharIndex).NPCNumber = NPCNumber

    'Set action index
    CharList(CharIndex).ActionIndex = 0

End Sub

Sub Engine_Char_Move_ByHead(ByVal CharIndex As Integer, ByVal nHeading As Byte)

'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************

Dim AddX As Integer
Dim AddY As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

'Check for a valid CharIndex

    If CharIndex <= 0 Then Exit Sub

    X = CharList(CharIndex).Pos.X
    Y = CharList(CharIndex).Pos.Y

    'Figure out which way to move
    Select Case nHeading
    Case NORTH
        AddY = -1
    Case EAST
        AddX = 1
    Case SOUTH
        AddY = 1
    Case WEST
        AddX = -1
    Case NORTHEAST
        AddY = -1
        AddX = 1
    Case SOUTHEAST
        AddY = 1
        AddX = 1
    Case SOUTHWEST
        AddY = 1
        AddX = -1
    Case NORTHWEST
        AddY = -1
        AddX = -1
    End Select

    'Update the character position and settings
    nX = X + AddX
    nY = Y + AddY
    CharList(CharIndex).Pos.X = nX
    CharList(CharIndex).Pos.Y = nY
    CharList(CharIndex).MoveOffset.X = -(TilePixelWidth * AddX)
    CharList(CharIndex).MoveOffset.Y = -(TilePixelHeight * AddY)
    CharList(CharIndex).Moving = 1
    CharList(CharIndex).Heading = nHeading
    CharList(CharIndex).HeadHeading = nHeading
    CharList(CharIndex).ScrollDirectionX = AddX
    CharList(CharIndex).ScrollDirectionY = AddY
    CharList(CharIndex).ActionIndex = 1

End Sub

Sub Engine_Char_Move_ByPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)

'*****************************************************************
'Starts the movement of a character to nX,nY
'*****************************************************************

Dim X As Integer
Dim Y As Integer
Dim AddX As Integer
Dim AddY As Integer
Dim nHeading As Byte

    X = CharList(CharIndex).Pos.X
    Y = CharList(CharIndex).Pos.Y
    AddX = nX - X
    AddY = nY - Y

    'Figure out the direction the character is going
    If Sgn(AddX) = 1 Then nHeading = EAST
    If Sgn(AddX) = -1 Then nHeading = WEST
    If Sgn(AddY) = -1 Then nHeading = NORTH
    If Sgn(AddY) = 1 Then nHeading = SOUTH
    If Sgn(AddX) = 1 And Sgn(AddY) = -1 Then
        nHeading = NORTHEAST
    End If
    If Sgn(AddX) = 1 And Sgn(AddY) = 1 Then
        nHeading = SOUTHEAST
    End If
    If Sgn(AddX) = -1 And Sgn(AddY) = 1 Then
        nHeading = SOUTHWEST
    End If
    If Sgn(AddX) = -1 And Sgn(AddY) = -1 Then
        nHeading = NORTHWEST
    End If

    'Update the character position and settings
    CharList(CharIndex).Pos.X = nX
    CharList(CharIndex).Pos.Y = nY
    CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * AddX)
    CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * AddY)
    CharList(CharIndex).Moving = 1
    CharList(CharIndex).Heading = nHeading
    CharList(CharIndex).HeadHeading = nHeading
    CharList(CharIndex).ScrollDirectionX = Sgn(AddX)
    CharList(CharIndex).ScrollDirectionY = Sgn(AddY)
    CharList(CharIndex).ActionIndex = 1

End Sub

Sub Engine_ClearMapArray()

'*****************************************************************
'Clears all layers
'*****************************************************************

Dim i As Integer
Dim Y As Byte
Dim X As Byte

    For Y = 1 To MapInfo.Height
        For X = 1 To MapInfo.Width

            'Change blockes status
            MapData(X, Y).Blocked = 0

            'Erase layer 1 and 4
            MapData(X, Y).Graphic(1).GrhIndex = 0
            MapData(X, Y).Graphic(2).GrhIndex = 0
            MapData(X, Y).Graphic(3).GrhIndex = 0
            MapData(X, Y).Graphic(4).GrhIndex = 0

        Next X
    Next Y

    'Erase characters
    For i = 1 To LastChar
        If CharList(i).Active Then Engine_Char_Erase i
    Next i

    'Erase objects
    For i = 1 To LastObj
        If OBJList(i).Grh.GrhIndex Then Engine_OBJ_Erase i
    Next i

End Sub

Sub Engine_ConvertCPtoTP(ByVal StartPixelLeft As Integer, ByVal StartPixelTop As Integer, ByVal cx As Integer, ByVal cy As Integer, ByRef tX As Integer, ByRef tY As Integer)

'******************************************
'Converts where the user clicks in the main window
'to a tile position
'******************************************
    
    If TilePixelWidth = 0 Then Exit Sub
    If TilePixelHeight = 0 Then Exit Sub

    tX = UserPos.X + (cx - StartPixelLeft) \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.Y + (cy - StartPixelTop) \ TilePixelHeight - WindowTileHeight \ 2

End Sub

Private Function Engine_ElapsedTime() As Long

'**************************************************************
'Gets the time that past since the last call
'**************************************************************

Dim Start_Time As Long

'Get current time

    Start_Time = timeGetTime

    'Calculate elapsed time
    Engine_ElapsedTime = Start_Time - EndTime
    If Engine_ElapsedTime > 1000 Then Engine_ElapsedTime = 1000
    
    'Get next end time
    EndTime = Start_Time

End Function

Function Engine_FileExist(File As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

    Engine_FileExist = (LenB(Dir$(File, FileType)) <> 0)

End Function

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single

'************************************************************
'Gets the angle between two points in a 2d plane
'************************************************************

    On Error GoTo ErrOut
Dim SideA As Single
Dim SideC As Single

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then

        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90

            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If

        'Exit the function
        Exit Function

    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then

        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360

            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If

        'Exit the function
        Exit Function

    End If

    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)

    'Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle

    'Exit function

Exit Function

    'Check for error
ErrOut:

    'Return a 0 saying there was an error
    Engine_GetAngle = 0

Exit Function

End Function

Public Function Engine_GetTextSize(ByVal Text As String) As POINTAPI

'***************************************************
'Returns the size of text
'***************************************************
'Get the size of the text

    GetTextExtentPoint32 frmScreen.hDC, Text, Len(Text), Engine_GetTextSize

End Function

Sub Engine_Init_BodyData()

'*****************************************************************
'Loads Body.dat
'*****************************************************************
Dim LoopC As Long
Dim j As Long
'Get number of bodies

    NumBodies = CInt(Var_Get(DataPath & "Body.dat", "INIT", "NumBodies"))
    'Resize array
    ReDim BodyData(0 To NumBodies) As BodyData
    'Fill list
    For LoopC = 1 To NumBodies
        For j = 1 To 8
            Engine_Init_Grh BodyData(LoopC).Walk(j), CLng(Var_Get(DataPath & "Body.dat", Str$(LoopC), Str$(j))), 0
            Engine_Init_Grh BodyData(LoopC).Attack(j), CLng(Var_Get(DataPath & "Body.dat", Str$(LoopC), "a" & j)), 1
        Next j
        BodyData(LoopC).HeadOffset.X = CLng(Var_Get(DataPath & "Body.dat", Str$(LoopC), "HeadOffsetX"))
        BodyData(LoopC).HeadOffset.Y = CLng(Var_Get(DataPath & "Body.dat", Str$(LoopC), "HeadOffsetY"))
    Next LoopC

End Sub

Sub Engine_Init_WingData()

'*****************************************************************
'Loads Wing.dat
'*****************************************************************
Dim LoopC As Long
Dim j As Long

    'Get number of wings
    NumWings = CInt(Var_Get(DataPath & "Wing.dat", "INIT", "NumWings"))
    
    'Resize array
    ReDim WingData(0 To NumWings) As WingData
    
    'Fill list
    For LoopC = 1 To NumWings
        For j = 1 To 8
            Engine_Init_Grh WingData(LoopC).Walk(j), CLng(Var_Get(DataPath & "Wing.dat", Str(LoopC), Str(j))), 0
            Engine_Init_Grh WingData(LoopC).Attack(j), CLng(Var_Get(DataPath & "Wing.dat", Str(LoopC), "a" & j)), 1
        Next j
    Next LoopC

End Sub

Private Function Engine_Init_D3DDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS)

'************************************************************
'Initialize the Direct3D Device - start off trying with the
'best settings and move to the worst until one works
'************************************************************

Dim DispMode As D3DDISPLAYMODE          'Describes the display mode
Dim i As Byte

'When there is an error, destroy the D3D device and get ready to make a new one

    On Error GoTo ErrOut

    'Retrieve current display mode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

    'Set format for windowed mode
    D3DWindow.Windowed = 1  'State that using windowed mode
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
    D3DWindow.BackBufferFormat = DispMode.Format    'Use format just retrieved

    'Set the D3DDevices
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmScreen.hWnd, D3DCREATEFLAGS, D3DWindow)

    'Store the create flags
    UsedCreateFlags = D3DCREATEFLAGS

    'Everything was successful
    Engine_Init_D3DDevice = 1

Exit Function

ErrOut:

    'Destroy the D3DDevice so it can be remade
    Set D3DDevice = Nothing

    'Return a failure - 0
    Engine_Init_D3DDevice = 0

End Function

Sub Engine_Init_Grh(ByRef Grh As Grh, ByVal GrhIndex As Long, Optional ByVal Started As Byte = 2)

'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************

    If GrhIndex <= 0 Then
        Grh.GrhIndex = 0
        Grh.FrameCounter = 0
        Grh.LastCount = 0
        Grh.Started = 0
        Grh.SpeedCounter = 0
        Exit Sub
    End If
        
    Grh.GrhIndex = GrhIndex
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then
            Started = 0
        End If
        Grh.Started = Started
    End If
    Grh.FrameCounter = 1
    Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed

End Sub

Sub Engine_Init_GrhData()
Dim HighestTextureNum As Long
Dim FileNum As Byte
Dim Grh As Long
Dim Frame As Long
Dim TempSplit() As String
Dim j As String
Dim i As Long

    'Get Number of Graphics
    NumGrhs = CLng(Var_Get(DataPath & "Grh.ini", "INIT", "NumGrhs"))
    
    'Resize arrays
    ReDim GrhData(1 To NumGrhs) As GrhData
    ReDim GrhCatFlags(1 To NumGrhs)
    
    'Get the category information
    FileNum = FreeFile
    Open Data2Path & "GrhRaw.txt" For Input As #FileNum
        Do While EOF(FileNum) = False
            Line Input #FileNum, j
            If LenB(j) <> 0 Then
                If InStr(1, j, "(") Then
                    If InStr(1, j, "=") Then
                        
                        'Get the category flags
                        TempSplit = Split(j, "(")
                        Frame = Val(Left$(TempSplit(1), Len(TempSplit(1)) - 1))
                        
                        'Get the Grh
                        TempSplit = Split(j, "=")
                        Grh = Val(Right$(TempSplit(0), Len(TempSplit(0)) - 3))
                        
                        'Store
                        GrhCatFlags(Grh) = Frame

                    End If
                End If
            End If
        Loop
    Close #FileNum
                              
    Frame = 0
    Grh = 0
                              
    'Open files
    Open DataPath & "Grh.dat" For Binary As #1
    Seek #1, 1
    
    'Fill Grh List
    Get #1, , Grh
    
    Do Until Grh <= 0
    
        'Get number of frames
        Get #1, , GrhData(Grh).NumFrames
        If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
        
        If GrhData(Grh).NumFrames > 1 Then
        
            'Read a animation GRH set
            ReDim GrhData(Grh).Frames(1 To GrhData(Grh).NumFrames)
            For Frame = 1 To GrhData(Grh).NumFrames
                Get #1, , GrhData(Grh).Frames(Frame)
                If GrhData(Grh).Frames(Frame) <= 0 Then
                    GoTo ErrorHandler
                End If
            Next Frame
            
            Get #1, , GrhData(Grh).Speed
            GrhData(Grh).Speed = GrhData(Grh).Speed * 0.075 * EngineBaseSpeed
            If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
            
            'Compute width and height
            GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
            GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
            GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
            If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
            GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
            If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
            
        Else
        
            'Read in normal GRH data
            ReDim GrhData(Grh).Frames(1 To 1)
            Get #1, , GrhData(Grh).FileNum
            If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
            Get #1, , GrhData(Grh).sX
            If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
            Get #1, , GrhData(Grh).sY
            If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            Get #1, , GrhData(Grh).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
            Get #1, , GrhData(Grh).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
            
            If HighestTextureNum < GrhData(Grh).FileNum Then HighestTextureNum = GrhData(Grh).FileNum
            
            'Compute width and height
            GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
            GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
            GrhData(Grh).Frames(1) = Grh
            
        End If
        
        'Get Next Grh Number
        Get #1, , Grh
        
    Loop
    '************************************************
    Close #1
    
    'Get the texture descs
    Open Data2Path & "TextureDescs.txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, j
            j = Trim$(j)
            TempSplit = Split(j, "=", 2)
            If UBound(TempSplit) > 0 Then
                If Val(TempSplit(0)) > 0 Then
                    i = Val(TempSplit(0))
                    If i > NumTextureDesc Then
                        NumTextureDesc = i
                        ReDim Preserve TextureDesc(1 To i)
                    End If
                    TextureDesc(i) = Trim$(TempSplit(1))
                End If
            End If
        Loop
    Close #1

Exit Sub

ErrorHandler:
    Close #1
    MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Sub Engine_Init_HairData()

'*****************************************************************
'Loads Hair.dat
'*****************************************************************

Dim LoopC As Long
Dim i As Integer
'Get Number of hairs

    NumHairs = CInt(Var_Get(DataPath & "Hair.dat", "INIT", "NumHairs"))
    'Resize array
    ReDim HairData(0 To NumHairs) As HairData
    'Fill List
    For LoopC = 1 To NumHairs
        For i = 1 To 8
            Engine_Init_Grh HairData(LoopC).Hair(i), CLng(Var_Get(DataPath & "Hair.dat", Str$(LoopC), Str$(i))), 0
        Next i
    Next LoopC

End Sub

Sub Engine_Init_HeadData()

'*****************************************************************
'Loads Head.dat
'*****************************************************************

Dim LoopC As Long
Dim i As Integer
'Get Number of heads

    NumHeads = CInt(Var_Get(DataPath & "Head.dat", "INIT", "NumHeads"))
    'Resize array
    ReDim HeadData(0 To NumHeads) As HeadData
    'Fill List
    For LoopC = 1 To NumHeads
        For i = 1 To 8
            Engine_Init_Grh HeadData(LoopC).Head(i), CLng(Var_Get(DataPath & "Head.dat", Str$(LoopC), Str(i))), 0
            Engine_Init_Grh HeadData(LoopC).Blink(i), CLng(Var_Get(DataPath & "Head.dat", Str$(LoopC), "b" & i)), 0
            Engine_Init_Grh HeadData(LoopC).AgrHead(i), CLng(Var_Get(DataPath & "Head.dat", Str$(LoopC), "a" & i)), 0
            Engine_Init_Grh HeadData(LoopC).AgrBlink(i), CLng(Var_Get(DataPath & "Head.dat", Str$(LoopC), "ab" & i)), 0
        Next i
    Next LoopC

End Sub

Sub Engine_Init_MapData()

'*****************************************************************
'Load Map.dat
'*****************************************************************
'Get Number of Maps

    NumMaps = CInt(Var_Get(DataPath & "Map.dat", "INIT", "NumMaps"))

End Sub

Sub Engine_Init_ParticleEngine(Optional ByVal SkipToTextures As Boolean = False)

'*****************************************************************
'Loads all particles into memory - unlike normal textures, these stay in memory. This isn't
'done for any reason in particular, they just use so little memory since they are so small
'*****************************************************************
Dim i As Byte

    If Not SkipToTextures Then
    
        'Set the particles texture
        NumEffects = Var_Get(DataPath & "Game.ini", "INIT", "NumEffects")
        ReDim Effect(1 To NumEffects)
    
    End If
    
    For i = 1 To UBound(ParticleTexture())
        If ParticleTexture(i) Is Nothing Then Set ParticleTexture(i) = Nothing
        Set ParticleTexture(i) = D3DX.CreateTextureFromFileEx(D3DDevice, GrhPath & "p" & i & ".png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)
    Next i

End Sub

Private Sub Engine_Init_RenderStates()

'************************************************************
'Set the render states of the Direct3D Device
'This is in a seperate sub since if using Fullscreen and device is lost
'this is eventually called to restore settings.
'************************************************************
'Set the shader to be used

    D3DDevice.SetVertexShader FVF

    'Set the render states
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

    'Particle engine settings
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0

    'Set the texture stage stats (filters)
    '//D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    '//D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR

End Sub

Sub Engine_Init_Texture(ByVal TextureNum As Integer)

'*****************************************************************
'Loads a texture into memory
'*****************************************************************

Dim TexInfo As D3DXIMAGE_INFO_A
Dim FilePath As String

    'Get the path
    FilePath = GrhPath & TextureNum & ".png"

    'Make sure the texture exists
    If Engine_FileExist(FilePath, vbNormal) Then

        'Set the texture
        Set SurfaceDB(TextureNum) = D3DX.CreateTextureFromFileEx(D3DDevice, FilePath, D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFF000000, TexInfo, ByVal 0)

        'Set the size
        SurfaceSize(TextureNum).X = TexInfo.Width
        SurfaceSize(TextureNum).Y = TexInfo.Height

        'Set the texture timer
        SurfaceTimer(TextureNum) = SurfaceTimerMax

    End If

End Sub

Function Engine_Init_TileEngine(ByRef setDisplayFormhWnd As Long, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal Engine_Speed As Single) As Boolean

'*****************************************************************
'Init Tile Engine
'*****************************************************************

Dim Fnt As New StdFont
Dim i As Long
Dim s As String
    
    'Set the text height
    For i = 0 To 255
        s = s & Chr$(i)
    Next i
    lngTextHeight = Engine_GetTextSize(s).Y
    SfxPath = Var_Get(DataPath & "Game.ini", "INIT", "SoundPath")

    'Fill startup variables
    DisplayFormhWnd = setDisplayFormhWnd
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    EngineBaseSpeed = Engine_Speed

    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder

    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = 36
    ScrollPixelsPerFrameY = 36

    'Set the array sizes by the number of graphic files
    NumGrhFiles = CLng(Var_Get(DataPath & "Grh.ini", "INIT", "NumGrhFiles"))
    ReDim SurfaceDB(1 To NumGrhFiles)
    ReDim SurfaceSize(1 To NumGrhFiles)
    ReDim SurfaceTimer(1 To NumGrhFiles)

    '****** INIT DirectX ******
    ' Create the root D3D objects
    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate()
    Set D3DX = New D3DX8

    'Create the D3D Device
    If Engine_Init_D3DDevice(D3DCREATE_PUREDEVICE) = 0 Then
        If Engine_Init_D3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) = 0 Then
            If Engine_Init_D3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) = 0 Then
                MsgBox "Could not init D3DDevice. Exiting..."
                Engine_Init_UnloadTileEngine
                Engine_UnloadAllForms
                End
            End If
        End If
    End If
    Engine_Init_RenderStates

    'Set up the font
    Fnt.Name = "Arial"
    Fnt.Size = 8
    Fnt.Bold = False
    Set MainFontDesc = Fnt
    Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)

    'Load graphic data into memory
    Engine_Init_GrhData
    Engine_Init_BodyData
    Engine_Init_WingData
    Engine_Init_WeaponData
    Engine_Init_HeadData
    Engine_Init_HairData
    Engine_Init_MapData
    Engine_Init_ParticleEngine

    'Initialize DirectSound
    '//InitializeSound DirectX

    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60

    'Set high resolution timer
    timeBeginPeriod 1
    EndTime = timeGetTime

    'Start the engine
    Engine_Init_TileEngine = True
    EngineRun = True

    UserPos.X = 13
    UserPos.Y = 10

End Function

Public Sub Engine_Init_UnloadTileEngine()

'*****************************************************************
'Shutsdown engine
'*****************************************************************

    On Error Resume Next

    Dim LoopC As Long

        EngineRun = False

        '****** Clear DirectX objects ******
        Set DX = Nothing
        Set D3DDevice = Nothing
        Set MainFont = Nothing
        Set D3DX = Nothing
        
        'Clear GRH memory
        For LoopC = 1 To NumGrhFiles
            Set SurfaceDB(LoopC) = Nothing
        Next LoopC

        For LoopC = 1 To UBound(ParticleTexture)
            Set ParticleTexture(LoopC) = Nothing
        Next LoopC
        
        'Clear DirectSound objects
        '//DeInitiliazeSound

End Sub

Sub Engine_Init_WeaponData()

'*****************************************************************
'Loads Weapon.dat
'*****************************************************************

Dim LoopC As Long
'Get number of weapons

    NumWeapons = CLng(Var_Get(DataPath & "Weapon.dat", "INIT", "NumWeapons"))
    'Resize array
    ReDim WeaponData(0 To NumWeapons) As WeaponData
    'Fill listn
    For LoopC = 1 To NumWeapons
        Engine_Init_Grh WeaponData(LoopC).Walk(1), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk1")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(2), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk2")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(3), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk3")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(4), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk4")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(5), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk5")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(6), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk6")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(7), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk7")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(8), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk8")), 0
        Engine_Init_Grh WeaponData(LoopC).Attack(1), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack1")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(2), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack2")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(3), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack3")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(4), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack4")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(5), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack5")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(6), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack6")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(7), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack7")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(8), CLng(Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack8")), 1
    Next LoopC

End Sub

Sub Engine_Weather_Update()

'*****************************************************************
'Initializes the weather effects
'*****************************************************************

    'Check if we're using weather
    'If UseWeather = 0 Then Exit Sub

    'Only update the weather settings if it has changed!
    If LastWeather <> MapInfo.Weather Then
    
        'Set the lastweather to the current weather
        LastWeather = MapInfo.Weather
        
        'Erase sounds
        'Sound_Erase WeatherSfx1
        'Sound_Erase WeatherSfx2
    
        Select Case LastWeather
        
        Case 1  'Snow (light fall)
            If WeatherEffectIndex <= 0 Then
                WeatherEffectIndex = Effect_Snow_Begin(1, 400)
            ElseIf Effect(WeatherEffectIndex).EffectNum <> EffectNum_Snow Then
                Effect_Kill WeatherEffectIndex
                WeatherEffectIndex = Effect_Snow_Begin(1, 400)
            ElseIf Not Effect(WeatherEffectIndex).Used Then
                WeatherEffectIndex = Effect_Snow_Begin(1, 400)
            End If
            WeatherDoLightning = 0
            WeatherDoFog = 0
            
        Case 2  'Rain Storm (heavy rain + lightning)
            If WeatherEffectIndex <= 0 Then
                WeatherEffectIndex = Effect_Rain_Begin(9, 300)
            ElseIf Effect(WeatherEffectIndex).EffectNum <> EffectNum_Rain Then
                Effect_Kill WeatherEffectIndex
                WeatherEffectIndex = Effect_Rain_Begin(9, 300)
            ElseIf Not Effect(WeatherEffectIndex).Used Then
                WeatherEffectIndex = Effect_Rain_Begin(9, 300)
            End If
            WeatherDoLightning = 1  'We take our rain with a bit of lightning on top >:D
            WeatherDoFog = 0
            'Sound_Set WeatherSfx1, 3
            'Sound_Set WeatherSfx2, 2
            'Sound_Play WeatherSfx1, DSBPLAY_LOOPING
            
        Case 3  'Inside of a house in a storm (lightning + muted rain sound)
            If WeatherEffectIndex > 0 Then  'Kill the weather effect if used
                If Effect(WeatherEffectIndex).Used Then Effect_Kill WeatherEffectIndex
            End If
            WeatherDoLightning = 1
            WeatherDoFog = 0
            'Sound_Set WeatherSfx1, 4
            'Sound_Set WeatherSfx2, 6
            'Sound_Play WeatherSfx1, DSBPLAY_LOOPING
            
        Case 4  'Inside of a cave in a storm (lightning + muted rain sound + fog)
            If WeatherEffectIndex > 0 Then  'Kill the weather effect if used
                If Effect(WeatherEffectIndex).Used Then Effect_Kill WeatherEffectIndex
            End If
            WeatherDoLightning = 1
            WeatherDoFog = 10    'This will make it nice and spooky! >:D
            'Sound_Set WeatherSfx1, 4
            'Sound_Set WeatherSfx2, 6
            'Sound_Play WeatherSfx1, DSBPLAY_LOOPING
            
        Case Else   'None
            If WeatherEffectIndex > 0 Then  'Kill the weather effect if used
                If Effect(WeatherEffectIndex).Used Then Effect_Kill WeatherEffectIndex
                'Sound_Erase WeatherSfx1  'Remove the sounds
                'Sound_Erase WeatherSfx2
            End If
            WeatherDoLightning = 0
            WeatherDoFog = 0
            
        End Select
        
    End If
    
    'Update fog
    If WeatherDoFog Then Engine_Weather_UpdateFog

    'Update lightning
    If WeatherDoLightning Then Engine_Weather_UpdateLightning

End Sub

Sub Engine_Weather_UpdateFog()

'*****************************************************************
'Update the fog effects
'*****************************************************************
Dim TempGrh As Grh
Dim i As Long
Dim X As Long
Dim Y As Long
Dim c As Long

    'Make sure we have the fog value
    If WeatherFogCount = 0 Then WeatherFogCount = 13
    
    'Update the fog's position
    WeatherFogX1 = WeatherFogX1 + (ElapsedTime * (0.018 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY1 = WeatherFogY1 + (ElapsedTime * (0.013 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)
    Do While WeatherFogX1 < -512
        WeatherFogX1 = WeatherFogX1 + 512
    Loop
    Do While WeatherFogY1 < -512
        WeatherFogY1 = WeatherFogY1 + 512
    Loop
    Do While WeatherFogX1 > 0
        WeatherFogX1 = WeatherFogX1 - 512
    Loop
    Do While WeatherFogY1 > 0
        WeatherFogY1 = WeatherFogY1 - 512
    Loop
    
    WeatherFogX2 = WeatherFogX2 - (ElapsedTime * (0.037 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY2 = WeatherFogY2 - (ElapsedTime * (0.021 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)
    Do While WeatherFogX2 < -512
        WeatherFogX2 = WeatherFogX2 + 512
    Loop
    Do While WeatherFogY2 < -512
        WeatherFogY2 = WeatherFogY2 + 512
    Loop
    Do While WeatherFogX2 > 0
        WeatherFogX2 = WeatherFogX2 - 512
    Loop
    Do While WeatherFogY2 > 0
        WeatherFogY2 = WeatherFogY2 - 512
    Loop

    TempGrh.FrameCounter = 1
    
    'Render fog 2
    TempGrh.GrhIndex = 4
    X = 2
    Y = -1
    c = D3DColorARGB(100, 255, 255, 255)
    For i = 1 To WeatherFogCount
        Engine_Render_Grh TempGrh, (X * 512) + WeatherFogX2, (Y * 512) + WeatherFogY2, 0, 0, False, c, c, c, c
        X = X + 1
        If X > (1 + (ScreenWidth \ 512)) Then
            X = 0
            Y = Y + 1
        End If
    Next i
            
    'Render fog 1
    TempGrh.GrhIndex = 3
    X = 0
    Y = 0
    c = D3DColorARGB(75, 255, 255, 255)
    For i = 1 To WeatherFogCount
        Engine_Render_Grh TempGrh, (X * 512) + WeatherFogX1, (Y * 512) + WeatherFogY1, 0, 0, False, c, c, c, c
        X = X + 1
        If X > (2 + (ScreenWidth \ 512)) Then
            X = 0
            Y = Y + 1
        End If
    Next i

End Sub

Sub Engine_Weather_UpdateLightning()

'*****************************************************************
'Updates the lightning count-down and creates the flash if its ready
'*****************************************************************
Dim X As Long
Dim Y As Long
Dim i As Long

    'Check if we are in the middle of a flash
    If FlashTimer > 0 Then
        FlashTimer = FlashTimer - ElapsedTime
        
        'The flash has run out
        If FlashTimer <= 0 Then
            If BrightChkValue = 0 Then
            
                'Change the light of all the tiles back
                For X = 1 To MapInfo.Width
                    For Y = 1 To MapInfo.Height
                        For i = 1 To 24
                            MapData(X, Y).Light(i) = SaveLightBuffer(X, Y).Light(i)
                        Next i
                    Next Y
                Next X
            
            End If
        End If
        
    'Update the timer, see if it is time to flash
    Else
        LightningTimer = LightningTimer - ElapsedTime
        
        'Flash me, baby!
        If LightningTimer <= 0 Then
            LightningTimer = 15000 + (Rnd * 15000)  'Reset timer (flash every 15 to 30 seconds)
            FlashTimer = 250    'How long the flash is (miliseconds)
            
            'Randomly place the lightning
            LightningX = 50 + Rnd * 700
            LightningY = Rnd * -200
            'Sound_Play WeatherSfx2, DSBPLAY_DEFAULT  'BAM!
            
            'Change the light of all the tiles to white
            For X = 1 To MapInfo.Width
                For Y = 1 To MapInfo.Height
                    For i = 1 To 24
                        MapData(X, Y).Light(i) = -1
                    Next i
                Next Y
            Next X
            
        End If
        
    End If

End Sub

Function Engine_LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean

'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************

Dim i As Integer

'Check to see if its out of bounds

    If X < MinXBorder Then Exit Function
    If X > MaxXBorder Then Exit Function
    If Y < MinYBorder Then Exit Function
    If Y > MaxYBorder Then Exit Function

    'Check to see if its blocked
    If MapData(X, Y).Blocked = 1 Then Exit Function

    'Check for character
    For i = 1 To LastChar
        If CharList(i).Active Then
            If CharList(i).Pos.X = X Then
                If CharList(i).Pos.Y = Y Then Exit Function
            End If
        End If
    Next i

    'The position is legal
    Engine_LegalPos = True

End Function

Sub Engine_Input_CheckKeys()

    If GetActiveWindow = 0 Then Exit Sub

    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
    
        'Move Up-Right
        If GetKeyState(vbKeyUp) < 0 And GetKeyState(vbKeyRight) < 0 Then
            Engine_MoveUser NORTHEAST
            Exit Sub
        End If
        'Move Up-Left
        If GetKeyState(vbKeyUp) < 0 And GetKeyState(vbKeyLeft) < 0 Then
            Engine_MoveUser NORTHWEST
            Exit Sub
        End If
        'Move Down-Right
        If GetKeyState(vbKeyDown) < 0 And GetKeyState(vbKeyRight) < 0 Then
            Engine_MoveUser SOUTHEAST
            Exit Sub
        End If
        'Move Down-Left
        If GetKeyState(vbKeyDown) < 0 And GetKeyState(vbKeyLeft) < 0 Then
            Engine_MoveUser SOUTHWEST
            Exit Sub
        End If
        'Move Up
        If GetKeyState(vbKeyUp) < 0 Then
            Engine_MoveUser NORTH
            Exit Sub
        End If
        'Move Right
        If GetKeyState(vbKeyRight) < 0 Then
            Engine_MoveUser EAST
            Exit Sub
        End If
        'Move down
        If GetKeyState(vbKeyDown) < 0 Then
            Engine_MoveUser SOUTH
            Exit Sub
        End If
        'Move left
        If GetKeyState(vbKeyLeft) < 0 Then
            Engine_MoveUser WEST
            Exit Sub
        End If
        
        'Move Up-Right
        If GetKeyState(vbKeyW) < 0 And GetKeyState(vbKeyD) < 0 Then
            Engine_MoveUser NORTHEAST
            Exit Sub
        End If
        'Move Up-Left
        If GetKeyState(vbKeyW) < 0 And GetKeyState(vbKeyA) < 0 Then
            Engine_MoveUser NORTHWEST
            Exit Sub
        End If
        'Move Down-Right
        If GetKeyState(vbKeyS) < 0 And GetKeyState(vbKeyD) < 0 Then
            Engine_MoveUser SOUTHEAST
            Exit Sub
        End If
        'Move Down-Left
        If GetKeyState(vbKeyS) < 0 And GetKeyState(vbKeyA) < 0 Then
            Engine_MoveUser SOUTHWEST
            Exit Sub
        End If
        'Move Up
        If GetKeyState(vbKeyW) < 0 Then
            Engine_MoveUser NORTH
            Exit Sub
        End If
        'Move Right
        If GetKeyState(vbKeyD) < 0 Then
            Engine_MoveUser EAST
            Exit Sub
        End If
        'Move down
        If GetKeyState(vbKeyS) < 0 Then
            Engine_MoveUser SOUTH
            Exit Sub
        End If
        'Move left
        If GetKeyState(vbKeyA) < 0 Then
            Engine_MoveUser WEST
            Exit Sub
        End If

    End If

End Sub

Sub Engine_MoveScreen(ByVal Heading As Byte)

'******************************************
'Starts the screen moving in a direction
'******************************************

Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer

    'Figure out which way to move
    Select Case Heading
        Case NORTH
            Y = -1
        Case EAST
            X = 1
        Case SOUTH
            Y = 1
        Case WEST
            X = -1
        Case NORTHEAST
            Y = -1
            X = 1
        Case SOUTHEAST
            Y = 1
            X = 1
        Case SOUTHWEST
            Y = 1
            X = -1
        Case NORTHWEST
            Y = -1
            X = -1
    End Select
    
    If X = -1 And X + UserPos.X < MinXBorder Then Exit Sub
    If X = 1 And X + UserPos.X > MaxXBorder Then Exit Sub
    If Y = -1 And Y + UserPos.Y < MinYBorder Then Exit Sub
    If Y = 1 And Y + UserPos.Y > MaxYBorder Then Exit Sub
    
    AddtoUserPos.X = X
    AddtoUserPos.Y = Y
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y

    If tX < MinXBorder Then tX = MinXBorder
    If tX > MaxXBorder Then tX = MaxXBorder
    If tY < MinYBorder Then tY = MinYBorder
    If tY > MaxYBorder Then tY = MaxYBorder

    'Start moving... MainLoop does the rest
    UserPos.X = tX
    UserPos.Y = tY
    UserMoving = True

End Sub

Sub Engine_MoveUser(ByVal Direction As Byte)

'*****************************************************************
'Move user in appropriate direction
'*****************************************************************

Dim aX As Integer
Dim aY As Integer
'Figure out the AddX and AddY values

    Select Case Direction
    Case NORTHEAST
        aX = 1
        aY = -1
    Case NORTHWEST
        aX = -1
        aY = -1
    Case SOUTHEAST
        aX = 1
        aY = 1
    Case SOUTHWEST
        aX = -1
        aY = 1
    Case NORTH
        aX = 0
        aY = -1
    Case EAST
        aX = 1
        aY = 0
    Case SOUTH
        aX = 0
        aY = 1
    Case WEST
        aX = -1
        aY = 0
    End Select

    'Move the char
    Engine_MoveScreen Direction

End Sub

Sub Engine_OBJ_Create(ByVal GrhIndex As Long, ByVal X As Byte, ByVal Y As Byte)

'*****************************************************************
'Create an object on the map and update LastOBJ value
'*****************************************************************

Dim ObjIndex As Integer

'Get the next open obj slot

    Do
        ObjIndex = ObjIndex + 1

        'Update LastObj if we go over the size of the current array
        If ObjIndex > LastObj Then
            LastObj = ObjIndex
            ReDim Preserve OBJList(1 To ObjIndex)
            Exit Do
        End If

    Loop While OBJList(ObjIndex).Grh.GrhIndex > 0

    'Set the object position
    OBJList(ObjIndex).Pos.X = X
    OBJList(ObjIndex).Pos.Y = Y

    'Create the object
    Engine_Init_Grh OBJList(ObjIndex).Grh, GrhIndex

End Sub

Sub Engine_OBJ_Erase(ByVal ObjIndex As Integer)

'*****************************************************************
'Erase an object from the map and update the LastOBJ value
'*****************************************************************

Dim j As Integer

'Check for a valid object

    If ObjIndex > LastObj Then Exit Sub
    If ObjIndex <= 0 Then Exit Sub

    'Erase the object
    OBJList(ObjIndex).Grh.GrhIndex = 0
    OBJList(ObjIndex).Pos.X = 0
    OBJList(ObjIndex).Pos.Y = 0

    'Update LastOBJ
    If j = LastObj Then
        Do Until OBJList(LastObj).Grh.GrhIndex > 1
            'Move down one object
            LastObj = LastObj - 1
            If LastObj = 0 Then Exit Do
        Loop
        If j <> LastObj Then
            'We still have objects, resize the array to end at the last used slot
            If j <> 0 Then
                ReDim Preserve OBJList(1 To LastObj)
            Else
                ReDim OBJList(1 To 1)
            End If
        End If
    End If

End Sub

Function Engine_PixelPosX(ByVal X As Integer) As Integer

'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************

    Engine_PixelPosX = (X - 1) * TilePixelWidth

End Function

Function Engine_PixelPosY(ByVal Y As Integer) As Integer

'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************

    Engine_PixelPosY = (Y - 1) * TilePixelHeight

End Function

Private Function Engine_ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String

'*****************************************************************
'Gets a field from a string
'*****************************************************************

Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String

    Seperator = Chr$(SepASCII)

    LastPos = 0
    FieldNum = 0
    For i = 1 To Len(Text)
        CurChar = Mid$(Text, i, 1)
        If CurChar = Seperator Then
            FieldNum = FieldNum + 1
            If FieldNum = Pos Then
                Engine_ReadField = Mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    If FieldNum = Pos Then
        Engine_ReadField = Mid$(Text, LastPos + 1)
    End If

End Function

Function Engine_RectCollision(ByVal x1 As Integer, ByVal y1 As Integer, ByVal Width1 As Integer, ByVal Height1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal Width2 As Integer, ByVal Height2 As Integer)

'******************************************
'Check for collision between two rectangles
'******************************************

Dim RetRect As RECT
Dim Rect1 As RECT
Dim RECT2 As RECT

'Build the rectangles

    Rect1.Left = x1
    Rect1.Right = x1 + Width1
    Rect1.Top = y1
    Rect1.bottom = y1 + Height1
    RECT2.Left = x2
    RECT2.Right = x2 + Width2
    RECT2.Top = y2
    RECT2.bottom = y2 + Height2

    'Call collision API
    Engine_RectCollision = IntersectRect(RetRect, Rect1, RECT2)

End Function

Public Sub Engine_SetTileSelectionArray()

'***************************************************
'Create the tile selection array, starting at tsStart and skipping unused graphics
'***************************************************
Dim CurrentGrh As Long
Dim i As Long
Dim j As Long
Dim b As Byte

    'We are changing out list, so we will draw everything again
    frmTileSelect.Cls
    tsDrawAll = 1

    'Check for valid values
    If tsStart < 1 Then tsStart = 1
    
    'Set the current grh
    CurrentGrh = tsStart
    i = 1
    
    'Loop until we hit either the end of the grhdata array, or the end of the previewgrhlist array
    Do
        If CurrentGrh > UBound(GrhData) Then Exit Do
        
        'Only use graphics in use
        If GrhData(CurrentGrh).NumFrames > 0 Then
            
            'Reset variable
            b = 0
        
            'Put uncategorized graphics in the Misc(Hidden) section
            If GrhCatFlags(CurrentGrh) = 0 Then
                If frmTSOpt.CatChk(7).Value = 1 Then b = 1
            End If
                
            'If the appropriate flags are ticked
            If Not b Then
                For j = frmTSOpt.CatChk.LBound To frmTSOpt.CatChk.UBound
                    If frmTSOpt.CatChk(j).Value = 1 Then
                        If GrhCatFlags(CurrentGrh) And (2 ^ j) Then
                            b = 1
                            Exit For
                        End If
                    End If
                Next j
            End If
                
            'Draw the grh
            If b Then

                'Set the graphic
                Engine_Init_Grh PreviewGrhList(i), CurrentGrh
                
                'Check if we hit the end of the list
                i = i + 1
                If i > UBound(PreviewGrhList) Then Exit Do
            
            End If
        End If
        
        'Update the currentgrh value - allows us to set the next grh
        CurrentGrh = CurrentGrh + 1
        
    Loop

End Sub

Public Sub Engine_Render_TileSelection()

'***************************************************
'Draw the tile selection screen - this completely replaces the tile engine drawing
'***************************************************
Dim dest As RECT
Dim src As RECT
Dim i As Long
Dim X As Long
Dim Y As Long
Dim j As Integer

    'Check for valid values
    If tsTileWidth = 0 Then tsTileWidth = 32
    If tsTileHeight = 0 Then tsTileHeight = 32
    
    'Set up our display rect - this is so we only grab a 64x64 chunk out of the buffer instead of the whole thing
    src.Right = tsTileWidth
    src.bottom = tsTileHeight

    'Loop through the array
    X = 0
    Y = 0
    
    For i = 1 To UBound(PreviewGrhList)
    
        'If we ever hit a point where the GrhIndex = 0, then we must have hit the end since we do not put empty graphics in the array
        If PreviewGrhList(i).GrhIndex = 0 Then Exit For
        
        'Update the position to render
        Y = Y + 1
        If Y = tsHeight Then
            Y = 0
            X = X + 1                       '                                                                 *cling*
            If X = tsWidth Then Exit For    'We've run out of space, omg!! congratz, yer dun rendurin!! (>^_^)>[][]<(^_^<)
        End If
        dest.Top = Y * tsTileHeight
        dest.Left = X * tsTileWidth
        dest.Right = dest.Left + tsTileWidth
        dest.bottom = dest.Top + tsTileHeight

        'Render - this is a very messy method and I dont reccomend it, but I am trying to prevent using another device ;)
        'We can not render it all at once because our buffer will be the screen size (default 800x600), so it'd be too small
        'Unless this is the first time rendering this list, we will only draw updated animations
        If GrhData(PreviewGrhList(i).GrhIndex).NumFrames > 1 Or tsDrawAll = 1 Then
            If Engine_ValidateDevice Then
                j = Int(PreviewGrhList(i).FrameCounter)
                D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
                D3DDevice.BeginScene
                Engine_Render_Grh PreviewGrhList(i), 0, 0, 0, 1
                If tsDrawAll = 1 Then j = -1    'tsDrawAll flag will force the draw
                If j = Int(PreviewGrhList(i).FrameCounter) Then
                    D3DDevice.EndScene
                Else
                    Engine_Render_Text PreviewGrhList(i).GrhIndex, 0, 0, 30, -16777216
                    Engine_Render_Text PreviewGrhList(i).GrhIndex, 1, 1, 30, -1
                    If GrhData(PreviewGrhList(i).GrhIndex).NumFrames > 1 Then Engine_Render_Text "A", tsTileWidth - 8, tsTileHeight - lngTextHeight, 8, -16711936
                    D3DDevice.EndScene
                    D3DDevice.Present src, dest, frmTileSelect.hWnd, ByVal 0
                End If
            End If
        End If
        
    Next i
    
    'Clear the "draw all"
    tsDrawAll = 0
    
    'Update FPS shiznites
    ElapsedTime = 17    'Run at about 60 FPS - the fact that we only draw sometimes really throws the FPS out of wack if done normally
    TickPerFrame = (ElapsedTime * EngineBaseSpeed)
    TimerMultiplier = TickPerFrame * 0.075

End Sub

Private Sub Engine_Render_Char(ByVal CharIndex As Long, ByVal PixelOffsetX As Single, ByVal PixelOffsetY As Single)

'***************************************************
'Draw a character to the screen by the CharIndex
'First variables are set, then all shadows drawn, then character drawn, then extras (emoticons, icons, etc)
'Any variables not handled in "Set the variables" are set in Shadow calls - do not call a second time in the
'normal character rendering calls
'***************************************************

Dim TempGrh As Grh
Dim Moved As Boolean
Dim IconCount As Byte
Dim IconOffset As Integer
Dim LoopC As Byte
Dim Green As Byte
Dim RenderColor(1 To 4) As Long
Dim TempBlock As MapBlock
Dim TempBlock2 As MapBlock
Dim HeadGrh As Grh
Dim BodyGrh As Grh
Dim WeaponGrh As Grh
Dim HairGrh As Grh
Dim WingsGrh As Grh

'***** Set the variables *****

'Set the map block the char is on to the TempBlock, and the block above the user as TempBlock2

    TempBlock = MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y)
    If CharList(CharIndex).Pos.Y > 1 Then TempBlock2 = MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y - 1)

    RenderColor(1) = TempBlock2.Light(1)
    RenderColor(2) = TempBlock2.Light(2)
    RenderColor(3) = TempBlock.Light(3)
    RenderColor(4) = TempBlock.Light(4)

    If CharList(CharIndex).Moving Then

        'If needed, move left and right
        If CharList(CharIndex).ScrollDirectionX <> 0 Then
            CharList(CharIndex).MoveOffset.X = CharList(CharIndex).MoveOffset.X + ScrollPixelsPerFrameX * Sgn(CharList(CharIndex).ScrollDirectionX) * TickPerFrame

            'Start animation
            CharList(CharIndex).Body.Walk(CharList(CharIndex).Heading).Started = 1

            'Char moved
            Moved = True

            'Check if we already got there
            If (Sgn(CharList(CharIndex).ScrollDirectionX) = 1 And CharList(CharIndex).MoveOffset.X >= 0) Or (Sgn(CharList(CharIndex).ScrollDirectionX) = -1 And CharList(CharIndex).MoveOffset.X <= 0) Then
                CharList(CharIndex).MoveOffset.X = 0
                CharList(CharIndex).ScrollDirectionX = 0
            End If

        End If

        'If needed, move up and down
        If CharList(CharIndex).ScrollDirectionY <> 0 Then
            CharList(CharIndex).MoveOffset.Y = CharList(CharIndex).MoveOffset.Y + ScrollPixelsPerFrameY * Sgn(CharList(CharIndex).ScrollDirectionY) * TickPerFrame

            'Start animation
            CharList(CharIndex).Body.Walk(CharList(CharIndex).Heading).Started = 1

            'Char moved
            Moved = True

            'Check if we already got there
            If (Sgn(CharList(CharIndex).ScrollDirectionY) = 1 And CharList(CharIndex).MoveOffset.Y >= 0) Or (Sgn(CharList(CharIndex).ScrollDirectionY) = -1 And CharList(CharIndex).MoveOffset.Y <= 0) Then
                CharList(CharIndex).MoveOffset.Y = 0
                CharList(CharIndex).ScrollDirectionY = 0
            End If

        End If
    End If

    'Update movement reset timer
    If CharList(CharIndex).ScrollDirectionX = 0 Or CharList(CharIndex).ScrollDirectionY = 0 Then

        'If done moving stop animation
        If Not Moved Then
            If CharList(CharIndex).Body.Walk(CharList(CharIndex).Heading).Started Then

                'Stop animation
                CharList(CharIndex).Body.Walk(CharList(CharIndex).Heading).Started = 0
                CharList(CharIndex).Body.Walk(CharList(CharIndex).Heading).FrameCounter = 1
                CharList(CharIndex).Moving = False
                If CharList(CharIndex).ActionIndex = 1 Then CharList(CharIndex).ActionIndex = 0

            End If
        End If
    End If

    'Set the pixel offset
    PixelOffsetX = PixelOffsetX + CharList(CharIndex).MoveOffset.X
    PixelOffsetY = PixelOffsetY + CharList(CharIndex).MoveOffset.Y

    'Save the values in the realpos variable
    CharList(CharIndex).RealPos.X = PixelOffsetX
    CharList(CharIndex).RealPos.Y = PixelOffsetY

    '***** Render Shadows *****

    'Draw Body
    If CharList(CharIndex).ActionIndex <= 1 Then

        'Shadow
        Engine_Render_Grh CharList(CharIndex).Body.Walk(CharList(CharIndex).Heading), PixelOffsetX, PixelOffsetY, 1, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, , 1
        Engine_Render_Grh CharList(CharIndex).Weapon.Walk(CharList(CharIndex).Heading), PixelOffsetX, PixelOffsetY, True, True, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, , 1

    Else

        'Start attack animation
        CharList(CharIndex).Body.Attack(CharList(CharIndex).Heading).Started = 0
        CharList(CharIndex).Weapon.Attack(CharList(CharIndex).Heading).FrameCounter = 1

        'Shadow
        Engine_Render_Grh CharList(CharIndex).Body.Attack(CharList(CharIndex).Heading), PixelOffsetX, PixelOffsetY, 1, 1, False, ShadowColor, ShadowColor, ShadowColor, ShadowColor, , 1
        Engine_Render_Grh CharList(CharIndex).Weapon.Attack(CharList(CharIndex).Heading), PixelOffsetX, PixelOffsetY, True, True, False, ShadowColor, ShadowColor, ShadowColor, ShadowColor, , 1

        'Check if animation has stopped
        If CharList(CharIndex).Body.Attack(CharList(CharIndex).Heading).Started = 0 Then CharList(CharIndex).ActionIndex = 0

    End If

    'Draw Head
    If CharList(CharIndex).Aggressive > 0 Then
        'Aggressive
        If CharList(CharIndex).BlinkTimer > 0 Then
            CharList(CharIndex).BlinkTimer = CharList(CharIndex).BlinkTimer - ElapsedTime
            'Blinking
            Engine_Render_Grh CharList(CharIndex).Head.AgrBlink(CharList(CharIndex).HeadHeading), PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, True, False, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, , 1
        Else
            'Normal
            Engine_Render_Grh CharList(CharIndex).Head.AgrHead(CharList(CharIndex).HeadHeading), PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, True, False, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, , 1
        End If
    Else
        'Not Aggressive
        If CharList(CharIndex).BlinkTimer > 0 Then
            CharList(CharIndex).BlinkTimer = CharList(CharIndex).BlinkTimer - ElapsedTime
            'Blinking
            Engine_Render_Grh CharList(CharIndex).Head.Blink(CharList(CharIndex).HeadHeading), PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, True, False, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, , 1
        Else
            'Normal
            Engine_Render_Grh CharList(CharIndex).Head.Head(CharList(CharIndex).HeadHeading), PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, True, False, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, , 1
        End If
    End If

    'Hair
    Engine_Render_Grh CharList(CharIndex).Hair.Hair(CharList(CharIndex).HeadHeading), PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, True, False, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, , 1

    '***** Render Character *****
    '***** (When updating this, make sure you copy it to the NPCEditor and MapEditor, too!) *****
    CharList(CharIndex).Weapon.Walk(CharList(CharIndex).Heading).FrameCounter = CharList(CharIndex).Body.Walk(CharList(CharIndex).Heading).FrameCounter

    'The body, weapon and wings
    If CharList(CharIndex).ActionIndex <= 1 Then
        'Walking
        BodyGrh = CharList(CharIndex).Body.Walk(CharList(CharIndex).Heading)
        WeaponGrh = CharList(CharIndex).Weapon.Walk(CharList(CharIndex).Heading)
        WingsGrh = CharList(CharIndex).Wings.Walk(CharList(CharIndex).Heading)
    Else
        'Attacking
        BodyGrh = CharList(CharIndex).Body.Attack(CharList(CharIndex).Heading)
        WeaponGrh = CharList(CharIndex).Weapon.Attack(CharList(CharIndex).Heading)
        WingsGrh = CharList(CharIndex).Wings.Attack(CharList(CharIndex).Heading)
    End If
    
    'The head
    If CharList(CharIndex).Aggressive > 0 Then  'Aggressive
        If CharList(CharIndex).BlinkTimer > 0 Then HeadGrh = CharList(CharIndex).Head.AgrBlink(CharList(CharIndex).HeadHeading) Else HeadGrh = CharList(CharIndex).Head.AgrHead(CharList(CharIndex).HeadHeading)
    Else    'Non-aggressive
        If CharList(CharIndex).BlinkTimer > 0 Then HeadGrh = CharList(CharIndex).Head.Blink(CharList(CharIndex).HeadHeading) Else HeadGrh = CharList(CharIndex).Head.Head(CharList(CharIndex).HeadHeading)
    End If
    
    'The hair
    HairGrh = CharList(CharIndex).Hair.Hair(CharList(CharIndex).HeadHeading)
    
    'Make the paperdoll layering based off the direction they are heading
        
    '*** NORTH / NORTHEAST *** (1.Weapon 2.Body 3.Head 4.Hair 5.Wings)
    If CharList(CharIndex).Heading = NORTH Or CharList(CharIndex).Heading = NORTHEAST Then
        Engine_Render_Grh WeaponGrh, PixelOffsetX, PixelOffsetY, True, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh BodyGrh, PixelOffsetX, PixelOffsetY, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh HeadGrh, PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh HairGrh, PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh WingsGrh, PixelOffsetX, PixelOffsetY, True, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        
    '*** EAST / SOUTHEAST *** (1.Body 2.Head 3.Hair 4.Wings 5.Weapon)
    ElseIf CharList(CharIndex).Heading = EAST Or CharList(CharIndex).Heading = SOUTHEAST Then
        Engine_Render_Grh BodyGrh, PixelOffsetX, PixelOffsetY, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh HeadGrh, PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh HairGrh, PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh WingsGrh, PixelOffsetX, PixelOffsetY, True, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh WeaponGrh, PixelOffsetX, PixelOffsetY, True, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        
    '*** SOUTH / SOUTHWEST *** (1.Wings 2.Body 3.Head 4.Hair 5.Weapon)
    ElseIf CharList(CharIndex).Heading = SOUTH Or CharList(CharIndex).Heading = SOUTHWEST Then
        Engine_Render_Grh WingsGrh, PixelOffsetX, PixelOffsetY, True, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh BodyGrh, PixelOffsetX, PixelOffsetY, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh HeadGrh, PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh HairGrh, PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh WeaponGrh, PixelOffsetX, PixelOffsetY, True, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        
    '*** WEST / NORTHWEST *** (1.Weapon 1.Body 2.Head 3.Hair 4.Wings)
    ElseIf CharList(CharIndex).Heading = WEST Or CharList(CharIndex).Heading = NORTHWEST Then
        Engine_Render_Grh WeaponGrh, PixelOffsetX, PixelOffsetY, True, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh BodyGrh, PixelOffsetX, PixelOffsetY, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh HeadGrh, PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh HairGrh, PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh WingsGrh, PixelOffsetX, PixelOffsetY, True, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        
    End If
    
    '***** Render Extras *****

    'Check to draw the health or not
    If CharList(CharIndex).HealthPercent > 0 Then

        'Draw name/health over head
        Engine_Render_Text CharList(CharIndex).Name & " " & CharList(CharIndex).HealthPercent & "%", PixelOffsetX - 32, PixelOffsetY - 36, 96, RenderColor(1), DT_TOP Or DT_CENTER

    Else

        'Draw name over head
        Engine_Render_Text CharList(CharIndex).Name, PixelOffsetX - 32, PixelOffsetY - 36, 96, RenderColor(1), DT_TOP Or DT_CENTER

    End If

End Sub

Private Function Engine_UpdateGrh(ByRef Grh As Grh, Optional ByVal LoopAnim As Boolean = True) As Boolean

'*****************************************************************
'Updates the grh's animation
'*****************************************************************

    'Check that the grh is started
    If Grh.Started = 1 Then
    
        'Update the frame counter
        Grh.FrameCounter = Grh.FrameCounter + ((timeGetTime - Grh.LastCount) * GrhData(Grh.GrhIndex).Speed)
        Grh.LastCount = timeGetTime
        
        'If the frame counter is higher then the number of frames...
        If Grh.FrameCounter >= GrhData(Grh.GrhIndex).NumFrames + 1 Then
        
            'Loop the animation
            If LoopAnim Then
                Do While Grh.FrameCounter >= GrhData(Grh.GrhIndex).NumFrames + 1
                    Grh.FrameCounter = Grh.FrameCounter - GrhData(Grh.GrhIndex).NumFrames
                Loop
            
            'Looping isn't set, just kill the animation
            Else
                Grh.Started = 0
                Exit Function
            End If
            
        End If
        
    End If
    
    'The grpahic will be rendered
    Engine_UpdateGrh = True
    
End Function

Sub Engine_Render_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, Optional ByVal LoopAnim As Boolean = True, Optional ByVal Light1 As Long = -1, Optional ByVal Light2 As Long = -1, Optional ByVal Light3 As Long = -1, Optional ByVal Light4 As Long = -1, Optional ByVal Shadow As Byte = 0, Optional ByVal Angle As Single = 0)

'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
Dim CurrGrhIndex As Long    'The grh index we will be working with (acquired after updating animations)
Dim FileNum As Integer

    'Check to make sure it is legal
    If Grh.GrhIndex < 1 Then Exit Sub
    If GrhData(Grh.GrhIndex).NumFrames < 1 Then Exit Sub
    If Int(Grh.FrameCounter) > GrhData(Grh.GrhIndex).NumFrames Then Grh.FrameCounter = 1
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrGrhIndex = GrhData(Grh.GrhIndex).Frames(Int(Grh.FrameCounter))

    'Check for in-bounds
    If X + GrhData(CurrGrhIndex).pixelWidth > 0 Then
        If Y + GrhData(CurrGrhIndex).pixelHeight > 0 Then
            If X < frmMain.ScaleWidth Then
                If Y < frmMain.ScaleHeight Then
                
                    'Update the animation frame
                    If Animate Then
                        If Not Engine_UpdateGrh(Grh, LoopAnim) Then Exit Sub
                    End If
                    
                    'Set the file number in a shorter variable
                    FileNum = GrhData(CurrGrhIndex).FileNum
                
                    'Center Grh over X,Y pos
                    If Center Then
                        If GrhData(CurrGrhIndex).TileWidth > 1 Then
                            X = X - GrhData(CurrGrhIndex).TileWidth * TilePixelWidth \ 2 + TilePixelWidth \ 2
                        End If
                        If GrhData(CurrGrhIndex).TileHeight > 1 Then
                            Y = Y - GrhData(CurrGrhIndex).TileHeight * TilePixelHeight + TilePixelHeight
                        End If
                    End If
                
                    'Check the rendering method to use
                    'If AlternateRender = 0 Then
                      
                        'Render the texture with 2 triangles on a triangle strip
                        Engine_Render_Rectangle X, Y, GrhData(CurrGrhIndex).pixelWidth, GrhData(CurrGrhIndex).pixelHeight, GrhData(CurrGrhIndex).sX, _
                            GrhData(CurrGrhIndex).sY, GrhData(CurrGrhIndex).pixelWidth, GrhData(CurrGrhIndex).pixelHeight, , , Angle, FileNum, Light1, Light2, Light3, Light4, Shadow
                        
                    'Else
                        
                        'Render the texture as a D3DXSprite
                    '    Engine_Render_D3DXSprite X, Y, GrhData(CurrGrhIndex).pixelWidth, GrhData(CurrGrhIndex).pixelHeight, GrhData(CurrGrhIndex).sX, GrhData(CurrGrhIndex).sY, Light1, FileNum, Angle
                        
                    'End If
                    
                End If
            End If
        End If
    End If

End Sub

Sub Engine_Render_FullTexture(ByVal hWnd As Long, ByVal TextureNum As Integer)

'************************************************************
'Does whatever the hell I want it to! >:D
'************************************************************
Dim VertexArray(0 To 3) As TLVERTEX
Dim SrcBitmapWidth As Long
Dim SrcBitmapHeight As Long
    
    'Load the surface into memory if it is not in memory and reset the timer
    If TextureNum > 0 Then
        If SurfaceTimer(TextureNum) = 0 Then Engine_Init_Texture TextureNum
        SurfaceTimer(TextureNum) = SurfaceTimerMax
    End If

    'Set the texture
    If LastTexture <> TextureNum Then
        If TextureNum <= 0 Then
            D3DDevice.SetTexture 0, Nothing
        Else
            D3DDevice.SetTexture 0, SurfaceDB(TextureNum)
        End If
        LastTexture = TextureNum
    End If

    'Set the bitmap dimensions if needed
    SrcBitmapWidth = SurfaceSize(TextureNum).X
    SrcBitmapHeight = SurfaceSize(TextureNum).Y
    
    'Set the top-left corner
    With VertexArray(0)
        .Color = -1
        .Tu = 0
        .Tv = 0
        .X = 0
        .Y = 0
        .Rhw = 1
    End With

    'Set the top-right corner
    With VertexArray(1)
        .Color = -1
        .Tu = 1
        .Tv = 0
        .X = SrcBitmapWidth
        .Y = 0
        .Rhw = 1
    End With

    'Set the bottom-left corner
    With VertexArray(2)
        .X = 0
        .Y = SrcBitmapHeight
        .Color = -1
        .Tu = 0
        .Tv = 1
        .Rhw = 1
    End With

    'Set the bottom-right corner
    With VertexArray(3)
        .X = SrcBitmapWidth
        .Y = SrcBitmapHeight
        .Color = -1
        .Tu = 1
        .Tv = 1
        .Rhw = 1
    End With

    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size

End Sub

Private Sub Engine_ReadyTexture(ByVal TextureNum As Long)

'************************************************************
'Gets a texture ready to for usage
'************************************************************

    'Load the surface into memory if it is not in memory and reset the timer
    If TextureNum > 0 Then
        If SurfaceTimer(TextureNum) = 0 Then Engine_Init_Texture TextureNum
        SurfaceTimer(TextureNum) = SurfaceTimerMax
    End If

    'Set the texture
    If LastTexture <> TextureNum Then
        If TextureNum <= 0 Then
            D3DDevice.SetTexture 0, Nothing
        Else
            D3DDevice.SetTexture 0, SurfaceDB(TextureNum)
        End If
        LastTexture = TextureNum
    End If
    
End Sub

Sub Engine_Render_Rectangle(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal SrcX As Single, ByVal SrcY As Single, ByVal SrcWidth As Single, ByVal SrcHeight As Single, Optional ByVal SrcBitmapWidth As Long = -1, Optional ByVal SrcBitmapHeight As Long = -1, Optional ByVal Degrees As Single = 0, Optional ByVal TextureNum As Long, Optional ByVal Color0 As Long = -1, Optional ByVal Color1 As Long = -1, Optional ByVal Color2 As Long = -1, Optional ByVal Color3 As Long = -1, Optional ByVal Shadow As Byte = 0, Optional ByVal GrhIndex As Long = 0, Optional ByVal InBoundsCheck As Boolean = True)

'************************************************************
'Render a square/rectangle based on the specified values then rotate it if needed
'************************************************************
Dim VertexArray(0 To 3) As TLVERTEX
Dim RadAngle As Single 'The angle in Radians
Dim CenterX As Single
Dim CenterY As Single
Dim Index As Integer
Dim NewX As Single
Dim NewY As Single
Dim SinRad As Single
Dim CosRad As Single
Dim ShadowAdd As Single

    'Perform in-bounds check if needed
    If InBoundsCheck Then
        If X + SrcWidth <= 0 Then Exit Sub
        If Y + SrcHeight <= 0 Then Exit Sub
        If X >= ScreenWidth Then Exit Sub
        If Y >= ScreenHeight Then Exit Sub
    End If

    'Ready the texture
    Engine_ReadyTexture TextureNum

    'Set the bitmap dimensions if needed
    If SrcBitmapWidth = -1 Then SrcBitmapWidth = SurfaceSize(TextureNum).X
    If SrcBitmapHeight = -1 Then SrcBitmapHeight = SurfaceSize(TextureNum).Y
    
    'Set the RHWs (must always be 1)
    VertexArray(0).Rhw = 1
    VertexArray(1).Rhw = 1
    VertexArray(2).Rhw = 1
    VertexArray(3).Rhw = 1
    
    'Apply the colors
    VertexArray(0).Color = Color0
    VertexArray(1).Color = Color1
    VertexArray(2).Color = Color2
    VertexArray(3).Color = Color3

    If Shadow Then

        'To make things easy, we just do a completely separate calculation the top two points
        ' with an uncropped tU / tV algorithm
        VertexArray(0).X = X + (Width * 0.5)
        VertexArray(0).Y = Y - (Height * 0.5)
        VertexArray(0).Tu = (SrcX / SrcBitmapWidth)
        VertexArray(0).Tv = (SrcY / SrcBitmapHeight)
        
        VertexArray(1).X = VertexArray(0).X + Width
        VertexArray(1).Tu = ((SrcX + Width) / SrcBitmapWidth)

        VertexArray(2).X = X
        VertexArray(2).Tu = (SrcX / SrcBitmapWidth)

        VertexArray(3).X = X + Width
        VertexArray(3).Tu = (SrcX + SrcWidth + ShadowAdd) / SrcBitmapWidth

    Else
        
        'If we are NOT using shadows, then we add +1 to the width/height (trust me, just do it... :p)
        ShadowAdd = 1

        'Find the left side of the rectangle
        VertexArray(0).X = X
        VertexArray(0).Tu = (SrcX / SrcBitmapWidth)

        'Find the top side of the rectangle
        VertexArray(0).Y = Y
        VertexArray(0).Tv = (SrcY / SrcBitmapHeight)
    
        'Find the right side of the rectangle
        VertexArray(1).X = X + Width
        VertexArray(1).Tu = (SrcX + SrcWidth + ShadowAdd) / SrcBitmapWidth

        'These values will only equal each other when not a shadow
        VertexArray(2).X = VertexArray(0).X
        VertexArray(3).X = VertexArray(1).X

    End If
    
    'Find the bottom of the rectangle
    VertexArray(2).Y = Y + Height
    VertexArray(2).Tv = (SrcY + SrcHeight + ShadowAdd) / SrcBitmapHeight

    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(1).Y = VertexArray(0).Y
    VertexArray(1).Tv = VertexArray(0).Tv
    VertexArray(2).Tu = VertexArray(0).Tu
    VertexArray(3).Y = VertexArray(2).Y
    VertexArray(3).Tu = VertexArray(1).Tu
    VertexArray(3).Tv = VertexArray(2).Tv
    
    'Check if a rotation is required
    If Degrees <> 0 Or Degrees <> 360 Then

        'Converts the angle to rotate by into radians
        RadAngle = Degrees * DegreeToRadian

        'Set the CenterX and CenterY values
        CenterX = X + (Width * 0.5)
        CenterY = Y + (Height * 0.5)

        'Pre-calculate the cosine and sine of the radiant
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)

        'Loops through the passed vertex buffer
        For Index = 0 To 3

            'Calculates the new X and Y co-ordinates of the vertices for the given angle around the center co-ordinates
            NewX = CenterX + (VertexArray(Index).X - CenterX) * CosRad - (VertexArray(Index).Y - CenterY) * SinRad
            NewY = CenterY + (VertexArray(Index).Y - CenterY) * CosRad + (VertexArray(Index).X - CenterX) * SinRad

            'Applies the new co-ordinates to the buffer
            VertexArray(Index).X = NewX
            VertexArray(Index).Y = NewY

        Next Index

    End If

    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size

End Sub

Public Sub Engine_BuildMiniMap()

'***************************************************
'Builds the array for the minimap. Theres multiple styles available, but only one
'is used in the demo, so experiment with them and check which one you like!
'***************************************************
Const UseOption As Byte = 2 'Change to the type of map you want
Dim MMC_Blocked As Long
Dim MMC_Exit As Long
Dim MMC_Sign As Long
Dim X As Byte
Dim Y As Byte
Dim j As Byte

    'Create the colors (character colors are defined in Engine_RenderScreen when it is rendered)
    MMC_Blocked = D3DColorARGB(75, 255, 255, 255)   'Blocked tiles
    MMC_Exit = D3DColorARGB(150, 255, 0, 0)         'Exit tiles (warps)
    MMC_Sign = D3DColorARGB(125, 255, 255, 0)       'Tiles with a sign
    
    'Clear the old array by resizing to the largest array we can possibly use
    ReDim MiniMapTile(1 To CLng(MapInfo.Width) * CLng(MapInfo.Height)) As MiniMapTile
    NumMiniMapTiles = 0
    
    Select Case UseOption
        
        '***** Option 1 *****
        Case 1

            For Y = 1 To MapInfo.Height
                For X = 1 To MapInfo.Width
                    
                    'Check for signs
                    If MapData(X, Y).Sign > 1 Then
                        NumMiniMapTiles = NumMiniMapTiles + 1
                        MiniMapTile(NumMiniMapTiles).X = X
                        MiniMapTile(NumMiniMapTiles).Y = Y
                        MiniMapTile(NumMiniMapTiles).Color = MMC_Sign
                    Else
                    
                        'Check for exits
                        If MapData(X, Y).TileExit.Map > 0 Then
                            NumMiniMapTiles = NumMiniMapTiles + 1
                            MiniMapTile(NumMiniMapTiles).X = X
                            MiniMapTile(NumMiniMapTiles).Y = Y
                            MiniMapTile(NumMiniMapTiles).Color = MMC_Exit
                        Else
                            
                            'Check for blocked tiles
                            If MapData(X, Y).Blocked = 0 Then
                                NumMiniMapTiles = NumMiniMapTiles + 1
                                MiniMapTile(NumMiniMapTiles).X = X
                                MiniMapTile(NumMiniMapTiles).Y = Y
                                MiniMapTile(NumMiniMapTiles).Color = MMC_Blocked
                            End If
                        End If
                    End If
                    
                Next X
            Next Y
                
        '***** Option 2 *****
        Case 2

            For Y = 1 To MapInfo.Height
                j = 0   'Clear the row settings
                For X = 1 To MapInfo.Width
                    
                    'Check if there is a sign
                    If MapData(X, Y).Sign > 1 Then
                        NumMiniMapTiles = NumMiniMapTiles + 1
                        MiniMapTile(NumMiniMapTiles).X = X
                        MiniMapTile(NumMiniMapTiles).Y = Y
                        MiniMapTile(NumMiniMapTiles).Color = MMC_Sign
                    Else
                    
                        'Check if there is an exit
                        If MapData(X, Y).TileExit.Map > 0 Then
                            NumMiniMapTiles = NumMiniMapTiles + 1
                            MiniMapTile(NumMiniMapTiles).X = X
                            MiniMapTile(NumMiniMapTiles).Y = Y
                            MiniMapTile(NumMiniMapTiles).Color = MMC_Exit
                        Else
                            
                            'Only check blocked tiles
                            If MapData(X, Y).Blocked > 0 Then
        
                                'If the row is set to draw, just keep drawing
                                If j = 1 Then
                                    NumMiniMapTiles = NumMiniMapTiles + 1
                                    MiniMapTile(NumMiniMapTiles).X = X
                                    MiniMapTile(NumMiniMapTiles).Y = Y
                                    MiniMapTile(NumMiniMapTiles).Color = MMC_Blocked
                                    
                                'The row isn't drawing, check if it is time to draw it
                                Else
        
                                    'If the next tile is not blocked, this one will be (to draw an outline)
                                    If j = 0 Then
                                        If X + 1 <= MapInfo.Width Then
                                            If MapData(X + 1, Y).Blocked = 0 Then
                                                NumMiniMapTiles = NumMiniMapTiles + 1
                                                MiniMapTile(NumMiniMapTiles).X = X
                                                MiniMapTile(NumMiniMapTiles).Y = Y
                                                MiniMapTile(NumMiniMapTiles).Color = MMC_Blocked
                                                j = 1
                                            End If
                                        End If
                                    End If
                                    
                                    'If the tile above or below is blocked, draw the tile
                                    If j = 0 Then
                                        If Y > 1 Then
                                            If MapData(X, Y - 1).Blocked = 0 Then
                                                NumMiniMapTiles = NumMiniMapTiles + 1
                                                MiniMapTile(NumMiniMapTiles).X = X
                                                MiniMapTile(NumMiniMapTiles).Y = Y
                                                MiniMapTile(NumMiniMapTiles).Color = MMC_Blocked
                                                j = 1
                                            End If
                                        End If
                                    End If
                                    If j = 0 Then
                                        If Y < MapInfo.Height Then
                                            If MapData(X, Y + 1).Blocked = 0 Then
                                                NumMiniMapTiles = NumMiniMapTiles + 1
                                                MiniMapTile(NumMiniMapTiles).X = X
                                                MiniMapTile(NumMiniMapTiles).Y = Y
                                                MiniMapTile(NumMiniMapTiles).Color = MMC_Blocked
                                                j = 1
                                            End If
                                        End If
                                    End If
                                    
                                    'If we STILL haven't drawn the tile, check to the diagonals (this makes corners smoothed)
                                    If j = 0 Then
                                        If Y > 1 Then
                                            If Y < MapInfo.Height Then
                                                If X > 1 Then
                                                    If X < MapInfo.Width Then
                                                        If MapData(X - 1, Y - 1).Blocked = 0 Or MapData(X - 1, Y + 1).Blocked = 0 Or MapData(X + 1, Y - 1).Blocked = 0 Or MapData(X + 1, Y + 1).Blocked = 0 Then
                                                            NumMiniMapTiles = NumMiniMapTiles + 1
                                                            MiniMapTile(NumMiniMapTiles).X = X
                                                            MiniMapTile(NumMiniMapTiles).Y = Y
                                                            MiniMapTile(NumMiniMapTiles).Color = MMC_Blocked
                                                            j = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    
                                End If
                                
                                'If the next tile isn't blocked, we remove the row drawing
                                If j = 1 Then
                                    If X < MapInfo.Width Then
                                        If MapData(X + 1, Y).Blocked > 0 Then j = 0
                                    End If
                                End If
                                
                            End If
                        End If
                    End If
                Next X
            Next Y

    End Select
    
    'Resize the array to fit the amount of data we have
    If NumMiniMapTiles = 0 Then
        Erase MiniMapTile
    Else
        ReDim Preserve MiniMapTile(1 To NumMiniMapTiles)
    End If

End Sub

Function Engine_ValidateDevice() As Boolean

'***********************************************
'Makes sure the device settings are valid
'***********************************************
Dim j As Long
Dim DispMode As D3DDISPLAYMODE          'Describes the display mode
Dim i As Byte
    
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then
    
        On Error GoTo ErrOut
        
        'Do a loop while device is lost
        If D3DDevice.TestCooperativeLevel = D3DERR_DEVICELOST Then
            Engine_ValidateDevice = False
            Exit Function
        End If
        
        'Clear all the textures
        LastTexture = -999
        For j = 1 To NumGrhFiles
            Set SurfaceDB(j) = Nothing
            SurfaceTimer(j) = 0
            SurfaceSize(j).X = 0
            SurfaceSize(j).Y = 0
        Next j

        D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
        D3DWindow.Windowed = 1
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
        D3DWindow.BackBufferFormat = DispMode.Format
        Set D3DDevice = Nothing
        Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmScreen.hWnd, UsedCreateFlags, D3DWindow)

        'Reset the render states
        Engine_Init_RenderStates
        
        Engine_Init_ParticleEngine True

        On Error GoTo 0

    End If
    
    'Everything is fine
    Engine_ValidateDevice = True
    
    Exit Function
    
ErrOut:

    Engine_ValidateDevice = False
        
End Function

Sub Engine_Render_Screen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

'***********************************************
'Draw current visible to scratch area based on TileX and TileY
'***********************************************
Dim ScreenX As Integer 'Keeps track of where to place tile on screen
Dim ScreenY As Integer
Dim ChrID() As Integer
Dim ChrY() As Integer
Dim Grh As Grh
Dim x2 As Long
Dim y2 As Long
Dim Y As Long    'Keeps track of where on map we are
Dim X As Long
Dim j As Long
Dim Layer As Byte

    minXOffset = 0
    minYOffset = 0

    'Check if we need to update the graphics
    If TileX <> LastTileX Or TileY <> LastTileY Then
    
        'Figure out Ends and Starts of screen
        ScreenMinY = TileY - (WindowTileHeight \ 2)
        ScreenMaxY = TileY + (WindowTileHeight \ 2)
        ScreenMinX = TileX - (WindowTileWidth \ 2)
        ScreenMaxX = TileX + (WindowTileWidth \ 2)
        minY = ScreenMinY - TileBufferSize
        maxY = ScreenMaxY + TileBufferSize
        minX = ScreenMinX - TileBufferSize
        maxX = ScreenMaxX + TileBufferSize
        
        'Update the last position
        LastTileX = TileX
        LastTileY = TileY
        
        'Re-create the tile layers
        Engine_CreateTileLayers
        
    End If

    'Calculate the particle offset values
    'Do NOT move this any farther down in the module or you will get "jumps" as the left/top borders on particles
    ParticleOffsetX = (Engine_PixelPosX(ScreenMinX) - PixelOffsetX) * 1
    ParticleOffsetY = (Engine_PixelPosY(ScreenMinY) - PixelOffsetY) * 1
    
    'Check if we have the device
    If Not Engine_ValidateDevice Then Exit Sub
    
    Engine_EndScreenRender
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    D3DDevice.BeginScene
    
    DrawingGameScreen = True
    
    '************** Layer 1 to 3 **************
    
    'Loop through the lower 3 layers
    For Layer = 1 To 3
        
        'Loop through all the tiles we know we will draw for this layer
        For j = 1 To TileLayer(Layer).NumTiles
            With TileLayer(Layer).Tile(j)
                
                'Check if we have to draw with a shadow or not (slighty changes because we have to animate on the shadow, not the main render)
                If MapData(.TileX, .TileY).Shadow(Layer) = 1 Then
                    Engine_Render_Grh MapData(.TileX, .TileY).Graphic(Layer), .PixelPosX + PixelOffsetX, .PixelPosY + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
                    Engine_Render_Grh MapData(.TileX, .TileY).Graphic(Layer), .PixelPosX + PixelOffsetX, .PixelPosY + PixelOffsetY, 0, 0, True, MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 1), MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 2), MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 3), MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 4)
                Else
                    Engine_Render_Grh MapData(.TileX, .TileY).Graphic(Layer), .PixelPosX + PixelOffsetX, .PixelPosY + PixelOffsetY, 0, 1, True, MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 1), MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 2), MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 3), MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 4)
                End If
                
            End With
        Next j
        
        'Tile preview
        If frmSetTile.Visible Then
            If DrawLayer = Layer Then
                Grh.FrameCounter = 1
                Grh.GrhIndex = Val(frmSetTile.GrhTxt.Text)
                j = D3DColorARGB(200, 255, 255, 255)
                If frmSetTile.ShadowChk.Value = 1 Then
                    If Val(frmSetTile.ShadowTxt.Text) = 1 Then Engine_Render_Grh Grh, Engine_PixelPosX(minXOffset + (HovertX - minX)) + (32 * (10 - TileBufferSize)) + PixelOffsetX, Engine_PixelPosY(minYOffset + (HovertY - minY)) + (32 * (10 - TileBufferSize)) + PixelOffsetY, 0, 0, False, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
                End If
                Engine_Render_Grh Grh, Engine_PixelPosX(minXOffset + (HovertX - minX)) + (32 * (10 - TileBufferSize)) + PixelOffsetX, Engine_PixelPosY(minYOffset + (HovertY - minY)) + (32 * (10 - TileBufferSize)) + PixelOffsetY, 0, 0, False, j, j, j, j, 0
            End If
        End If
        
    Next Layer

    '************** Objects **************
    'We don't need to show objects in the map editor...
    'For j = 1 To LastObj
    '    If OBJList(j).Grh.GrhIndex Then
    '        X = Engine_PixelPosX(minXOffset + (OBJList(j).Pos.X - minX)) + PixelOffsetX
    '        Y = Engine_PixelPosY(minYOffset + (OBJList(j).Pos.Y - minY)) + PixelOffsetY
    '        If Y >= -32 Then
    '            If Y <= 632 Then
    '                If X >= -32 Then
    '                    If X <= 832 Then
    '                        X2 = minXOffset + (OBJList(j).Pos.X - minX)
    '                        Y2 = minYOffset + (OBJList(j).Pos.Y - minY)
    '                        Engine_Render_Grh OBJList(j).Grh, X, Y, 1, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 0, 1
    '                        Engine_Render_Grh OBJList(j).Grh, X, Y, 1, 0, True, MapData(X2, Y2).Light(1), MapData(X2, Y2).Light(2), MapData(X2, Y2).Light(3), MapData(X2, Y2).Light(4)
    '                    End If
    '                End If
    '            End If
    '        End If
    '    End If
    'Next j

    '************** Characters **************
    If CharsChkValue = 1 Then
        'Size the to the smallest safe size (LastChar)
        If LastChar > 0 Then
            ReDim ChrID(1 To LastChar)
            ReDim ChrY(1 To LastChar)
        
            'Fill the array
            For j = 1 To LastChar
                ChrY(j) = CharList(j).Pos.Y
                ChrID(j) = j
            Next j
        
            'Sort the char list
            Engine_SortIntArray ChrY, ChrID, 1, LastChar
        
            'Loop through the sorted characters
            For j = 1 To LastChar
                If CharList(ChrID(j)).Active Then
                    X = Engine_PixelPosX(minXOffset + (CharList(ChrID(j)).Pos.X - minX)) + PixelOffsetX + ((10 - TileBufferSize) * 32)
                    Y = Engine_PixelPosY(minYOffset + (CharList(ChrID(j)).Pos.Y - minY)) + PixelOffsetY + ((10 - TileBufferSize) * 32)
                    If Y >= -32 Then
                        If Y <= 632 Then
                            If X >= -32 Then
                                If X <= 832 Then
                                    Engine_Render_Char ChrID(j), X, Y
                                End If
                            End If
                        End If
                    End If
                End If
            Next j
        End If
    End If

    '************** Layer 4 to 6 **************
    For Layer = 4 To 6
        For j = 1 To TileLayer(Layer).NumTiles
            With TileLayer(Layer).Tile(j)
                If MapData(.TileX, .TileY).Shadow(Layer) = 1 Then
                    Engine_Render_Grh MapData(.TileX, .TileY).Graphic(Layer), .PixelPosX + PixelOffsetX, .PixelPosY + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
                    Engine_Render_Grh MapData(.TileX, .TileY).Graphic(Layer), .PixelPosX + PixelOffsetX, .PixelPosY + PixelOffsetY, 0, 0, True, MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 1), MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 2), MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 3), MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 4)
                Else
                    Engine_Render_Grh MapData(.TileX, .TileY).Graphic(Layer), .PixelPosX + PixelOffsetX, .PixelPosY + PixelOffsetY, 0, 1, True, MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 1), MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 2), MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 3), MapData(.TileX, .TileY).Light(((Layer - 1) * 4) + 4)
                End If
            End With
        Next j
        
        'Tile preview
        If frmSetTile.Visible Then
            If DrawLayer = Layer Then
                Grh.FrameCounter = 1
                Grh.GrhIndex = Val(frmSetTile.GrhTxt.Text)
                j = D3DColorARGB(200, 255, 255, 255)
                If frmSetTile.ShadowChk.Value = 1 Then
                    If Val(frmSetTile.ShadowTxt.Text) = 1 Then Engine_Render_Grh Grh, Engine_PixelPosX(minXOffset + (HovertX - minX)) + (32 * (10 - TileBufferSize)) + PixelOffsetX, Engine_PixelPosY(minYOffset + (HovertY - minY)) + (32 * (10 - TileBufferSize)) + PixelOffsetY, 0, 0, False, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
                End If
                Engine_Render_Grh Grh, Engine_PixelPosX(minXOffset + (HovertX - minX)) + (32 * (10 - TileBufferSize)) + PixelOffsetX, Engine_PixelPosY(minYOffset + (HovertY - minY)) + (32 * (10 - TileBufferSize)) + PixelOffsetY, 0, 0, False, j, j, j, j, 0
            End If
        End If
        
    Next Layer
    
    '************** Info **************
    If InfoChkValue = 1 Then
        For X = 1 To InfoLayer.NumTiles
            For Y = 1 To InfoLayer.Tile(X).NumGrhs
                With InfoLayer.Tile(X).Grh(Y)
                    Engine_Render_Grh .Grh, .PixelPosX + PixelOffsetX, .PixelPosY + PixelOffsetY, 0, 0, False
                End With
            Next Y
        Next X
    End If
    
    '************** Grid **************
    If GridChkValue = 1 Then
        j = D3DColorARGB(25, 255, 255, 255)
        Grh.GrhIndex = 2
        Grh.FrameCounter = 1
        Grh.Started = 0
        ScreenY = minYOffset
        For Y = minY To maxY
            ScreenX = minXOffset
            For X = minX To maxX
                Engine_Render_Grh Grh, Engine_PixelPosX(ScreenX) + PixelOffsetX + ((10 - TileBufferSize) * 32), Engine_PixelPosY(ScreenY) + PixelOffsetY + ((10 - TileBufferSize) * 32), 0, 0, , j, j, j, j
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If

    '************** Update weather **************
    If WeatherChkValue = 1 Then
    
        'Make sure the right weather is going on
        Engine_Weather_Update
    
        'Update the weather
        If WeatherEffectIndex Then
            If ParticleOffsetX <> 0 Then
                If ParticleOffsetY <> 0 Then
                    Effect(WeatherEffectIndex).ShiftX = (LastOffsetX - ParticleOffsetX)
                    Effect(WeatherEffectIndex).ShiftY = (LastOffsetY - ParticleOffsetY)
                End If
            End If
        End If
        
    Else
    
        If WeatherEffectIndex > 0 Then
            If Effect(WeatherEffectIndex).Used Then Effect_Kill WeatherEffectIndex
        End If
        
        'Change the light of all the tiles back
        If FlashTimer > 0 Then
            For X = 1 To MapInfo.Width
                For Y = 1 To MapInfo.Height
                    For x2 = 1 To 4
                        MapData(X, Y).Light(x2) = SaveLightBuffer(X, Y).Light(x2)
                    Next x2
                Next Y
            Next X
            FlashTimer = 0
        End If
        
    End If
    
    '************** Misc Rendering **************

    'Update and render particle effects
    Effect_UpdateAll

    'Clear the shift-related variables
    LastOffsetX = ParticleOffsetX
    LastOffsetY = ParticleOffsetY
    
    '************** Mini-map **************
    Const tS As Single = 2  'Size of the mini-map dots
    
    If ShowMiniMap Then
    
        'Draw the map outline
        For X = 1 To NumMiniMapTiles
            Engine_Render_Rectangle MiniMapTile(X).X * tS, MiniMapTile(X).Y * tS, tS, tS, 1, 1, 1, 1, 1, 1, 0, 0, MiniMapTile(X).Color, MiniMapTile(X).Color, MiniMapTile(X).Color, MiniMapTile(X).Color
        Next X
        
        'Draw the characters
        j = D3DColorARGB(200, 0, 255, 255)
        For X = 1 To LastChar
            Engine_Render_Rectangle CharList(X).Pos.X * tS, CharList(X).Pos.Y * tS, tS, tS, 1, 1, 1, 1, 1, 1, 0, 0, j, j, j, j
        Next X
        
        'Draw the position indicator
        j = D3DColorARGB(200, 0, 255, 0)
        Engine_Render_Rectangle UserPos.X * tS, UserPos.Y * tS, tS, tS, 1, 1, 1, 1, 1, 1, 0, 0, j, j, j, j
        
    End If

End Sub

Public Sub Engine_EndScreenRender()

    If DrawingGameScreen Then
        
        If Not Engine_ValidateDevice Then Exit Sub
    
        D3DDevice.EndScene
        
        'Display the textures drawn to the device
        D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
        
        DrawingGameScreen = False
        
    End If

End Sub

Public Sub Engine_CreateTileLayers()

'************************************************************
'Creates the tile layers used for rendering the tiles so they can be drawn faster
'Has to happen every time the user warps or moves a whole tile
'************************************************************
Dim Layer As Byte
Dim ScreenX As Byte
Dim ScreenY As Byte
Dim tBuf As Integer
Dim pX As Integer
Dim pY As Integer
Dim X As Long
Dim Y As Long
Dim j As Long

    'Raise the buffer up + 1 to prevent graphical errors
    tBuf = TileBufferSize '+ 1
    
    'Loop through each layer and check which tiles there are that will need to be drawn
    For Layer = 1 To 6
        
        'Clear the number of tiles
        TileLayer(Layer).NumTiles = 0
        
        'Allocate enough memory for all the tiles
        ReDim TileLayer(Layer).Tile(1 To ((maxY - minY + 1) * (maxX - minX + 1)))
        
        'Loop through all the tiles within the buffer's range
        ScreenY = (10 - tBuf)
        For Y = minY To maxY
            ScreenX = (10 - tBuf)
            For X = minX To maxX
            
                'Check that the tile is in the range of the map
                If X >= 1 Then
                    If Y >= 1 Then
                        If X <= MapInfo.Width Then
                            If Y <= MapInfo.Height Then
                        
                                'Check if the tile even has a graphic on it
                                If MapData(X, Y).Graphic(Layer).GrhIndex Then
                                
                                    'Calculate the pixel values
                                    pX = Engine_PixelPosX(ScreenX) - 288
                                    pY = Engine_PixelPosY(ScreenY) - 288
                                    
                                    'Check that the tile is in the screen
                                    With GrhData(MapData(X, Y).Graphic(Layer).GrhIndex)
                                        If pX >= -.pixelWidth Then
                                            If pX <= ScreenWidth + .pixelWidth Then
                                                If pY >= -.pixelHeight Then
                                                    If pY <= ScreenHeight + .pixelHeight Then
                                                        
                                                        'The tile is valid to be used, so raise the count
                                                        TileLayer(Layer).NumTiles = TileLayer(Layer).NumTiles + 1
                                                        
                                                        'Store the needed information
                                                        TileLayer(Layer).Tile(TileLayer(Layer).NumTiles).TileX = X
                                                        TileLayer(Layer).Tile(TileLayer(Layer).NumTiles).TileY = Y
                                                        TileLayer(Layer).Tile(TileLayer(Layer).NumTiles).PixelPosX = pX + 288
                                                        TileLayer(Layer).Tile(TileLayer(Layer).NumTiles).PixelPosY = pY + 288
    
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End With
    
                                End If
                                
                            End If
                        End If
                    End If
                End If
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    
        'We got all the information we need, now resize the array as small as possible to save RAM, then do the same for every other layer :o
        If TileLayer(Layer).NumTiles > 0 Then
            ReDim Preserve TileLayer(Layer).Tile(1 To TileLayer(Layer).NumTiles)
        Else
            Erase TileLayer(Layer).Tile
        End If
        
    Next Layer
    
    '***** Information Layer *****
    'This is just like the layers above, but holds information a bit differently
    InfoLayer.NumTiles = 1
    ReDim InfoLayer.Tile(1 To ((maxY - minY + 1) * (maxX - minX + 1)))
    
    ScreenY = (10 - tBuf)
    For Y = minY To maxY
        ScreenX = (10 - tBuf)
        For X = minX To maxX
        
            If X >= 1 Then
                If Y >= 1 Then
                    If X <= MapInfo.Width Then
                        If Y <= MapInfo.Height Then

                            pX = Engine_PixelPosX(ScreenX)
                            pY = Engine_PixelPosY(ScreenY)
                            
                            If pX - 288 > -32 Then
                                If pY - 288 < ScreenWidth + 32 Then
                                    If pY - 288 > -32 Then
                                        If pY - 288 < ScreenHeight + 32 Then
                
                                            'Blocked tiles
                                            If MapData(X, Y).Blocked And 1 Then 'North
                                                InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs = InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs + 1
                                                ReDim Preserve InfoLayer.Tile(InfoLayer.NumTiles).Grh(1 To InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs)
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).Grh.GrhIndex = 650
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosX = pX
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosY = pY
                                            End If
                                            If MapData(X, Y).Blocked And 2 Then 'East
                                                InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs = InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs + 1
                                                ReDim Preserve InfoLayer.Tile(InfoLayer.NumTiles).Grh(1 To InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs)
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).Grh.GrhIndex = 651
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosX = pX
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosY = pY
                                            End If
                                            If MapData(X, Y).Blocked And 4 Then 'South
                                                InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs = InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs + 1
                                                ReDim Preserve InfoLayer.Tile(InfoLayer.NumTiles).Grh(1 To InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs)
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).Grh.GrhIndex = 652
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosX = pX
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosY = pY
                                            End If
                                            If MapData(X, Y).Blocked And 8 Then 'West
                                                InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs = InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs + 1
                                                ReDim Preserve InfoLayer.Tile(InfoLayer.NumTiles).Grh(1 To InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs)
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).Grh.GrhIndex = 653
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosX = pX
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosY = pY
                                            End If
                                            
                                            'No-attack tiles
                                            If Not (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
                                                If MapData(X, Y).BlockedAttack Then
                                                    InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs = InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs + 1
                                                    ReDim Preserve InfoLayer.Tile(InfoLayer.NumTiles).Grh(1 To InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs)
                                                    InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).Grh.GrhIndex = 10
                                                    InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosX = pX
                                                    InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosY = pY
                                                End If
                                            End If
                                            
                                            'Warp Tiles
                                            If MapData(X, Y).TileExit.X <> 0 Then
                                                InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs = InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs + 1
                                                ReDim Preserve InfoLayer.Tile(InfoLayer.NumTiles).Grh(1 To InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs)
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).Grh.GrhIndex = 65
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosX = pX + 8
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosY = pY + 2
                                            End If
                                            
                                            'Mailbox Tiles
                                            If MapData(X, Y).Mailbox > 0 Then
                                                InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs = InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs + 1
                                                ReDim Preserve InfoLayer.Tile(InfoLayer.NumTiles).Grh(1 To InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs)
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).Grh.GrhIndex = 66
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosX = pX + 14
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosY = pY + 2
                                            End If
                                            
                                            'Sfx Tiles
                                            If MapData(X, Y).Sfx > 0 Then
                                                InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs = InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs + 1
                                                ReDim Preserve InfoLayer.Tile(InfoLayer.NumTiles).Grh(1 To InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs)
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).Grh.GrhIndex = 655
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosX = pX + 20
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosY = pY + 2
                                            End If
                                            
                                            'Sign tiles
                                            If MapData(X, Y).Sign > 0 Then
                                                InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs = InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs + 1
                                                ReDim Preserve InfoLayer.Tile(InfoLayer.NumTiles).Grh(1 To InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs)
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).Grh.GrhIndex = 13
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosX = pX + 24
                                                InfoLayer.Tile(InfoLayer.NumTiles).Grh(InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs).PixelPosY = pY + 2
                                            End If
                                            
                                            If InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs > 0 Then
                                            
                                                'Assign the frame to 1
                                                For j = 1 To InfoLayer.Tile(InfoLayer.NumTiles).NumGrhs
                                                    InfoLayer.Tile(InfoLayer.NumTiles).Grh(j).Grh.FrameCounter = 1
                                                Next j
                                            
                                                'Raise the tiles count
                                                InfoLayer.NumTiles = InfoLayer.NumTiles + 1
                                                ReDim Preserve InfoLayer.Tile(1 To InfoLayer.NumTiles)
                                                
                                            End If
                                            
                                        End If
                                    End If
                                End If
                            End If
                            
                        End If
                    End If
                End If
            End If
            
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    If InfoLayer.NumTiles > 0 Then
        ReDim Preserve InfoLayer.Tile(1 To InfoLayer.NumTiles)
    Else
        Erase InfoLayer.Tile
    End If

End Sub

Public Sub Engine_Render_Text(ByVal Text As String, ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Color As Long, Optional ByVal Format As Long = DT_LEFT Or DT_TOP)

'************************************************************
'Draw text on D3DDevice
'************************************************************

Dim Tempstr() As String
Dim TxtRect As RECT
Dim i As Long

'Check for valid text to render

    If LenB(Text) = 0 Then Exit Sub

    'Set up the text rectangle
    TxtRect.Left = X
    TxtRect.Right = X + Width

    'Get the text into arrays (split by vbCrLf)
    Tempstr = Split(Text, vbCrLf)

    'Check for valid areas to draw the text
    If TxtRect.Left < 0 Then
        TxtRect.Left = 0
        TxtRect.Right = Width
        Format = DT_LEFT
    ElseIf TxtRect.Left > 800 Then
        TxtRect.Left = 800 - Width
        TxtRect.Right = 800
        Format = DT_RIGHT
    End If
    If TxtRect.Top < 0 Then
        TxtRect.Top = 0
        TxtRect.bottom = lngTextHeight * UBound(Tempstr)
        Format = DT_TOP
    ElseIf TxtRect.bottom > 600 Then
        TxtRect.Top = 600 - lngTextHeight * UBound(Tempstr)
        TxtRect.bottom = 600
        Format = DT_BOTTOM
    End If

    'Alphablending must be disabled for text to display properly
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False

    'Draw the text
    For i = 0 To UBound(Tempstr)
        TxtRect.Top = Y + (i * lngTextHeight)
        TxtRect.bottom = TxtRect.Top + lngTextHeight
        D3DX.DrawText MainFont, Color, Tempstr(i), TxtRect, Format
    Next i

    'Re-enable alphablending
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True

End Sub

Sub Engine_ShowNextFrame()

'***********************************************
'Updates and draws next frame to screen
'***********************************************
'***** Check if engine is allowed to run ******

    If EngineRun Then
        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * TickPerFrame * (1 + -(GetAsyncKeyState(vbKeyShift) <> 0) * 2)
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * TickPerFrame * (1 + -(GetAsyncKeyState(vbKeyShift) <> 0) * 2)
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
        
        '****** Update preview *****
        If UpdatePreview Then
            Engine_Render_PreviewScreens
        End If

        '****** Update screen ******
        Call Engine_Render_Screen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX - 288, OffsetCounterY - 288)

        'Get timing info
        ElapsedTime = Engine_ElapsedTime()
        TickPerFrame = (ElapsedTime * EngineBaseSpeed)
        TimerMultiplier = TickPerFrame * 0.075
        If FPS_Last_Check + 1000 < timeGetTime Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            FPS_Last_Check = timeGetTime
            'Show FPS
            frmMain.FPSLbl.Caption = "FPS: " & FPS
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
    End If
    
    Engine_EndScreenRender

End Sub

Public Sub Engine_Render_PreviewScreens()

'*****************************************************************
'Sort an array of integers
'*****************************************************************
Dim i As Long
Dim R As RECT
Dim j As Long

    'This timer brought to you by a lazy ass programmer :)
    'Screw speed, this is a tool!

    If D3DDevice Is Nothing Then Exit Sub
    If Not Engine_ValidateDevice Then Exit Sub
    If SearchTextureFileNum > 0 Then
    
        Engine_EndScreenRender
    
        If ShownTextureGrhs.NumGrhs > 0 Then
            
            D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
            D3DDevice.BeginScene
            
            If LastThingy <> SearchTextureFileNum Then
                frmSearchTexture.Refresh
                LastThingy = SearchTextureFileNum
            End If
            Engine_Render_FullTexture frmSearchTexture.hWnd, SearchTextureFileNum
            
            j = D3DColorARGB(150, 0, 255, 0)
            For i = 1 To ShownTextureGrhs.NumGrhs
                With ShownTextureGrhs.Grh(i)
                    Engine_Render_Rectangle .X, .Y, 1, .Height, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                    Engine_Render_Rectangle .X, .Y, .Width, 1, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                    Engine_Render_Rectangle .X + .Width, .Y, 1, .Height, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                    Engine_Render_Rectangle .X, .Y + .Height, .Width, 1, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                End With
            Next i
            
            D3DDevice.EndScene
            
            R.Left = 0
            R.Top = 0
            R.Right = SurfaceSize(SearchTextureFileNum).X
            R.bottom = SurfaceSize(SearchTextureFileNum).Y
            
            D3DDevice.Present R, R, frmSearchTexture.hWnd, ByVal 0
        
        End If
        
        If ShownTextureAnims.NumGrhs > 0 Then
            
            D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
            D3DDevice.BeginScene
            
            If LastThingy <> SearchTextureFileNum Then
                frmSearchTexture.Refresh
                LastThingy = SearchTextureFileNum
            End If
            
            For i = 1 To ShownTextureAnims.NumGrhs
                With ShownTextureAnims.Grh(i)
                    Engine_Render_Grh .Grh, .X, .Y, 0, 1, True
                End With
            Next i
            
            j = D3DColorARGB(150, 0, 255, 0)
            For i = 1 To ShownTextureGrhs.NumGrhs
                With ShownTextureGrhs.Grh(i)
                    Engine_Render_Rectangle .X, .Y, 1, .Height, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                    Engine_Render_Rectangle .X, .Y, .Width, 1, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                    Engine_Render_Rectangle .X + .Width, .Y, 1, .Height, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                    Engine_Render_Rectangle .X, .Y + .Height, .Width, 1, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                End With
            Next i
            
            D3DDevice.EndScene
            
            R.Left = 0
            R.Top = 0
            R.Right = STAWidth
            R.bottom = STAHeight
            
            D3DDevice.Present R, R, frmSearchAnim.hWnd, ByVal 0
            
        End If
        
    End If
    
    DrawPreview
    DrawTileInfoPreview
    UpdatePreview = False

End Sub

Public Sub Engine_SortIntArray(TheArray() As Integer, TheIndex() As Integer, ByVal LowerBound As Integer, ByVal UpperBound As Integer)

'*****************************************************************
'Sort an array of integers
'*****************************************************************

Dim s(1 To 64) As Integer   'Stack space for pending Subarrays
Dim indxt As Long   'Stored index
Dim swp As Integer  'Swap variable
Dim F As Integer    'Subarray Minimum
Dim g As Integer    'Subarray Maximum
Dim h As Integer    'Subarray Middle
Dim i As Integer    'Subarray Low  Scan Index
Dim j As Integer    'Subarray High Scan Index
Dim t As Integer    'Stack pointer

'Set the array boundries to f and g

    F = LowerBound
    g = UpperBound

    'Start the loop
    Do

        For j = F + 1 To g
            indxt = TheIndex(j)
            swp = TheArray(indxt)
            For i = j - 1 To F Step -1
                If TheArray(TheIndex(i)) <= swp Then Exit For
                TheIndex(i + 1) = TheIndex(i)
            Next i
            TheIndex(i + 1) = indxt
        Next j

        'Finished sorting when t = 0
        If t = 0 Then Exit Do

        'Pop stack and begin new partitioning round
        g = s(t)
        F = s(t - 1)
        t = t - 2

    Loop

End Sub

Sub Engine_UnloadAllForms()

'*****************************************************************
'Unloads all forms
'*****************************************************************

Dim frm As Form

    For Each frm In VB.Forms
        If frm.Name <> frmMain.Name Then Unload frm
    Next

End Sub

Function Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
Dim sSpaces As String

    sSpaces = Space$(500)
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    Var_Get = Trim$(sSpaces)
    Var_Get = Left$(Var_Get, Len(Var_Get) - 1)

End Function

Sub Var_Write(File As String, Main As String, Var As String, Value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    writeprivateprofilestring Main, Var, Value, File

End Sub
