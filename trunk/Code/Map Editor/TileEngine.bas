Attribute VB_Name = "TileEngine"
Option Explicit

Public Const ShadowColor As Long = 1677721600  'ARGB 100/0/0/0
Public Const PreviewColor As Long = -1258291201 'ARGB 180/255/255/255

Public ParticleOffsetX As Long
Public ParticleOffsetY As Long
Public LastOffsetX As Integer           'The last offset values stored, used to get the offset difference
Public LastOffsetY As Integer           ' so the particle engine can adjust weather particles accordingly

Public lngTextHeight As Long

Public DevMode As Byte

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

'Map sizes in tiles
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

'********** WEATHER ***********
Public Type LightType
    Light(1 To 24) As Long
End Type
Public SaveLightBuffer(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As LightType
Public WeatherEffectIndex As Long   'Index returned by the weather effect initialization
Public DoLightning As Byte          'Are we using lightning? 1 = Yes, 0 = No
Public LightningTimer As Single     'How long until our next lightning bolt strikes
Public FlashTimer As Single         'How long until the flash goes away (being > 0 states flash is happening)

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
    Graphic(1 To 6) As Grh      'Index of the 4 graphic layers
    Light(1 To 24) As Long      'Holds the light values - retrieve with Index = Light * Layer
    UserIndex As Integer        'Index of the user on the tile
    NPCIndex As Integer         'Index of the NPC on the tile
    ObjInfo As OBJ              'Information of the object on the tile
    TileExit As WorldPosEX      'Warp location when user touches the tile
    Mailbox As Byte             'If there is a mailbox on the tile
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

'Describes the damage counters
Public Type DamageTxt
    Pos As FloatPos
    Value As String
    Counter As Single
    Width As Integer
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
Public UserCharIndex As Integer
Public EngineRun As Boolean
Public FPS As Long
Private FramesPerSecCounter As Long
Private FPS_Last_Check As Long

'Main view size size in tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

'How many tiles the engine "looks ahead" when drawing the screen
Private TileBufferSize As Integer

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

'********** Direct X ***********
Public Const SurfaceTimerMax As Single = 30000  'How long a texture stays in memory unused (miliseconds)
Public SurfaceDB() As Direct3DTexture8          'The list of all the textures
Public SurfaceTimer() As Integer                'How long until the surface unloads
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
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    Rhw As Single
    Color As Long
    Specular As Long
    Tu As Single
    Tv As Single
End Type

Private VertexArray(0 To 3) As TLVERTEX

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
Public PreviewMapGrh(1 To 6) As Grh

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

Private Function Engine_AllFilesInFolders(ByVal sFolderPath As String, Optional bWithSubFolders As Boolean = False) As String()

'************************************************************
'Returns a list of all the files in a folder
'************************************************************

Dim sTemp As String
Dim sDirIn As String
Dim i As Integer
Dim j As Integer

'Clear the arrays

    ReDim sFilelist(0) As String
    ReDim sSubFolderList(0) As String
    ReDim sToProcessFolderList(0) As String

    'Set the initial directory
    sDirIn = sFolderPath

    'Make sure we have a slash
    If Not (Right$(sDirIn, 1) = "\") Then sDirIn = sDirIn & "\"

    'Resume on errors - we can handle them ourselves
    On Error Resume Next

        'Loop through the files in the targeted folder
        sTemp = Dir$(sDirIn & "*.*")
        Do While sTemp <> ""
            Engine_AddItem2Array1D sFilelist(), sDirIn & sTemp
            sTemp = Dir
        Loop

        'Loop through the files in the sub folders to the targeted folder
        If bWithSubFolders Then

            'Loop through the subdirectories
            sTemp = Dir$(sDirIn & "*.*", vbDirectory)
            Do While sTemp <> ""
                If sTemp <> "." And sTemp <> ".." Then
                    If (GetAttr(sDirIn & sTemp) And vbDirectory) = vbDirectory Then Engine_AddItem2Array1D sToProcessFolderList, sDirIn & sTemp
                End If
                sTemp = Dir
            Loop
            If UBound(sToProcessFolderList) > 0 Or UBound(sToProcessFolderList) = 0 And sToProcessFolderList(0) <> "" Then
                For i = 0 To UBound(sToProcessFolderList)
                    sSubFolderList = Engine_AllFilesInFolders(sToProcessFolderList(i), bWithSubFolders)
                    If UBound(sSubFolderList) > 0 Or UBound(sSubFolderList) = 0 And sSubFolderList(0) <> "" Then
                        For j = 0 To UBound(sSubFolderList)
                            Engine_AddItem2Array1D sFilelist(), sSubFolderList(j)
                        Next
                    End If
                Next
            End If

        End If

        'Return the values
        Engine_AllFilesInFolders = sFilelist

End Function

Sub Engine_Char_Erase(ByVal CharIndex As Integer)

'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************

    'Check for valid position
    If CharList(CharIndex).Pos.X <= XMinMapSize Then Exit Sub
    If CharList(CharIndex).Pos.X >= XMaxMapSize Then Exit Sub
    If CharList(CharIndex).Pos.Y <= YMinMapSize Then Exit Sub
    If CharList(CharIndex).Pos.Y >= YMaxMapSize Then Exit Sub

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

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

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

    'Get next end time
    EndTime = Start_Time

End Function

Function Engine_FileExist(file As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

    Engine_FileExist = (Dir$(file, FileType) <> "")

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

    GetTextExtentPoint32 frmMain.hdc, Text, Len(Text), Engine_GetTextSize

End Function

Sub Engine_Init_BodyData()

'*****************************************************************
'Loads Body.dat
'*****************************************************************
Dim LoopC As Long
Dim j As Long
'Get number of bodies

    NumBodies = CInt(Engine_Var_Get(DataPath & "Body.dat", "INIT", "NumBodies"))
    'Resize array
    ReDim BodyData(0 To NumBodies) As BodyData
    'Fill list
    For LoopC = 1 To NumBodies
        For j = 1 To 8
            Engine_Init_Grh BodyData(LoopC).Walk(j), CInt(Engine_Var_Get(DataPath & "Body.dat", Str$(LoopC), Str$(j))), 0
            Engine_Init_Grh BodyData(LoopC).Attack(j), CInt(Engine_Var_Get(DataPath & "Body.dat", Str$(LoopC), "a" & j)), 1
        Next j
        BodyData(LoopC).HeadOffset.X = CLng(Engine_Var_Get(DataPath & "Body.dat", Str$(LoopC), "HeadOffsetX"))
        BodyData(LoopC).HeadOffset.Y = CLng(Engine_Var_Get(DataPath & "Body.dat", Str$(LoopC), "HeadOffsetY"))
    Next LoopC

End Sub

Sub Engine_Init_WingData()

'*****************************************************************
'Loads Wing.dat
'*****************************************************************
Dim LoopC As Long
Dim j As Long

    'Get number of wings
    NumWings = CInt(Engine_Var_Get(DataPath & "Wing.dat", "INIT", "NumWings"))
    
    'Resize array
    ReDim WingData(0 To NumWings) As WingData
    
    'Fill list
    For LoopC = 1 To NumWings
        For j = 1 To 8
            Engine_Init_Grh WingData(LoopC).Walk(j), CInt(Engine_Var_Get(DataPath & "Wing.dat", Str(LoopC), Str(j))), 0
            Engine_Init_Grh WingData(LoopC).Attack(j), CInt(Engine_Var_Get(DataPath & "Wing.dat", Str(LoopC), "a" & j)), 1
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
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.ScreenPic.hwnd, D3DCREATEFLAGS, D3DWindow)

    'Store the create flags
    UsedCreateFlags = D3DCREATEFLAGS

    'Everything was successful
    Engine_Init_D3DDevice = 1

    'The Rhw will always be 1, so set it now instead of every call
    For i = 0 To 3
        VertexArray(i).Rhw = 1
    Next i

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

    If GrhIndex <= 0 Then Exit Sub
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
Dim Grh As Long
Dim Frame As Long

    'Get Number of Graphics
    GrhPath = App.Path & "\Grh\"
    NumGrhs = CLng(Engine_Var_Get(DataPath & "Grh.ini", "INIT", "NumGrhs"))
    
    'Resize arrays
    ReDim GrhData(1 To NumGrhs) As GrhData
    
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
                If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > NumGrhs Then
                    GoTo ErrorHandler
                End If
            Next Frame
            
            Get #1, , GrhData(Grh).Speed
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

    NumHairs = CInt(Engine_Var_Get(DataPath & "Hair.dat", "INIT", "NumHairs"))
    'Resize array
    ReDim HairData(0 To NumHairs) As HairData
    'Fill List
    For LoopC = 1 To NumHairs
        For i = 1 To 8
            Engine_Init_Grh HairData(LoopC).Hair(i), CInt(Engine_Var_Get(DataPath & "Hair.dat", Str$(LoopC), Str$(i))), 0
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

    NumHeads = CInt(Engine_Var_Get(DataPath & "Head.dat", "INIT", "NumHeads"))
    'Resize array
    ReDim HeadData(0 To NumHeads) As HeadData
    'Fill List
    For LoopC = 1 To NumHeads
        For i = 1 To 8
            Engine_Init_Grh HeadData(LoopC).Head(i), CInt(Engine_Var_Get(DataPath & "Head.dat", Str$(LoopC), Str(i))), 0
            Engine_Init_Grh HeadData(LoopC).Blink(i), CInt(Engine_Var_Get(DataPath & "Head.dat", Str$(LoopC), "b" & i)), 0
            Engine_Init_Grh HeadData(LoopC).AgrHead(i), CInt(Engine_Var_Get(DataPath & "Head.dat", Str$(LoopC), "a" & i)), 0
            Engine_Init_Grh HeadData(LoopC).AgrBlink(i), CInt(Engine_Var_Get(DataPath & "Head.dat", Str$(LoopC), "ab" & i)), 0
        Next i
    Next LoopC

End Sub

Sub Engine_Init_MapData()

'*****************************************************************
'Load Map.dat
'*****************************************************************
'Get Number of Maps

    NumMaps = CInt(Engine_Var_Get(DataPath & "Map.dat", "INIT", "NumMaps"))

End Sub

Sub Engine_Init_ParticleEngine()

'*****************************************************************
'Loads all particles into memory - unlike normal textures, these stay in memory. This isn't
'done for any reason in particular, they just use so little memory since they are so small
'*****************************************************************

Dim i As Byte

    'Set the particles texture
    NumEffects = Engine_Var_Get(DataPath & "Game.ini", "INIT", "NumEffects")
    ReDim Effect(1 To NumEffects)
    
    'Update effects list
    UpdateEffectList

    For i = 1 To UBound(ParticleTexture())
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
        Set SurfaceDB(TextureNum) = D3DX.CreateTextureFromFileEx(D3DDevice, FilePath, D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, TexInfo, ByVal 0)

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
    SfxPath = Engine_Var_Get(DataPath & "Game.ini", "INIT", "SoundPath")

    'Fill startup variables
    DisplayFormhWnd = setDisplayFormhWnd
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    EngineBaseSpeed = Engine_Speed

    'Setup borders
    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder

    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = 36
    ScrollPixelsPerFrameY = 36

    'Set the array sizes by the number of graphic files
    NumGrhFiles = CInt(Engine_Var_Get(DataPath & "Grh.ini", "INIT", "NumGrhFiles"))
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

End Function

Public Sub Engine_Init_UnloadTileEngine()

'*****************************************************************
'Shutsdown engine
'*****************************************************************

    On Error Resume Next

    Dim LoopC As Long

        EngineRun = False

        '****** Clear DirectX objects ******
        Set D3DDevice = Nothing
        Set MainFont = Nothing
        Set D3DX = Nothing

        'Clear GRH memory
        For LoopC = 1 To NumGrhFiles
            Set SurfaceDB(LoopC) = Nothing
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

    NumWeapons = CInt(Engine_Var_Get(DataPath & "Weapon.dat", "INIT", "NumWeapons"))
    'Resize array
    ReDim WeaponData(0 To NumWeapons) As WeaponData
    'Fill listn
    For LoopC = 1 To NumWeapons
        Engine_Init_Grh WeaponData(LoopC).Walk(1), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk1")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(2), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk2")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(3), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk3")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(4), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk4")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(5), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk5")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(6), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk6")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(7), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk7")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(8), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Walk8")), 0
        Engine_Init_Grh WeaponData(LoopC).Attack(1), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack1")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(2), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack2")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(3), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack3")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(4), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack4")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(5), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack5")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(6), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack6")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(7), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack7")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(8), CInt(Engine_Var_Get(DataPath & "Weapon.dat", "Weapon" & LoopC, "Attack8")), 1
    Next LoopC

End Sub

Sub Engine_Init_Weather()

'*****************************************************************
'Initializes the weather effects
'*****************************************************************
Dim TempGrh As Grh
Dim X As Byte
Dim Y As Byte
Dim i As Byte

    Select Case MapInfo.Weather
    Case 0  'None
        If WeatherEffectIndex > 0 Then
            If Effect(WeatherEffectIndex).Used Then Effect_Kill WeatherEffectIndex
        End If
        
    Case 1  'Snow (light fall)
        If WeatherEffectIndex <= 0 Then
            WeatherEffectIndex = Effect_Snow_Begin(1, 400)
        ElseIf Effect(WeatherEffectIndex).EffectNum <> EffectNum_Snow Then
            Effect_Kill WeatherEffectIndex
            WeatherEffectIndex = Effect_Snow_Begin(1, 400)
        ElseIf Effect(WeatherEffectIndex).Used = False Then
            WeatherEffectIndex = Effect_Snow_Begin(1, 400)
        End If
        DoLightning = 0
        
    Case 2  'Rain Storm (heavy rain + lightning)
        If WeatherEffectIndex <= 0 Then
            WeatherEffectIndex = Effect_Rain_Begin(9, 400)
        ElseIf Effect(WeatherEffectIndex).EffectNum <> EffectNum_Rain Then
            Effect_Kill WeatherEffectIndex
            WeatherEffectIndex = Effect_Rain_Begin(9, 400)
        ElseIf Effect(WeatherEffectIndex).Used = False Then
            WeatherEffectIndex = Effect_Rain_Begin(9, 400)
        End If
        DoLightning = 1 'We take our rain with a bit of lightning on top >:D
        
    End Select
    
    'Update lightning
    If DoLightning Then
        
        'Check if we are in the middle of a flash
        If FlashTimer > 0 Then
            FlashTimer = FlashTimer - ElapsedTime
            
            'The flash has run out
            If FlashTimer <= 0 Then
            
                'Change the light of all the tiles back
                For X = XMinMapSize To XMaxMapSize
                    For Y = YMinMapSize To YMaxMapSize
                        For i = 1 To 4
                            MapData(X, Y).Light(i) = SaveLightBuffer(X, Y).Light(i)
                        Next i
                    Next Y
                Next X
            
            End If
            
        'Update the timer, see if it is time to flash
        Else
            LightningTimer = LightningTimer - ElapsedTime
            
            'Flash me, baby!
            If LightningTimer <= 0 Then
                LightningTimer = 15000 + (Rnd * 15000)  'Reset timer (flash every 15 to 30 seconds)
                FlashTimer = 250    'How long the flash is (miliseconds)

                'Change the light of all the tiles to white
                For X = XMinMapSize To XMaxMapSize
                    For Y = YMinMapSize To YMaxMapSize
                        For i = 1 To 4
                            MapData(X, Y).Light(i) = -1
                        Next i
                    Next Y
                Next X
                
            End If
            
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
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = True
    End If

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

Function Engine_RectCollision(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal Width1 As Integer, ByVal Height1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal Width2 As Integer, ByVal Height2 As Integer)

'******************************************
'Check for collision between two rectangles
'******************************************

Dim RetRect As RECT
Dim Rect1 As RECT
Dim Rect2 As RECT

'Build the rectangles

    Rect1.Left = X1
    Rect1.Right = X1 + Width1
    Rect1.Top = Y1
    Rect1.bottom = Y1 + Height1
    Rect2.Left = X2
    Rect2.Right = X2 + Width2
    Rect2.Top = Y2
    Rect2.bottom = Y2 + Height2

    'Call collision API
    Engine_RectCollision = IntersectRect(RetRect, Rect1, Rect2)

End Function

Public Sub Engine_SetTileSelectionArray()

'***************************************************
'Create the tile selection array, starting at tsStart and skipping unused graphics
'***************************************************
Dim CurrentGrh As Long
Dim i As Long

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
            
            'Set the graphic
            Engine_Init_Grh PreviewGrhList(i), CurrentGrh
            
            'Check if we hit the end of the list
            i = i + 1
            If i > UBound(PreviewGrhList) Then Exit Do
            
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
                D3DDevice.Present src, dest, frmTileSelect.hwnd, ByVal 0
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
    If CharList(CharIndex).Pos.Y > YMinMapSize Then TempBlock2 = MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y - 1)

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

Sub Engine_Render_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, Optional ByVal LoopAnim As Boolean = True, Optional ByVal Light1 As Long = -1, Optional ByVal Light2 As Long = -1, Optional ByVal Light3 As Long = -1, Optional ByVal Light4 As Long = -1, Optional ByVal Degrees As Byte = 0, Optional ByVal Shadow As Byte = 0)

'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************

Dim CurrentGrh As Grh

'Check to make sure it is legal

    If Grh.GrhIndex < 1 Then Exit Sub
    If GrhData(Grh.GrhIndex).NumFrames < 1 Then Exit Sub

    'Update the animation frame
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (TimerMultiplier * GrhData(Grh.GrhIndex).Speed)
            If Grh.FrameCounter >= GrhData(Grh.GrhIndex).NumFrames + 1 Then
                Do While Grh.FrameCounter >= GrhData(Grh.GrhIndex).NumFrames + 1
                    Grh.FrameCounter = Grh.FrameCounter - GrhData(Grh.GrhIndex).NumFrames
                Loop
                If LoopAnim <> True Then Grh.Started = 0
            End If
        End If
    End If

    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Int(Grh.FrameCounter))

    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * TilePixelWidth * 0.5) + TilePixelWidth * 0.5
        End If
        If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If

    'Draw
    If X + GrhData(CurrentGrh.GrhIndex).pixelWidth > 0 Then
        If Y + GrhData(CurrentGrh.GrhIndex).pixelHeight > 0 Then
            If X < frmMain.ScaleWidth Then
                If Y < frmMain.ScaleHeight Then
                    Engine_Render_Rectangle X, Y, GrhData(CurrentGrh.GrhIndex).pixelWidth, GrhData(CurrentGrh.GrhIndex).pixelHeight, GrhData(CurrentGrh.GrhIndex).sX, GrhData(CurrentGrh.GrhIndex).sY, GrhData(CurrentGrh.GrhIndex).pixelWidth, GrhData(CurrentGrh.GrhIndex).pixelHeight, , , 0, GrhData(CurrentGrh.GrhIndex).FileNum, Light1, Light2, Light3, Light4, Shadow
                End If
            End If
        End If
    End If

End Sub

Sub Engine_Render_Rectangle(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal SrcX As Single, ByVal SrcY As Single, ByVal SrcWidth As Single, ByVal SrcHeight As Single, Optional ByVal SrcBitmapWidth As Long = -1, Optional ByVal SrcBitmapHeight As Long = -1, Optional ByVal Degrees As Single = 0, Optional ByVal TextureNum As Long, Optional ByVal Color0 As Long = -1, Optional ByVal Color1 As Long = -1, Optional ByVal Color2 As Long = -1, Optional ByVal Color3 As Long = -1, Optional ByVal Shadow As Byte = 0)

'************************************************************
'Render a square/rectangle based on the specified values then rotate it if needed
'************************************************************

Dim RadAngle As Single 'The angle in Radians
Dim CenterX As Single
Dim CenterY As Single
Dim Index As Integer
Dim NewX As Single
Dim NewY As Single
Dim SinRad As Single
Dim CosRad As Single

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
    If SrcBitmapWidth = -1 Then SrcBitmapWidth = SurfaceSize(TextureNum).X
    If SrcBitmapHeight = -1 Then SrcBitmapHeight = SurfaceSize(TextureNum).Y

    'Set shadowed settings - shadows only change on the top 2 points
    If Shadow Then

        SrcWidth = SrcWidth - 1

        'Set the top-left corner
        With VertexArray(0)
            .X = X + (Width * 0.5)
            .Y = Y - (Height * 0.5)
        End With

        'Set the top-right corner
        With VertexArray(1)
            .X = X + Width + (Width * 0.5)
            .Y = Y - (Height * 0.5)
        End With

    Else

        SrcWidth = SrcWidth + 1
        SrcHeight = SrcHeight + 1

        'Set the top-left corner
        With VertexArray(0)
            .X = X
            .Y = Y
        End With

        'Set the top-right corner
        With VertexArray(1)
            .X = X + Width
            .Y = Y
        End With

    End If

    'Set the top-left corner
    With VertexArray(0)
        .Color = Color0
        .Tu = SrcX / SrcBitmapWidth
        .Tv = SrcY / SrcBitmapHeight
    End With

    'Set the top-right corner
    With VertexArray(1)
        .Color = Color1
        .Tu = (SrcX + SrcWidth) / SrcBitmapWidth
        .Tv = SrcY / SrcBitmapHeight
    End With

    'Set the bottom-left corner
    With VertexArray(2)
        .X = X
        .Y = Y + Height
        .Color = Color2
        .Tu = SrcX / SrcBitmapWidth
        .Tv = (SrcY + SrcHeight) / SrcBitmapHeight
    End With

    'Set the bottom-right corner
    With VertexArray(3)
        .X = X + Width
        .Y = Y + Height
        .Color = Color3
        .Tu = (SrcX + SrcWidth) / SrcBitmapWidth
        .Tv = (SrcY + SrcHeight) / SrcBitmapHeight
    End With

    'Check if a rotation is required
    If Degrees <> 0 Then

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
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), Len(VertexArray(0))

End Sub

Sub Engine_Render_Screen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

'***********************************************
'Draw current visible to scratch area based on TileX and TileY
'***********************************************
Dim ScreenX As Integer 'Keeps track of where to place tile on screen
Dim ScreenY As Integer
Dim ChrID() As Integer
Dim ChrY() As Integer
Dim Grh As Grh 'Temp Grh for show tile and blocked
Dim X2 As Long
Dim Y2 As Long
Dim Y As Long    'Keeps track of where on map we are
Dim X As Long
Dim j As Long

    minXOffset = 0
    minYOffset = 0

    'Figure out Ends and Starts of screen
    ScreenMinY = TileY - WindowTileHeight \ 2
    ScreenMaxY = TileY + WindowTileHeight \ 2
    ScreenMinX = TileX - WindowTileWidth \ 2
    ScreenMaxX = TileX + WindowTileWidth \ 2
    minY = ScreenMinY - TileBufferSize
    maxY = ScreenMaxY + TileBufferSize
    minX = ScreenMinX - TileBufferSize
    maxX = ScreenMaxX + TileBufferSize

    'Calculate the particle offset values
    'Do NOT move this any farther down in the module or you will get "jumps" as the left/top borders on particles
    ParticleOffsetX = (Engine_PixelPosX(ScreenMinX) - PixelOffsetX) * 1
    ParticleOffsetY = (Engine_PixelPosY(ScreenMinY) - PixelOffsetY) * 1

    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    If maxY > YMaxMapSize Then
        maxY = YMaxMapSize
    End If
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    If maxX > XMaxMapSize Then
        maxX = XMaxMapSize
    End If

    'If we can we render around the view area to make it smoother
    If ScreenMinY > YMinMapSize Then
        ScreenMinY = ScreenMinY - 1
    Else
        ScreenMinY = 1
        ScreenY = 1
    End If
    If ScreenMaxY < YMaxMapSize Then
        ScreenMaxY = ScreenMaxY + 1
    End If
    If ScreenMinX > XMinMapSize Then
        ScreenMinX = ScreenMinX - 1
    Else
        ScreenMinX = 1
        ScreenX = 1
    End If
    If ScreenMaxX < XMaxMapSize Then
        ScreenMaxX = ScreenMaxX + 1
    End If

    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    D3DDevice.BeginScene

    '************** Layer 1 **************
    If L1ChkValue = 1 Then
        For Y = ScreenMinY To ScreenMaxY
            For X = ScreenMinX To ScreenMaxX
                
                'Map preview
                If frmSetTile.Visible = True Then
                    If PreviewMapGrh(1).GrhIndex > 0 Then
                        If X = HovertX Then
                            If Y = HovertY Then
                                If frmSetTile.LayerChk(1).Value = 1 Then
                                    If frmSetTile.ShadowChk(1).Value = 1 Then
                                        Engine_Render_Grh PreviewMapGrh(1), Engine_PixelPosX(ScreenX) + PixelOffsetX + ((TileBufferSize - 1) * TilePixelWidth), Engine_PixelPosY(ScreenY) + PixelOffsetY + ((TileBufferSize - 1) * TilePixelHeight), 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 0, 1
                                        Engine_Render_Grh PreviewMapGrh(1), Engine_PixelPosX(ScreenX) + PixelOffsetX + ((TileBufferSize - 1) * TilePixelWidth), Engine_PixelPosY(ScreenY) + PixelOffsetY + ((TileBufferSize - 1) * TilePixelHeight), 0, 0, True, PreviewColor, PreviewColor, PreviewColor, PreviewColor
                                    Else
                                        Engine_Render_Grh PreviewMapGrh(1), Engine_PixelPosX(ScreenX) + PixelOffsetX + ((TileBufferSize - 1) * TilePixelWidth), Engine_PixelPosY(ScreenY) + PixelOffsetY + ((TileBufferSize - 1) * TilePixelHeight), 0, 1, True, PreviewColor, PreviewColor, PreviewColor, PreviewColor
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                If MapData(X, Y).Shadow(1) = 1 Then
                    Engine_Render_Grh MapData(X, Y).Graphic(1), Engine_PixelPosX(ScreenX) + PixelOffsetX + ((TileBufferSize - 1) * TilePixelWidth), Engine_PixelPosY(ScreenY) + PixelOffsetY + ((TileBufferSize - 1) * TilePixelHeight), 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 0, 1
                    Engine_Render_Grh MapData(X, Y).Graphic(1), Engine_PixelPosX(ScreenX) + PixelOffsetX + ((TileBufferSize - 1) * TilePixelWidth), Engine_PixelPosY(ScreenY) + PixelOffsetY + ((TileBufferSize - 1) * TilePixelHeight), 0, 0, True, MapData(X, Y).Light(1), MapData(X, Y).Light(2), MapData(X, Y).Light(3), MapData(X, Y).Light(4)
                Else
                    Engine_Render_Grh MapData(X, Y).Graphic(1), Engine_PixelPosX(ScreenX) + PixelOffsetX + ((TileBufferSize - 1) * TilePixelWidth), Engine_PixelPosY(ScreenY) + PixelOffsetY + ((TileBufferSize - 1) * TilePixelHeight), 0, 1, True, MapData(X, Y).Light(1), MapData(X, Y).Light(2), MapData(X, Y).Light(3), MapData(X, Y).Light(4)
                End If
                ScreenX = ScreenX + 1
            Next X
            ScreenX = ScreenX - X + ScreenMinX
            ScreenY = ScreenY + 1
        Next Y
    End If
    
    '************** Layer 2 **************
    If L2ChkValue = 1 Then
        ScreenY = minYOffset
        For Y = minY To maxY
            ScreenX = minXOffset
            For X = minX To maxX
            
                'Map preview
                If frmSetTile.Visible = True Then
                    If PreviewMapGrh(2).GrhIndex > 0 Then
                        If X = HovertX Then
                            If Y = HovertY Then
                                If frmSetTile.LayerChk(2).Value = 1 Then
                                    If frmSetTile.ShadowChk(2).Value = 1 Then
                                        Engine_Render_Grh PreviewMapGrh(2), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 0, 1
                                        Engine_Render_Grh PreviewMapGrh(2), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0, True, PreviewColor, PreviewColor, PreviewColor, PreviewColor
                                    Else
                                        Engine_Render_Grh PreviewMapGrh(2), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, PreviewColor, PreviewColor, PreviewColor, PreviewColor
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            
                If MapData(X, Y).Graphic(2).GrhIndex Then
                    If MapData(X, Y).Shadow(2) = 1 Then
                        Engine_Render_Grh MapData(X, Y).Graphic(2), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 0, 1
                        Engine_Render_Grh MapData(X, Y).Graphic(2), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0, True, MapData(X, Y).Light(5), MapData(X, Y).Light(6), MapData(X, Y).Light(7), MapData(X, Y).Light(8)
                    Else
                        Engine_Render_Grh MapData(X, Y).Graphic(2), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, MapData(X, Y).Light(5), MapData(X, Y).Light(6), MapData(X, Y).Light(7), MapData(X, Y).Light(8)
                    End If
                End If
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If
    
    '************** Layer 3 **************
    If L3ChkValue = 1 Then
        ScreenY = minYOffset
        For Y = minY To maxY
            ScreenX = minXOffset
            For X = minX To maxX
            
                'Map preview
                If frmSetTile.Visible = True Then
                    If PreviewMapGrh(3).GrhIndex > 0 Then
                        If X = HovertX Then
                            If Y = HovertY Then
                                If frmSetTile.LayerChk(3).Value = 1 Then
                                    If frmSetTile.ShadowChk(3).Value = 1 Then
                                        Engine_Render_Grh PreviewMapGrh(3), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 0, 1
                                        Engine_Render_Grh PreviewMapGrh(3), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0, True, PreviewColor, PreviewColor, PreviewColor, PreviewColor
                                    Else
                                        Engine_Render_Grh PreviewMapGrh(3), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, PreviewColor, PreviewColor, PreviewColor, PreviewColor
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            
                If MapData(X, Y).Graphic(3).GrhIndex Then
                    If MapData(X, Y).Shadow(3) = 1 Then
                        Engine_Render_Grh MapData(X, Y).Graphic(3), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 0, 1
                        Engine_Render_Grh MapData(X, Y).Graphic(3), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0, True, MapData(X, Y).Light(9), MapData(X, Y).Light(10), MapData(X, Y).Light(11), MapData(X, Y).Light(12)
                    Else
                        Engine_Render_Grh MapData(X, Y).Graphic(3), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, MapData(X, Y).Light(9), MapData(X, Y).Light(10), MapData(X, Y).Light(11), MapData(X, Y).Light(12)
                    End If
                End If
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If

    '************** Objects **************
    If ObjChkValue = 1 Then
        For j = 1 To LastObj
            If OBJList(j).Grh.GrhIndex Then
                X = Engine_PixelPosX(minXOffset + (OBJList(j).Pos.X - minX)) + PixelOffsetX
                Y = Engine_PixelPosY(minYOffset + (OBJList(j).Pos.Y - minY)) + PixelOffsetY
                If Y >= -32 Then
                    If Y <= 632 Then
                        If X >= -32 Then
                            If X <= 832 Then
                                X2 = minXOffset + (OBJList(j).Pos.X - minX)
                                Y2 = minYOffset + (OBJList(j).Pos.Y - minY)
                                Engine_Render_Grh OBJList(j).Grh, X, Y, 1, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 0, 1
                                Engine_Render_Grh OBJList(j).Grh, X, Y, 1, 0, True, MapData(X2, Y2).Light(1), MapData(X2, Y2).Light(2), MapData(X2, Y2).Light(3), MapData(X2, Y2).Light(4)
                            End If
                        End If
                    End If
                End If
            End If
        Next j
    End If

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
                    X = Engine_PixelPosX(minXOffset + (CharList(ChrID(j)).Pos.X - minX)) + PixelOffsetX
                    Y = Engine_PixelPosY(minYOffset + (CharList(ChrID(j)).Pos.Y - minY)) + PixelOffsetY
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

    '************** Layer 4 **************
    If L4ChkValue = 1 Then
        ScreenY = minYOffset
        For Y = minY To maxY
            ScreenX = minXOffset
            For X = minX To maxX
            
                'Map preview
                If frmSetTile.Visible = True Then
                    If PreviewMapGrh(4).GrhIndex > 0 Then
                        If X = HovertX Then
                            If Y = HovertY Then
                                If frmSetTile.LayerChk(4).Value = 1 Then
                                    If frmSetTile.ShadowChk(4).Value = 1 Then
                                        Engine_Render_Grh PreviewMapGrh(4), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 0, 1
                                        Engine_Render_Grh PreviewMapGrh(4), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0, True, PreviewColor, PreviewColor, PreviewColor, PreviewColor
                                    Else
                                        Engine_Render_Grh PreviewMapGrh(4), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, PreviewColor, PreviewColor, PreviewColor, PreviewColor
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            
                If MapData(X, Y).Graphic(4).GrhIndex Then
                    If MapData(X, Y).Shadow(4) = 1 Then
                        Engine_Render_Grh MapData(X, Y).Graphic(4), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 0, 1
                        Engine_Render_Grh MapData(X, Y).Graphic(4), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0, True, MapData(X, Y).Light(13), MapData(X, Y).Light(14), MapData(X, Y).Light(15), MapData(X, Y).Light(16)
                    Else
                        Engine_Render_Grh MapData(X, Y).Graphic(4), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, MapData(X, Y).Light(13), MapData(X, Y).Light(14), MapData(X, Y).Light(15), MapData(X, Y).Light(16)
                    End If
                End If
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If
    
    '************** Layer 5 **************
    If L5ChkValue = 1 Then
        ScreenY = minYOffset
        For Y = minY To maxY
            ScreenX = minXOffset
            For X = minX To maxX
            
                'Map preview
                If frmSetTile.Visible = True Then
                    If PreviewMapGrh(5).GrhIndex > 0 Then
                        If X = HovertX Then
                            If Y = HovertY Then
                                If frmSetTile.LayerChk(5).Value = 1 Then
                                    If frmSetTile.ShadowChk(5).Value = 1 Then
                                        Engine_Render_Grh PreviewMapGrh(5), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 0, 1
                                        Engine_Render_Grh PreviewMapGrh(5), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0, True, PreviewColor, PreviewColor, PreviewColor, PreviewColor
                                    Else
                                        Engine_Render_Grh PreviewMapGrh(5), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, PreviewColor, PreviewColor, PreviewColor, PreviewColor
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            
                If MapData(X, Y).Graphic(5).GrhIndex Then
                    If MapData(X, Y).Shadow(5) = 1 Then
                        Engine_Render_Grh MapData(X, Y).Graphic(5), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 0, 1
                        Engine_Render_Grh MapData(X, Y).Graphic(5), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0, True, MapData(X, Y).Light(17), MapData(X, Y).Light(18), MapData(X, Y).Light(19), MapData(X, Y).Light(20)
                    Else
                        Engine_Render_Grh MapData(X, Y).Graphic(5), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, MapData(X, Y).Light(17), MapData(X, Y).Light(18), MapData(X, Y).Light(19), MapData(X, Y).Light(20)
                    End If
                End If
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If
    
    '************** Layer 6 **************
    If L6ChkValue = 1 Then
        ScreenY = minYOffset
        For Y = minY To maxY
            ScreenX = minXOffset
            For X = minX To maxX
            
                'Map preview
                If frmSetTile.Visible = True Then
                    If PreviewMapGrh(6).GrhIndex > 0 Then
                        If X = HovertX Then
                            If Y = HovertY Then
                                If frmSetTile.LayerChk(6).Value = 1 Then
                                    If frmSetTile.ShadowChk(6).Value = 1 Then
                                        Engine_Render_Grh PreviewMapGrh(6), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 0, 1
                                        Engine_Render_Grh PreviewMapGrh(6), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0, True, PreviewColor, PreviewColor, PreviewColor, PreviewColor
                                    Else
                                        Engine_Render_Grh PreviewMapGrh(6), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, PreviewColor, PreviewColor, PreviewColor, PreviewColor
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            
                If MapData(X, Y).Graphic(6).GrhIndex Then
                    If MapData(X, Y).Shadow(6) = 1 Then
                        Engine_Render_Grh MapData(X, Y).Graphic(6), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 0, 1
                        Engine_Render_Grh MapData(X, Y).Graphic(6), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0, True, MapData(X, Y).Light(21), MapData(X, Y).Light(22), MapData(X, Y).Light(23), MapData(X, Y).Light(24)
                    Else
                        Engine_Render_Grh MapData(X, Y).Graphic(6), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, MapData(X, Y).Light(21), MapData(X, Y).Light(22), MapData(X, Y).Light(23), MapData(X, Y).Light(24)
                    End If
                End If
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If
    
    '************** Grid **************
    If GridChkValue = 1 Then
        Grh.GrhIndex = 2
        Grh.FrameCounter = 1
        Grh.Started = 0
        ScreenY = minYOffset
        For Y = minY To maxY
            ScreenX = minXOffset
            For X = minX To maxX
                Engine_Render_Grh Grh, Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If
    
    '************** Info **************
    If InfoChkValue = 1 Then
        ScreenY = minYOffset
        For Y = minY To maxY
            ScreenX = minXOffset
            For X = minX To maxX
                Grh.FrameCounter = 1
                Grh.Started = 0
                'Blocked Tiles
                If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
                    Grh.GrhIndex = 654
                    Engine_Render_Grh Grh, Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0
                Else
                    If MapData(X, Y).Blocked And 1 Then 'North
                        Grh.GrhIndex = 650
                        Engine_Render_Grh Grh, Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0
                    End If
                    If MapData(X, Y).Blocked And 2 Then 'East
                        Grh.GrhIndex = 651
                        Engine_Render_Grh Grh, Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0
                    End If
                    If MapData(X, Y).Blocked And 4 Then 'South
                        Grh.GrhIndex = 652
                        Engine_Render_Grh Grh, Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0
                    End If
                    If MapData(X, Y).Blocked And 8 Then 'West
                        Grh.GrhIndex = 653
                        Engine_Render_Grh Grh, Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0
                    End If
                End If
                'Warp Tiles
                If MapData(X, Y).TileExit.X <> 0 Then
                    Grh.GrhIndex = 65
                    Engine_Render_Grh Grh, Engine_PixelPosX(ScreenX) + PixelOffsetX + 8, Engine_PixelPosY(ScreenY) + PixelOffsetY + 2, 0, 0
                End If
                'Mailbox Tiles
                If MapData(X, Y).Mailbox > 0 Then
                    Grh.GrhIndex = 66
                    Engine_Render_Grh Grh, Engine_PixelPosX(ScreenX) + PixelOffsetX + 14, Engine_PixelPosY(ScreenY) + PixelOffsetY + 2, 0, 0
                End If
                'Sfx Tiles
                If MapData(X, Y).Sfx > 0 Then
                    Grh.GrhIndex = 655
                    Engine_Render_Grh Grh, Engine_PixelPosX(ScreenX) + PixelOffsetX + 20, Engine_PixelPosY(ScreenY) + PixelOffsetY + 2, 0, 0
                End If
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If

    '************** Update weather **************
    If WeatherChkValue = 1 Then
    
        'Make sure the right weather is going on
        Engine_Init_Weather
    
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
            For X = XMinMapSize To XMaxMapSize
                For Y = YMinMapSize To YMaxMapSize
                    For X2 = 1 To 4
                        MapData(X, Y).Light(X2) = SaveLightBuffer(X, Y).Light(X2)
                    Next X2
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

    'End the device rendering
    D3DDevice.EndScene

    'Display the textures drawn to the device
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Public Sub Engine_Render_Text(ByVal Text As String, ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Color As Long, Optional ByVal Format As Long = DT_LEFT Or DT_TOP)

'************************************************************
'Draw text on D3DDevice
'************************************************************

Dim Tempstr() As String
Dim TxtRect As RECT
Dim i As Long

'Check for valid text to render

    If Text = "" Then Exit Sub

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
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * TickPerFrame
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * TickPerFrame
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
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
        If frm.Caption <> frmMain.Caption Then Unload frm
    Next
    Unload frmMain  'Unloading the main form last will allow us to unload the other forms to save their position

End Sub

Function Engine_Var_Get(file As String, Main As String, Var As String) As String

'*****************************************************************
'Gets a Var from a text file
'*****************************************************************

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

    szReturn = ""

    sSpaces = Space$(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish
    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), file
    Engine_Var_Get = RTrim$(sSpaces)
    Engine_Var_Get = Left$(Engine_Var_Get, Len(Engine_Var_Get) - 1)

End Function

Sub Engine_Var_Write(file As String, Main As String, Var As String, Value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    writeprivateprofilestring Main, Var, Value, file

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:35)  Decl: 562  Code: 4465  Total: 5027 Lines
':) CommentOnly: 753 (15%)  Commented: 113 (2.2%)  Empty: 709 (14.1%)  Max Logic Depth: 12
