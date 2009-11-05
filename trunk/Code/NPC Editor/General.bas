Attribute VB_Name = "General"
'**       ____        _________   ______   ______  ______   _______           **
'**       \   \      /   /     \ /  ____\ /      \|      \ |   ____|          **
'**        \   \    /   /|      |  /     |        |       ||  |____           **
'***        \   \  /   / |     /| |  ___ |        |      / |   ____|         ***
'****        \   \/   /  |     \| |  \  \|        |   _  \ |  |____         ****
'******       \      /   |      |  \__|  |        |  | \  \|       |      ******
'********      \____/    |_____/ \______/ \______/|__|  \__\_______|    ********
'*******************************************************************************
'*******************************************************************************
'************ vbGORE - Visual Basic 6.0 Graphical Online RPG Engine ************
'************            Official Release: Version 0.1.1            ************
'************                 http://www.vbgore.com                 ************
'*******************************************************************************
'*******************************************************************************
'***** Source Distribution Information: ****************************************
'*******************************************************************************
'** If you wish to distribute this source code, you must distribute as-is     **
'** from the vbGORE website unless permission is given to do otherwise. This  **
'** comment block must remain in-tact in the distribution. If you wish to     **
'** distribute modified versions of vbGORE, please contact Spodi (info below) **
'** before distributing the source code. You may never label the source code  **
'** as the "Official Release" or similar unless the code and content remains  **
'** unmodified from the version downloaded from the official website.         **
'** You may also never sale the source code without permission first. If you  **
'** want to sell the code, please contact Spodi (below). This is to prevent   **
'** people from ripping off other people by selling an insignificantly        **
'** modified version of open-source code just to make a few quick bucks.      **
'*******************************************************************************
'***** Creating Engines With vbGORE: *******************************************
'*******************************************************************************
'** If you plan to create an engine with vbGORE that, please contact Spodi    **
'** before doing so. You may not sell the engine unless told elsewise (the    **
'** engine must has substantial modifications), and you may not claim it as   **
'** all your own work - credit must be given to vbGORE, along with a link to  **
'** the vbGORE homepage. Failure to gain approval from Spodi directly to      **
'** make a new engine with vbGORE will result in first a friendly reminder,   **
'** followed by much more drastic measures.                                   **
'*******************************************************************************
'***** Helping Out vbGORE: *****************************************************
'*******************************************************************************
'** If you want to help out with vbGORE's progress, theres a few things you   **
'** can do:                                                                   **
'**  *Donate - Great way to keep a free project going. :) Info and benifits   **
'**        for donating can be found at:                                      **
'**        http://www.vbgore.com/modules.php?name=Content&pa=showpage&pid=11  **
'**  *Contribute - Check out our forums, contribute ideas, report bugs, or    **
'**        create tutorials for the Knowledge Base. :)                        **
'**  *Ads - Advertisements have been placed on the site for those who can     **
'**        not or do not want to donate. Not donating is understandable - not **
'**        everyone has access to credit cards / paypal or spair money laying **
'**        around. These ads allow for a free way for you to help out the     **
'**        site. Those who do donate have the option to hide/remove the ads.  **
'*******************************************************************************
'***** Conact Information: *****************************************************
'*******************************************************************************
'** Please contact the creator of vbGORE (Spodi) directly with any questions: **
'** AIM: Spodii                          Yahoo: Spodii                        **
'** MSN: Spodii@hotmail.com              Email: spodi@vbgore.com              **
'** 2nd Email: spodii@hotmail.com        Website: http://www.vbgore.com       **
'*******************************************************************************
'***** Credits: ****************************************************************
'*******************************************************************************
'** Below are credits to those who have helped with the project or who have   **
'** distributed source code which has help this project's creation. The below **
'** is listed in no particular order of significance:                         **
'**                                                                           **
'** ORE (Aaron Perkins): Used as base engine and for learning experience      **
'**   http://www.baronsoft.com/                                               **
'** SOX (Trevor Herselman): Used for all the networking                       **
'**   http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=35239&lngWId=1      **
'** Compression Methods (Marco v/d Berg): Provided compression algorithms     **
'**   http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=37867&lngWId=1      **
'** All Files In Folder (Jorge Colaccini): Algorithm implimented into engine  **
'**   http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=51435&lngWId=1      **
'** Game Programming Wiki (All community): Help on many different subjects    **
'**   http://wwww.gpwiki.org/                                                 **
'** ORE Maraxus's Edition (Maraxus): Used the map editor from this project    **
'**                                                                           **
'** Also, all the members of the vbGORE community who have submitted          **
'** tutorials, bugs, suggestions, criticism and have just stuck around!!      **
'** Big thanks goes to Van, Nex666 and ChAsE01!                               **
'**                                                                           **
'** If you feel you belong in these credits, please contact Spodi (above).    **
'*******************************************************************************
'*******************************************************************************

Option Explicit

'********** Direct X ***********
Public Const SurfaceTimerMax As Single = 30000  'How long a texture stays in memory unused (miliseconds)
Public SurfaceDB() As Direct3DTexture8          'The list of all the textures
Public SurfaceTimer() As Integer                'How long until the surface unloads
Public LastTexture As Long                      'The last texture used
Public D3DWindow As D3DPRESENT_PARAMETERS       'Describes the viewport and used to restore when in fullscreen
Public UsedCreateFlags As CONST_D3DCREATEFLAGS  'The flags we used to create the device when it first succeeded

'DirectX 8 Objects
Private DX As DirectX8
Private D3D As Direct3D8
Public D3DX As D3DX8
Public D3DDevice As Direct3DDevice8

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

Public Const NumStats As Byte = 31
Public Type StatOrder
    Gold As Byte
    EXP As Byte
    ELV As Byte
    ELU As Byte
    MaxHIT As Byte
    MinHIT As Byte
    MinMAN As Byte
    MinHP As Byte
    MinSTA As Byte
    Points As Byte
    DEF As Byte
    MaxHP As Byte
    MaxSTA As Byte
    MaxMAN As Byte
    Str As Byte
    Agil As Byte
    Mag As Byte
    Regen As Byte
    Rest As Byte
    Meditate As Byte
    Fist As Byte
    Staff As Byte
    Sword As Byte
    Parry As Byte
    Dagger As Byte
    Clairovoyance As Byte
    Immunity As Byte
    DefensiveMag As Byte
    OffensiveMag As Byte
    SummoningMag As Byte
    WeaponSkill As Byte     'Only used on NPCs
End Type
Public SID As StatOrder 'Stat ID

'Heading constants
Public Const NORTH = 1
Public Const EAST = 2
Public Const SOUTH = 3
Public Const WEST = 4
Public Const NORTHEAST = 5
Public Const SOUTHEAST = 6
Public Const SOUTHWEST = 7
Public Const NORTHWEST = 8

'Holds a position on a 2d grid
Public Type Position
    X As Long
    Y As Long
End Type

'Holds data about where a png can be found,
'How big it is and animation info
Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Byte
    Frames() As Integer
    Speed As Single
End Type

'Points to a grhData and keeps animation info
Public Type Grh
    GrhIndex As Integer
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

'Weapons list
Public Type WeaponData
    Walk(1 To 8) As Grh
    Attack(1 To 8) As Grh
    WeaponOffset As Position
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

'Holds info about a object
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
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

'Holds a world position
Public Type WorldPos
    Map As Integer  'Map
    X As Integer       'X coordinate
    Y As Integer       'Y coordinate
End Type

'User skills casted
Type Skills
    IronSkin As Byte
    Bless As Integer
    Protect As Integer
    Strengthen As Integer
    WarCurse As Integer
End Type

Type NPCFlags
    NPCAlive As Byte    'If the NPC is alive and visible
    NPCActive As Byte   'If the NPC is active
    ActionDelay As Long     'How long until the NPC can perform another action
    WalkPath() As WorldPos  'The position the NPC will be traveling
    HasPath As Byte         'If the NPC has a path they are following
    PathPos As Integer      'The index in the WalkPath() the NPC is currently on
    GoalX As Byte           'The position the NPC is trying to get to with the walkpath
    GoalY As Byte
End Type

Type NPCCounters
    RespawnCounter As Long  'Stores the death time to respawn later
    BlinkCounter As Long    'How long until the NPC blinks again
    AggressiveCounter As Long   'How long the NPC will remain aggressive-faced

    SpellExhaustion As Long     'Time until another spell can be casted
    BlessCounter As Long        'Time left on bless
    ProtectCounter As Long      'Time left on protection
    StrengthenCounter As Long   'Time left on strengthen
    WarCurseCounter As Long     'Time left on warcry-curse
End Type

Type NPC
    Name As String      'Name of the NPC
    Char As CharShort   'Defines NPC looks
    Desc As String      'Description

    Pos As WorldPos         'Current NPC Postion
    StartPos As WorldPos    'Spawning location of the NPC
    NPCNumber As Integer    'The NPC index within NPC.dat
    Movement As Integer     'Movement style
    RespawnWait As Long     'How long for the NPC to respawn
    Attackable As Byte      'If the NPC is attackable
    Hostile As Byte         'If the NPC is hostile
    GiveEXP As Long         'How much exp given on death
    GiveGLD As Long         'How much gold given on death
    Quest As Integer        'Quest index

    Skills As Skills                'Declares the skills casted on the NPC
    BaseStat(1 To NumStats) As Long 'Declares the NPC's stats
    ModStat(1 To NumStats) As Long  'Declares the NPC's stats
    Flags As NPCFlags               'Declares the NPC's flags
    Counters As NPCCounters         'Declares the NPC's counters

    NumVendItems As Integer 'Number of items the NPC is vending
    VendItems() As Obj      'Information on the item the NPC is vending
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
    Moving As Byte
    Aggressive As Byte
    BlinkTimer As Single
    Name As String
    ActionIndex As Byte
    NPCNumber As Integer
End Type

Public Type ObjData
    Name As String              'Name
    ObjType As Byte             'Type (armor, weapon, item, etc)
    GrhIndex As Integer         'Graphic index
    MinHP As Integer            'Bonus HP regenerated
    MaxHP As Integer            'Bonus Max HP raised
    MinHIT As Integer           'Bonus minimum hit
    MaxHIT As Integer           'Bonus maximum hit
    DEF As Integer              'Bonus defence
    ArmorIndex As Byte          'Index of the body sprite
    WeaponIndex As Byte         'Index of the weapon sprite
    WeaponType As Byte          'What type of weapon, if it is a weapon
    Price As Long               'Price of the object
End Type

Public GrhData() As GrhData         'Holds data for the graphic structure
Public SurfaceSize() As Point       'Holds the size of the surfaces for SurfaceDB()
Public BodyData() As BodyData       'Holds data about body structure
Public HeadData() As HeadData       'Holds data about head structure
Public HairData() As HairData       'Holds data about hair structure
Public WeaponData() As WeaponData   'Holds data about weapon structure
Public ObjData() As ObjData

'FPS
Public FPS As Long
Public EndTime As Long
Public ElapsedTime As Single
Public TickPerFrame As Single
Public TimerMultiplier As Single
Public EngineBaseSpeed As Single
Public FramesPerSecCounter As Long

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

'File paths
Public GrhPath As String
Public NPCPath As String
Public OBJPath As String
Public IniPath As String

Private NumBodies As Integer
Private NumGrhs As Integer
Private NumHairs As Integer
Private NumObjs As Integer
Private NumHeads As Integer
Private NumGrhFiles As Integer
Private NumWeapons As Integer
Private TilePixelHeight As Integer
Private TilePixelWidth As Integer
Public EngineRun As Boolean

'The NPC we are editing
Public OpenNPC As NPC
Public NPCNumber As Integer
Public CharList(1) As Char

'********** OUTSIDE FUNCTIONS ***********
Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function writeprivateprofilestring Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Sub Editor_LoadOBJs()
'*****************************************************************
'Loads all the objects and places them in the OBJList on frmMain
'*****************************************************************
Dim FileNum As Byte
Dim i As Integer

    'Get the number of objects
    FileNum = FreeFile
    Open OBJPath & "Count.obj" For Binary As #FileNum
        Get #FileNum, , NumObjs
    Close #FileNum
    ReDim ObjData(1 To NumObjs)
    
    'Get the object information
    For i = 1 To NumObjs
        Open OBJPath & i & ".obj" For Binary As #FileNum
            Get #FileNum, , ObjData(i)
        Close #FileNum
    Next i

End Sub

Sub Editor_UpdateVendItems()
'*****************************************************************
'Display all the items for vending
'*****************************************************************
Dim i As Integer
On Error Resume Next

    With frmMain.OBJList
        .Clear
        If OpenNPC.NumVendItems > 0 Then
            For i = 1 To OpenNPC.NumVendItems
                .AddItem ObjData(OpenNPC.VendItems(i).OBJIndex).Name & " (ID: " & OpenNPC.VendItems(i).OBJIndex & " #: " & OpenNPC.VendItems(i).Amount & ")"
            Next i
        End If
    End With

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
        Grh.FrameCounter = Grh.FrameCounter + (0.0375 * GrhData(Grh.GrhIndex).Speed)
        If Grh.FrameCounter >= GrhData(Grh.GrhIndex).NumFrames + 1 Then
            Do While Grh.FrameCounter >= GrhData(Grh.GrhIndex).NumFrames + 1
                Grh.FrameCounter = Grh.FrameCounter - GrhData(Grh.GrhIndex).NumFrames
            Loop
            If LoopAnim <> True Then Grh.Started = 0
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
    Engine_Render_Rectangle X, Y, GrhData(CurrentGrh.GrhIndex).pixelWidth, GrhData(CurrentGrh.GrhIndex).pixelHeight, GrhData(CurrentGrh.GrhIndex).sX, GrhData(CurrentGrh.GrhIndex).sY, GrhData(CurrentGrh.GrhIndex).pixelWidth, GrhData(CurrentGrh.GrhIndex).pixelHeight, , , 0, GrhData(CurrentGrh.GrhIndex).FileNum, Light1, Light2, Light3, Light4, Shadow
    
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

    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), Len(VertexArray(0))

End Sub

Public Sub Engine_Render_Char(ByVal CharIndex As Long, ByVal PixelOffsetX As Single, ByVal PixelOffsetY As Single)

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
Dim ShadowColor As Long

    '***** Set the variables *****
    RenderColor(1) = -1
    RenderColor(2) = -1
    RenderColor(3) = -1
    RenderColor(4) = -1
    ShadowColor = D3DColorARGB(150, 0, 0, 0)

    '***** Render Shadows *****

    'Draw Body
    If CharList(CharIndex).ActionIndex <= 1 Then

        'Shadow
        Engine_Render_Grh CharList(CharIndex).Body.Walk(CharList(CharIndex).Heading), PixelOffsetX, PixelOffsetY, 1, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, , 1
        Engine_Render_Grh CharList(CharIndex).Weapon.Walk(CharList(CharIndex).Heading), PixelOffsetX + CharList(CharIndex).Weapon.WeaponOffset.X, PixelOffsetY + CharList(CharIndex).Weapon.WeaponOffset.Y, True, True, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, , 1

    Else

        'Start attack animation
        CharList(CharIndex).Body.Attack(CharList(CharIndex).Heading).Started = 0
        CharList(CharIndex).Weapon.Attack(CharList(CharIndex).Heading).FrameCounter = 1

        'Shadow
        Engine_Render_Grh CharList(CharIndex).Body.Attack(CharList(CharIndex).Heading), PixelOffsetX, PixelOffsetY, 1, 1, False, ShadowColor, ShadowColor, ShadowColor, ShadowColor, , 1
        Engine_Render_Grh CharList(CharIndex).Weapon.Attack(CharList(CharIndex).Heading), PixelOffsetX + CharList(CharIndex).Weapon.WeaponOffset.X, PixelOffsetY + CharList(CharIndex).Weapon.WeaponOffset.Y, True, True, False, ShadowColor, ShadowColor, ShadowColor, ShadowColor, , 1

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
    'Draw body and weapon
    If CharList(CharIndex).ActionIndex <= 1 Then
        'Walking
        Engine_Render_Grh CharList(CharIndex).Body.Walk(CharList(CharIndex).Heading), PixelOffsetX, PixelOffsetY, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh CharList(CharIndex).Weapon.Walk(CharList(CharIndex).Heading), PixelOffsetX + CharList(CharIndex).Weapon.WeaponOffset.X, PixelOffsetY + CharList(CharIndex).Weapon.WeaponOffset.Y, True, True, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
    Else
        'Attacking
        Engine_Render_Grh CharList(CharIndex).Body.Attack(CharList(CharIndex).Heading), PixelOffsetX, PixelOffsetY, 1, 0, False, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Engine_Render_Grh CharList(CharIndex).Weapon.Attack(CharList(CharIndex).Heading), PixelOffsetX + CharList(CharIndex).Weapon.WeaponOffset.X, PixelOffsetY + CharList(CharIndex).Weapon.WeaponOffset.Y, True, True, False, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
    End If

    'Draw Head
    If CharList(CharIndex).Aggressive > 0 Then
        'Aggressive
        If CharList(CharIndex).BlinkTimer > 0 Then
            'Blinking
            Engine_Render_Grh CharList(CharIndex).Head.AgrBlink(CharList(CharIndex).HeadHeading), PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Else
            'Normal
            Engine_Render_Grh CharList(CharIndex).Head.AgrHead(CharList(CharIndex).HeadHeading), PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        End If
    Else
        'Not Aggressive
        If CharList(CharIndex).BlinkTimer > 0 Then
            'Blinking
            Engine_Render_Grh CharList(CharIndex).Head.Blink(CharList(CharIndex).HeadHeading), PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        Else
            'Normal
            Engine_Render_Grh CharList(CharIndex).Head.Head(CharList(CharIndex).HeadHeading), PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
        End If
    End If

    'Draw Hair
    Engine_Render_Grh CharList(CharIndex).Hair.Hair(CharList(CharIndex).HeadHeading), PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, 1, 0, True, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)

End Sub

Sub Editor_SaveNPC(ByVal NPCNum As Integer)
Dim FileNum As Byte
Dim Count As Integer
Dim i As Byte

    NPCNumber = NPCNum
    
    'Check if to replace the highest NPC value (count.npc)
    FileNum = FreeFile
    Open NPCPath & "Count.npc" For Binary As #FileNum
        Get #FileNum, , Count
    Close #FileNum
    If Count < NPCNum Then
        Open NPCPath & "Count.npc" For Binary As #FileNum
            Put #FileNum, , NPCNum
        Close #FileNum
    End If
    
    'Set the information
    With frmMain
        .Caption = "NPC Editor - NPC: " & NPCNumber
        OpenNPC.Name = .NameTxt.Text
        OpenNPC.Movement = .AITxt.Text
        OpenNPC.Char.Weapon = .WeaponTxt.Text
        OpenNPC.Char.Body = .BodyTxt.Text
        OpenNPC.Char.Hair = .HairTxt.Text
        OpenNPC.Char.Head = .HeadTxt.Text
        OpenNPC.Char.Heading = .HeadingTxt.Text
        OpenNPC.GiveEXP = .GiveExpTxt.Text
        OpenNPC.GiveGLD = .GiveGoldTxt.Text
        OpenNPC.RespawnWait = .RespawnTxt.Text
        OpenNPC.Quest = .QuestTxt.Text
        OpenNPC.Desc = .DescTxt.Text
        If .HostileChk.Value = 1 Then OpenNPC.Hostile = 1 Else OpenNPC.Hostile = 0
        If .AttackChk.Value = 1 Then OpenNPC.Attackable = 1 Else OpenNPC.Attackable = 0
        For i = 1 To .StatTxt.UBound
            If i <= NumStats Then OpenNPC.BaseStat(i) = .StatTxt(i).Text
        Next i
    End With
    
    'Load up a NPC
    i = Len(OpenNPC.Desc)
    Open NPCPath & NPCNumber & ".npc" For Binary As #FileNum
        Put #FileNum, , OpenNPC
    Close #FileNum
    
    'Saved
    MsgBox "NPC " & NPCNumber & " successfully saved!", vbOKOnly

End Sub

Sub Editor_OpenNPC(ByVal NPCNum As Integer)
Dim FileNum As Byte
Dim i As Byte

    'Check that the file exists
    If Engine_FileExist(NPCPath & NPCNumber & ".npc", vbNormal) = False Then
        MsgBox "Error! NPC file (" & NPCPath & NPCNumber & ".npc) does not exist!", vbOKOnly
        Exit Sub
    End If

    'Load up a NPC
    NPCNumber = NPCNum
    FileNum = FreeFile
    Open NPCPath & NPCNumber & ".npc" For Binary As #FileNum
        Get #FileNum, , OpenNPC
    Close #FileNum
    
    'Display the information
    With frmMain
        .Caption = "NPC Editor - NPC: " & NPCNumber
        .NameTxt.Text = OpenNPC.Name
        .AITxt.Text = OpenNPC.Movement
        .WeaponTxt.Text = OpenNPC.Char.Weapon
        .BodyTxt.Text = OpenNPC.Char.Body
        .HairTxt.Text = OpenNPC.Char.Hair
        .HeadTxt.Text = OpenNPC.Char.Head
        .HeadingTxt.Text = OpenNPC.Char.Heading
        .GiveExpTxt.Text = OpenNPC.GiveEXP
        .GiveGoldTxt.Text = OpenNPC.GiveGLD
        .RespawnTxt.Text = OpenNPC.RespawnWait
        .QuestTxt.Text = OpenNPC.Quest
        .DescTxt.Text = OpenNPC.Desc
        If OpenNPC.Hostile Then .HostileChk.Value = 1 Else .HostileChk.Value = 0
        If OpenNPC.Attackable Then .AttackChk.Value = 1 Else .AttackChk.Value = 0
        
        'Stats
        For i = 1 To .StatTxt.UBound
            If i <= NumStats Then
                .StatTxt(i).Text = OpenNPC.BaseStat(i)
            Else
                .StatTxt(i).Text = "N/A"
                .StatTxt(i).Enabled = False
            End If
        Next i
        
        'Objects (vending)
        Editor_UpdateVendItems
            
    End With
    
    CharList(1).Body = BodyData(1)
    CharList(1).Head = HeadData(1)
    CharList(1).Hair = HairData(1)
    CharList(1).Weapon = WeaponData(0)
    CharList(1).Heading = SOUTH
    CharList(1).HeadHeading = SOUTH
    CharList(1).Moving = 1
    CharList(1).Active = 1

End Sub

Sub Editor_SetNPCGrhs()
Dim EmptyBodyData As BodyData
Dim EmptyHeadData As HeadData
Dim EmptyHairData As HairData
Dim EmptyWeaponData As WeaponData

    If Val(frmMain.BodyTxt.Text) > 0 And Val(frmMain.BodyTxt.Text) <= UBound(BodyData) Then CharList(1).Body = BodyData(Val(frmMain.BodyTxt.Text)) Else CharList(1).Body = EmptyBodyData
    If Val(frmMain.HeadTxt.Text) > 0 And Val(frmMain.HeadTxt.Text) <= UBound(HeadData) Then CharList(1).Head = HeadData(Val(frmMain.HeadTxt.Text)) Else CharList(1).Head = EmptyHeadData
    If Val(frmMain.HairTxt.Text) > 0 And Val(frmMain.HairTxt.Text) <= UBound(HairData) Then CharList(1).Hair = HairData(Val(frmMain.HairTxt.Text)) Else CharList(1).Hair = EmptyHairData
    If Val(frmMain.WeaponTxt.Text) > 0 And Val(frmMain.WeaponTxt.Text) <= UBound(WeaponData) Then CharList(1).Weapon = WeaponData(Val(frmMain.WeaponTxt.Text)) Else CharList(1).Weapon = EmptyWeaponData
    CharList(1).Heading = SOUTH
    CharList(1).HeadHeading = SOUTH
    CharList(1).Moving = 1
    CharList(1).Active = 1

End Sub

Sub Main()
Dim FilePath As String

    'Set file paths
    GrhPath = App.Path & "\Grh\"
    NPCPath = App.Path & "\NPCs\"
    OBJPath = App.Path & "\OBJs\"
    IniPath = App.Path & "\Data\"
    
    frmMain.Show
    
    Engine_Init_TileEngine frmMain.PreviewPic.hwnd, frmMain.PreviewPic.ScaleWidth, frmMain.PreviewPic.ScaleHeight, 32, 32, 1, 0.011
    
    'Check for the first NPC
    If Command$ = "" Then
        If Engine_FileExist(NPCPath & "1.npc", vbNormal) Then Editor_OpenNPC 1
    Else
        FilePath = Mid$(Command$, 2, Len(Command$) - 2) 'Retrieve the filepath from Command$ and crop off the "'s
        Editor_OpenNPC Val(Right$(FilePath, Len(FilePath) - Len(NPCPath)))
    End If
    
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

Sub Engine_UnloadAllForms()

'*****************************************************************
'Unloads all forms
'*****************************************************************

Dim frm As Form

    For Each frm In VB.Forms
        Unload frm
    Next

End Sub

Sub Engine_Init_BodyData()

'*****************************************************************
'Loads Body.dat
'*****************************************************************

Dim LoopC As Long
'Get number of bodies

    NumBodies = CInt(Engine_Var_Get(IniPath & "Body.dat", "INIT", "NumBodies"))
    'Resize array
    ReDim BodyData(1 To NumBodies) As BodyData
    'Fill list
    For LoopC = 1 To NumBodies
        Engine_Init_Grh BodyData(LoopC).Walk(1), CInt(Engine_Var_Get(IniPath & "Body.dat", "Body" & LoopC, "Walk1")), 0
        Engine_Init_Grh BodyData(LoopC).Walk(2), CInt(Engine_Var_Get(IniPath & "Body.dat", "Body" & LoopC, "Walk2")), 0
        Engine_Init_Grh BodyData(LoopC).Walk(3), CInt(Engine_Var_Get(IniPath & "Body.dat", "Body" & LoopC, "Walk3")), 0
        Engine_Init_Grh BodyData(LoopC).Walk(4), CInt(Engine_Var_Get(IniPath & "Body.dat", "Body" & LoopC, "Walk4")), 0
        BodyData(LoopC).Walk(5) = BodyData(LoopC).Walk(1)
        BodyData(LoopC).Walk(6) = BodyData(LoopC).Walk(2)
        BodyData(LoopC).Walk(7) = BodyData(LoopC).Walk(3)
        BodyData(LoopC).Walk(8) = BodyData(LoopC).Walk(4)
        BodyData(LoopC).HeadOffset.X = CLng(Engine_Var_Get(IniPath & "Body.dat", "Body" & LoopC, "HeadOffsetX"))
        BodyData(LoopC).HeadOffset.Y = CLng(Engine_Var_Get(IniPath & "Body.dat", "Body" & LoopC, "HeadOffsetY"))
        Engine_Init_Grh BodyData(LoopC).Attack(1), CInt(Engine_Var_Get(IniPath & "Body.dat", "Body" & LoopC, "Attack1")), 1
        Engine_Init_Grh BodyData(LoopC).Attack(2), CInt(Engine_Var_Get(IniPath & "Body.dat", "Body" & LoopC, "Attack2")), 1
        Engine_Init_Grh BodyData(LoopC).Attack(3), CInt(Engine_Var_Get(IniPath & "Body.dat", "Body" & LoopC, "Attack3")), 1
        Engine_Init_Grh BodyData(LoopC).Attack(4), CInt(Engine_Var_Get(IniPath & "Body.dat", "Body" & LoopC, "Attack4")), 1
        BodyData(LoopC).Attack(5) = BodyData(LoopC).Attack(1)
        BodyData(LoopC).Attack(6) = BodyData(LoopC).Attack(2)
        BodyData(LoopC).Attack(7) = BodyData(LoopC).Attack(3)
        BodyData(LoopC).Attack(8) = BodyData(LoopC).Attack(4)
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
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.PreviewPic.hwnd, D3DCREATEFLAGS, D3DWindow)

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

Sub Engine_Init_Grh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)

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

'*****************************************************************
'Loads Grh.dat
'*****************************************************************

Dim Grh As Integer
Dim Frame As Long

    'Get Number of Graphics
    NumGrhs = CInt(Engine_Var_Get(IniPath & "Grh.ini", "INIT", "NumGrhs"))
    
    'Resize arrays
    ReDim GrhData(1 To NumGrhs) As GrhData
    
    'Open files
    Open IniPath & "Grh.dat" For Binary As #1
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

    NumHairs = CInt(Engine_Var_Get(IniPath & "Hair.dat", "INIT", "NumHairs"))
    'Resize array
    ReDim HairData(0 To NumHairs) As HairData
    'Fill List
    For LoopC = 1 To NumHairs
        For i = 1 To 8
            Engine_Init_Grh HairData(LoopC).Hair(i), CInt(Engine_Var_Get(IniPath & "Hair.dat", Str$(LoopC), Str$(i))), 0
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

    NumHeads = CInt(Engine_Var_Get(IniPath & "Head.dat", "INIT", "NumHeads"))
    'Resize array
    ReDim HeadData(1 To NumHeads) As HeadData
    'Fill List
    For LoopC = 1 To NumHeads
        For i = 1 To 8
            Engine_Init_Grh HeadData(LoopC).Head(i), CInt(Engine_Var_Get(IniPath & "Head.dat", Str$(LoopC), "h" & i)), 0
            Engine_Init_Grh HeadData(LoopC).Blink(i), CInt(Engine_Var_Get(IniPath & "Head.dat", Str$(LoopC), "b" & i)), 0
            Engine_Init_Grh HeadData(LoopC).AgrHead(i), CInt(Engine_Var_Get(IniPath & "Head.dat", Str$(LoopC), "ah" & i)), 0
            Engine_Init_Grh HeadData(LoopC).AgrBlink(i), CInt(Engine_Var_Get(IniPath & "Head.dat", Str$(LoopC), "ab" & i)), 0
        Next i
    Next LoopC

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

Function Engine_Init_TileEngine(ByRef setDisplayFormhWnd As Long, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal Engine_Speed As Single) As Boolean

'*****************************************************************
'Init Tile Engine
'*****************************************************************

    'Set the array sizes by the number of graphic files
    NumGrhFiles = CInt(Engine_Var_Get(IniPath & "Grh.ini", "INIT", "NumGrhFiles"))
    ReDim SurfaceDB(1 To NumGrhFiles)
    ReDim SurfaceSize(1 To NumGrhFiles)
    ReDim SurfaceTimer(1 To NumGrhFiles)
    
    TilePixelWidth = setWindowTileWidth
    TilePixelHeight = setWindowTileHeight

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

    'Load graphic data into memory
    Engine_Init_GrhData
    Engine_Init_BodyData
    Engine_Init_WeaponData
    Engine_Init_HeadData
    Engine_Init_HairData
    Editor_LoadOBJs

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
        Set D3DX = Nothing

        'Clear GRH memory
        For LoopC = 1 To NumGrhFiles
            Set SurfaceDB(LoopC) = Nothing
        Next LoopC

End Sub

Sub Engine_Init_WeaponData()

'*****************************************************************
'Loads Weapon.dat
'*****************************************************************

Dim LoopC As Long
'Get number of weapons

    NumWeapons = CInt(Engine_Var_Get(IniPath & "Weapon.dat", "INIT", "NumWeapons"))
    'Resize array
    ReDim WeaponData(0 To NumWeapons) As WeaponData
    'Fill listn
    For LoopC = 1 To NumWeapons
        Engine_Init_Grh WeaponData(LoopC).Walk(1), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Walk1")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(2), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Walk2")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(3), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Walk3")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(4), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Walk4")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(5), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Walk5")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(6), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Walk6")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(7), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Walk7")), 0
        Engine_Init_Grh WeaponData(LoopC).Walk(8), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Walk8")), 0
        WeaponData(LoopC).WeaponOffset.X = CLng(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "WeaponOffsetX"))
        WeaponData(LoopC).WeaponOffset.Y = CLng(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "WeaponOffsetY"))
        Engine_Init_Grh WeaponData(LoopC).Attack(1), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Attack1")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(2), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Attack2")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(3), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Attack3")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(4), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Attack4")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(5), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Attack5")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(6), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Attack6")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(7), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Attack7")), 1
        Engine_Init_Grh WeaponData(LoopC).Attack(8), CInt(Engine_Var_Get(IniPath & "Weapon.dat", "Weapon" & LoopC, "Attack8")), 1
    Next LoopC

End Sub

Function Engine_FileExist(file As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

    Engine_FileExist = (Dir$(file, FileType) <> "")

End Function
