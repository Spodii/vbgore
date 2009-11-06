Attribute VB_Name = "TileEngine"
Option Explicit

Public Const ShadowColor As Long = 1677721600   'ARGB 100/0/0/0
Public Const HealthColor As Long = -1761673216  'ARGB 150/255/0/0
Public Const ManaColor As Long = -1778384641    'ARGB 150/0/0/255

Public ParticleOffsetX As Long
Public ParticleOffsetY As Long
Public LastOffsetX As Integer           'The last offset values stored, used to get the offset difference
Public LastOffsetY As Integer           ' so the particle engine can adjust weather particles accordingly

Public EnterText As Boolean             'If the text buffer is used (the user is typing a message)
Public EnterTextBuffer As String        'The text in the text buffer
Public EnterTextBufferWidth As Long     'Width of the text buffer

'********** CONSTANTS ***********
'Keep window in the game screen - dont let them move outside of the window bounds
Public Const WindowsInScreen As Boolean = True

'Heading constants
Public Const NORTH As Byte = 1
Public Const EAST As Byte = 2
Public Const SOUTH As Byte = 3
Public Const WEST As Byte = 4
Public Const NORTHEAST As Byte = 5
Public Const SOUTHEAST As Byte = 6
Public Const SOUTHWEST As Byte = 7
Public Const NORTHWEST As Byte = 8

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

'Font colors
Public Const FontColor_Talk As Long = -1
Public Const FontColor_Info As Long = -16711936
Public Const FontColor_Fight As Long = -65536
Public Const FontColor_Quest As Long = -256
Private Const ChatTextBufferSize As Integer = 200
Public Const DamageDisplayTime As Integer = 2000
Public Const BufferSize As Long = 40
Public Const MouseSpeed As Single = 1.5

'********** MUSIC ***********
Public Const Music_MaxVolume As Long = 100
Public Const Music_MaxBalance As Long = 100
Public Const Music_MaxSpeed As Long = 226
Public Const NumMusicBuffers As Long = 1
Public DirectShow_Event(1 To NumMusicBuffers) As IMediaEvent
Public DirectShow_Control(1 To NumMusicBuffers) As IMediaControl
Public DirectShow_Position(1 To NumMusicBuffers) As IMediaPosition
Public DirectShow_Audio(1 To NumMusicBuffers) As IBasicAudio

'********** Custom Fonts ************
'vbGORE Font Header
Private Type VFH
    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
End Type

Private Type CustomFont
    HeaderInfo As VFH           'Holds the header information
    Texture As Direct3DTexture8 'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
End Type

Public Font_Default As CustomFont   'Describes our custom font "default"

'********** WEATHER ***********
Public Type LightType
    Light(1 To 24) As Long
End Type
Public SaveLightBuffer(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As LightType
Public WeatherEffectIndex As Long   'Index returned by the weather effect initialization
Public DoLightning As Byte          'Are we using lightning? 1 = Yes, 2 = No
Public LightningTimer As Single     'How long until our next lightning bolt strikes
Public FlashTimer As Single         'How long until the flash goes away (being > 0 states flash is happening)
Public LightningX As Integer        'Position of the lightning (top-left corner)
Public LightningY As Integer
Public WeatherSfx1 As DirectSoundSecondaryBuffer8   'Weather buffers - dont add more unless you need more for
Public WeatherSfx2 As DirectSoundSecondaryBuffer8   ' one weather effect (ie rain, wind, lightning)

'********** TYPES ***********
'Text buffer
Type ChatTextBuffer
    Text As String
    Color As Long
    Width As Long
End Type

Private ChatTextBuffer(1 To ChatTextBufferSize) As ChatTextBuffer

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
Private Type WorldPos
    X As Byte
    Y As Byte
End Type

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
    Speed As Byte
    Running As Byte
    Aggressive As Byte
    MoveOffset As FloatPos
    BlinkTimer As Single        'The length of the actual blinking
    StartBlinkTimer As Single   'How long until a blink starts
    ScrollDirectionX As Integer
    ScrollDirectionY As Integer
    BubbleStr As String
    BubbleTime As Long
    name As String
    NameOffset As Integer       'Used to acquire the center position for the name
    ActionIndex As Byte
    HealthPercent As Byte
    ManaPercent As Byte
    CharStatus As CharStatus
    Emoticon As Grh
    EmoFade As Single
    EmoDir As Byte      'Direction the fading is going - 0 = Stopped, 1 = Up, 2 = Down
    NPCChatIndex As Byte
    NPCChatLine As Byte
    NPCChatDelay As Long
End Type

'Holds info about each tile position
Public Type MapBlock
    BlockedAttack As Byte
    Graphic(1 To 6) As Grh
    Light(1 To 24) As Long
    Shadow(1 To 6) As Byte
    Sign As Integer
    Blocked As Byte
    Mailbox As Byte
    Warp As Byte
    Sfx As DirectSoundSecondaryBuffer8
End Type

'Hold info about each map
Public Type MapInfo
    name As String
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
    Offset As Position
    Grh As Grh
End Type

'Describes the effects layer
Private Type EffectSurface
    Pos As WorldPos
    Grh As Grh
    Angle As Single
    Time As Long
    Animated As Byte
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
Private AddtoUserPos As Position    'For moving user
Public UserCharIndex As Integer
Public EngineRun As Boolean
Private FPS As Long
Private FramesPerSecCounter As Long
Private FPSLastCheck As Long
Private SaveLastCheck As Long

'Main view size size in tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

'How many tiles the engine "looks ahead" when drawing the screen
Public TileBufferSize As Integer

'Tile size in pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

'Totals
Private NumBodies As Integer    'Number of bodies
Private NumHeads As Integer     'Number of heads
Private NumHairs As Integer     'Number of hairs
Private NumWeapons As Integer   'Number of weapons
Private NumGrhs As Long         'Number of grhs
Private NumWings As Integer     'Number of wings
Public NumSfx As Integer        'Number of sound effects
Public NumMaps As Integer       'Number of maps
Public NumGrhFiles As Integer   'Number of pngs
Public LastChar As Integer      'Last character
Public LastObj As Integer       'Last object
Public LastBlood As Integer     'Last blood splatter index used
Public LastEffect As Integer    'Last effect index used
Public LastDamage As Integer    'Last damage counter text index used
Public LastProjectile As Integer    'Last projectile index used

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

'********** GAME WINDOWS ***********
Public Const SkillListX As Integer = 750            'Position where the skill list where appear
Public Const SkillListY As Integer = 525            ' (indicates the bottom-right corner)
Public Const SkillListWidth As Integer = 5          'How many skills wide the skill popup list is
Public Const GUIColorValue As Long = -1090519041    'ARGB 190/255/255/255
Public Const QuickBarWindow As Byte = 1
Public Const InventoryWindow As Byte = 2
Public Const ShopWindow As Byte = 3
Public Const MailboxWindow As Byte = 4
Public Const ViewMessageWindow As Byte = 5
Public Const WriteMessageWindow As Byte = 6
Public Const AmountWindow As Byte = 7
Public Const MenuWindow As Byte = 8
Public Const ChatWindow As Byte = 9
Public Const StatWindow As Byte = 10
Public Const BankWindow As Byte = 11
Private Const NumGameWindows As Byte = 11
Public Const MaxMailObjs As Byte = 10
Public SelGameWindow As Byte            'The selected game window (mouse is down, not last-clicked)
Public SelMessage As Byte               'The selected message in the mailbox
Public LastClickedWindow As Byte        'The last game window to be clicked
Public ShowGameWindow(1 To NumGameWindows) As Byte  'What game windows are visible
Public MailboxListBuffer As String      'Holds the list of text for the mailbox
Public AmountWindowValue As String      'How much of the item will be dropped from the amount window
Public AmountWindowItemIndex As Byte    'Index of the item to be dropped/sold/sent when the amount window pops up
Public AmountWindowUsage As Byte        'The usage combination for the amount window (as defined with below constants)
Public DrawSkillList As Byte            'If the skills list is to be drawn
Public QuickBarSetSlot As Byte          'What slot on the quickbar was clicked to be set
Public DragSourceWindow As Byte         'The window the item was dragged from
Public DragItemSlot As Byte             'Holds what slot an item is being dragged from in the inventory

'AmountWindowUsage constants
Public Const AW_Drop As Byte = 0
Public Const AW_InvToInv As Byte = 1
Public Const AW_InvToShop As Byte = 2
Public Const AW_InvToBank As Byte = 3
Public Const AW_InvToMail As Byte = 4
Public Const AW_ShopToInv As Byte = 5
Public Const AW_BankToInv As Byte = 6

Private Type QuickBarIDData
    Type As Byte    'Type of information in the quick bar (Item, Skill, etc)
    ID As Byte      'The ID of whatever is being held (Item = Inventory Slot, Skill = SkillID)
End Type
Public QuickBarID(1 To 12) As QuickBarIDData
Public Const QuickBarType_Skill As Byte = 1
Public Const QuickBarType_Item As Byte = 2

Private Type SkillListData
    SkillID As Byte
    X As Long
    Y As Long
End Type
Public SkillList() As SkillListData
Public SkillListSize As Byte

Private Type RMailData          'The mail data for the message being read
    Subject As String
    WriterName As String
    RecieveDate As Date
    Message As String
    New As Byte
    Obj(1 To MaxMailObjs) As Integer
    ObjName(1 To MaxMailObjs) As String
    ObjAmount(1 To MaxMailObjs) As Integer
End Type

Public ReadMailData As RMailData

Private Type WMailData          'The mail data for the message being written
    Subject As String
    RecieverName As String
    Message As String
    ObjIndex(1 To MaxMailObjs) As Integer
    ObjAmount(1 To MaxMailObjs) As Integer
End Type

Public WriteMailData As WMailData

Public Enum WriteMailSelectedControl
    wmFrom = 1
    wmSubject = 2
    wmMessage = 3
End Enum
#If False Then
Private From, Subject, Message
#End If
Public WMSelCon As WriteMailSelectedControl

Private Type Rectangle          'A normal little rectangle
    X As Integer
    Y As Integer
    Width As Integer
    Height As Integer
End Type

Private Type WindowMessage      'Write/Read message window
    Screen As Rectangle
    From As Rectangle
    Subject As Rectangle
    Message As Rectangle
    Image(1 To MaxMailObjs) As Rectangle
    SkinGrh As Grh
End Type

Private Type WindowQuickBar     'Quick bar window
    Screen As Rectangle
    Image(1 To 12) As Rectangle
    SkinGrh As Grh
End Type

Private Type WindowInventory    'User inventory window
    Screen As Rectangle
    Image(1 To 49) As Rectangle
    SkinGrh As Grh
End Type

Private Type WindowMailbox      'Mailbox window
    Screen As Rectangle
    WriteLbl As Rectangle
    DeleteLbl As Rectangle
    ReadLbl As Rectangle
    List As Rectangle
    SkinGrh As Grh
End Type

Private Type WindowAmount       'Amount window
    Screen As Rectangle
    Value As Rectangle
    SkinGrh As Grh
End Type

Private Type ChatWindow         'Chat buffer/input window
    Screen As Rectangle
    Text As Rectangle
    SkinGrh As Grh
End Type

Private Type WindowMenu
    Text As Rectangle
    Screen As Rectangle
    QuitLbl As Rectangle
    SkinGrh As Grh
End Type

Private Type StatWindow
    Screen As Rectangle
    AddStr As Rectangle
    AddAgi As Rectangle
    AddMag As Rectangle
    Str As Rectangle
    Agi As Rectangle
    Mag As Rectangle
    Points As Rectangle
    Dmg As Rectangle
    DEF As Rectangle
    Gold As Rectangle
    AddGrh As Grh
    SkinGrh As Grh
End Type

Public Type GameWindow          'List of all the different game windows
    QuickBar As WindowQuickBar
    Inventory As WindowInventory
    Shop As WindowInventory
    Mailbox As WindowMailbox
    ViewMessage As WindowMessage
    WriteMessage As WindowMessage
    Amount As WindowAmount
    Menu As WindowMenu
    ChatWindow As ChatWindow
    StatWindow As StatWindow
    Bank As WindowInventory
End Type

Public GameWindow As GameWindow

'********** Direct X ***********
Public Const SurfaceTimerMax As Single = 300000     'How long a texture stays in memory unused (miliseconds)
Public Const SoundBufferTimerMax As Single = 300000 'How long a sound stays in memory unused (miliseconds)
Public SurfaceDB() As Direct3DTexture8          'The list of all the textures
Public SurfaceTimer() As Long                   'How long until the surface unloads
Public SoundBufferTimer() As Long               'How long until the sound buffer unloads
Public LastTexture As Long                      'The last texture used
Public D3DWindow As D3DPRESENT_PARAMETERS       'Describes the viewport and used to restore when in fullscreen
Public UsedCreateFlags As CONST_D3DCREATEFLAGS  'The flags we used to create the device when it first succeeded

'Texture for particle effects - this is handled differently then the rest of the graphics
Public ParticleTexture(1 To 12) As Direct3DTexture8

'DirectX 8 Objects
Private DX As DirectX8
Private DI As DirectInput8
Private D3D As Direct3D8
Public D3DX As D3DX8
Public DIDevice As DirectInputDevice8
Public D3DDevice As Direct3DDevice8
Private DS As DirectSound8
Private DSBDesc As DSBUFFERDESC
Public DSBuffer() As DirectSoundSecondaryBuffer8
Public MousePos As POINTAPI
Public MousePosAdd As POINTAPI
Public MouseEvent As Long
Public MouseLeftDown As Byte
Public MouseRightDown As Byte

'Describes a transformable lit vertex
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    Rhw As Single
    Color As Long
    tu As Single
    tv As Single
End Type

'The size of a FVF vertex
Public Const FVF_Size As Long = 28

'Holds the general purpose vertex array (for building rectangles only)
Private VertexArray(0 To 3) As TLVERTEX

'Holds the temp vertex array to build vertex buffers
Private tVA() As TLVERTEX

'Chat vertex buffer information
Private ChatArrayUbound As Long
Private ChatVB As Direct3DVertexBuffer8

'Projectile information
Public Type Projectile
    X As Single
    Y As Single
    tX As Single
    tY As Single
    RotateSpeed As Byte
    Rotate As Single
    Grh As Grh
End Type

'Texture information
Public Type TexInfo
    X As Long
    Y As Long
    MipLevels As Long
    BmpFormat As Long
End Type

'If to use the sounds or not
Public UseSounds As Byte

'********** Public ARRAYS ***********
Public GrhData() As GrhData             'Holds data for the graphic structure
Public SurfaceSize() As TexInfo         'Holds the size of the surfaces for SurfaceDB()
Public BodyData() As BodyData           'Holds data about body structure
Public HeadData() As HeadData           'Holds data about head structure
Public HairData() As HairData           'Holds data about hair structure
Public WeaponData() As WeaponData       'Holds data about weapon structure
Public WingData() As WingData           'Holds data about wing structure
Public MapData() As MapBlock            'Holds map data for current map
Public MapInfo As MapInfo               'Holds map info for current map
Public CharList() As Char               'Holds info about all characters on the map
Public OBJList() As FloatSurface        'Holds info about all objects on the map
Public BloodList() As FloatSurface      'Holds info about all the active blood splatters
Public EffectList() As EffectSurface    'Holds info about all the active effects of all types
Public ProjectileList() As Projectile   'Holds info about all the active projectiles (arrows, ninja stars, bullets, etc)
Public DamageList() As DamageTxt        'Holds info on the damage displays

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

Private NotFirstRender As Byte

Public ShownText As String
Public LineBreakChr As String * 1

'Mini-map tiles
Public Type MiniMapTile
    X As Single         'X and Y index of the tile (using the tile position, not pixel position)
    Y As Single
    Color As Long       'The color of the tile
    RoundCorner As Byte 'What corner to round (0 = none, 1 = top-left, 2 = top-right, 3 = bottom-left, 4 = bottom-right)
End Type
Public MiniMapVBSize As Long    'Size of the vertex buffer (number of verticies, or Tiles x 8)
Public MiniMapVB As Direct3DVertexBuffer8   'Holds the information needed to render the mini-map (not including characters)
Public ShowMiniMap As Byte

'********** OUTSIDE FUNCTIONS ***********
Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function IntersectRect Lib "User32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

Sub Engine_MakeChatBubble(ByVal CharIndex As Integer, ByVal Text As String)

'************************************************************
'Adds text to a chat bubble
'************************************************************
    
    If LenB(Text) = 0 Then Exit Sub 'No text passed
    CharList(CharIndex).BubbleStr = Text
    CharList(CharIndex).BubbleTime = 5000
    
End Sub

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

Public Sub Engine_AddToChatTextBuffer(ByVal Text As String, ByVal Color As Long)

'************************************************************
'Adds text to the chat text buffer
'Buffer is order from bottom to top
'************************************************************
Dim TempSplit() As String
Dim TSLoop As Long
Dim LastSpace As Long
Dim Size As Long
Dim i As Long
Dim b As Long
Dim j As Long

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(Text, vbCrLf)
    
    For TSLoop = 0 To UBound(TempSplit)
    
        'Clear the values for the new line
        Size = 0
        b = 1
        LastSpace = 1
        
        'Loop through all the characters
        For i = 1 To Len(TempSplit(TSLoop))
        
            'If it is a space, store it so we can easily break at it
            Select Case Mid$(TempSplit(TSLoop), i, 1)
                Case " ": LastSpace = i
                Case "_": LastSpace = i
                Case "-": LastSpace = i
            End Select
            
            'Add up the size - Do not count the "|" character (high-lighter)!
            If Not Mid$(TempSplit(TSLoop), i, 1) = "|" Then
                Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
            End If
            
            'Check for too large of a size
            If Size > GameWindow.ChatWindow.Text.Width Then
                
                'Check if the last space was too far back
                If i - LastSpace > 10 Then
                    
                    'Too far away to the last space, so break at the last character
                    Engine_AddToChatTextBuffer2 Trim$(Mid$(TempSplit(TSLoop), b, (i - 1) - b)), Color
                    b = i - 1
                    Size = 0
                
                Else
                
                    'Break at the last space to preserve the word
                    Engine_AddToChatTextBuffer2 Trim$(Mid$(TempSplit(TSLoop), b, LastSpace - b)), Color
                    b = LastSpace + 1
                    
                    'Count all the words we ignored (the ones that weren't printed, but are before "i")
                    Size = Engine_GetTextWidth(Mid$(TempSplit(TSLoop), LastSpace, i - LastSpace))
 
                End If
                
            End If
            
            'This handles the remainder
            If i = Len(TempSplit(TSLoop)) Then
                If b <> i Then Engine_AddToChatTextBuffer2 Mid$(TempSplit(TSLoop), b, i), Color
            End If
            
        Next i
        
    Next TSLoop
    
    'Only update if we have set up the text (that way we can add to the buffer before it is even made)
    If Font_Default.RowPitch = 0 Then Exit Sub

    'Update the array
    Engine_UpdateChatArray

End Sub

Private Sub Engine_AddToChatTextBuffer2(ByVal Text As String, ByVal Color As Long)

'************************************************************
'Actually adds the text to the buffer
'************************************************************
Dim LoopC As Long

    'Move all other text up
    For LoopC = (ChatTextBufferSize - 1) To 1 Step -1
        ChatTextBuffer(LoopC + 1) = ChatTextBuffer(LoopC)
    Next LoopC
    
    'Set the values
    ChatTextBuffer(1).Width = Engine_GetTextWidth(Text)
    ChatTextBuffer(1).Text = Text
    ChatTextBuffer(1).Color = Color

End Sub

Public Sub Engine_UpdateChatArray()

'************************************************************
'Update the array representing the text in the chat buffer
'************************************************************
Dim Chunk As Integer
Dim Count As Integer
Dim LoopC As Byte
Dim Ascii As Byte
Dim Row As Long
Dim Pos As Long
Dim u As Single
Dim V As Single
Dim X As Single
Dim Y As Single
Dim Y2 As Single
Dim i As Long
Dim j As Long
Dim s As String
Dim Size As Integer
Dim KeyPhrase As Byte
Dim ResetColor As Byte
Dim TempColor As Long

    'Set the position
    If ChatBufferChunk <= 1 Then ChatBufferChunk = 1
    Chunk = 12
    
    'Get the number of characters in all the visible buffer
    Size = 0
    For LoopC = (Chunk * ChatBufferChunk) - 11 To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        Size = Size + Len(ChatTextBuffer(LoopC).Text)
        
        'Remove the "|"'s from the count
        For i = 1 To Size
            If Mid$(ChatTextBuffer(LoopC).Text, i, 1) = "|" Then j = j + 1
        Next i
        
    Next LoopC
    Size = Size - j
    ChatArrayUbound = Size * 6 - 1
    ReDim tVA(0 To ChatArrayUbound) 'Size our array to fix the 6 verticies of each character

    'Set the base position
    X = GameWindow.ChatWindow.Screen.X + GameWindow.ChatWindow.Text.X
    Y = GameWindow.ChatWindow.Screen.Y + GameWindow.ChatWindow.Text.X 'We assume the border is the same size on all sides

    'Loop through each buffer string
    For LoopC = (Chunk * ChatBufferChunk) - 11 To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        If ChatBufferChunk * Chunk > ChatTextBufferSize Then ChatBufferChunk = ChatBufferChunk - 1
        
        'Set the temp color
        TempColor = ChatTextBuffer(LoopC).Color
        
        'Set the Y position to be used
        Y2 = Y - (LoopC * 10) + (Chunk * ChatBufferChunk * 10)
        
        'Loop through each line if there are line breaks (vbCrLf)
        Count = 0   'Counts the offset value we are on
        If ChatTextBuffer(LoopC).Text <> "" Then 'Dont bother with empty strings
            
            'Loop through the characters
            For j = 1 To Len(ChatTextBuffer(LoopC).Text)
            
                'Convert the character to the ascii value
                Ascii = Asc(Mid$(ChatTextBuffer(LoopC).Text, j, 1))
                
                'Check for a key phrase
                If Ascii = 124 Then
                    KeyPhrase = (Not KeyPhrase)
                    If KeyPhrase Then TempColor = D3DColorARGB(255, 255, 0, 0) Else ResetColor = 1
                Else
                
                    'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
                    Row = (Ascii - Font_Default.HeaderInfo.BaseCharOffset) \ Font_Default.RowPitch
                    u = ((Ascii - Font_Default.HeaderInfo.BaseCharOffset) - (Row * Font_Default.RowPitch)) * Font_Default.ColFactor
                    V = Row * Font_Default.RowFactor

                    'Set up the verticies
                    '    4____5
                    ' 1|\\    |  1 = 4
                    '  | \\   |  3 = 6
                    '  |  \\  |
                    '  |   \\ |
                    ' 2|____\\|
                    '       3 6
                    
                    'Triangle 1
                    With tVA(0 + (6 * Pos))   'Top-left corner
                        .Color = TempColor
                        .X = X + Count
                        .Y = Y2
                        .tu = u
                        .tv = V
                        .Rhw = 1
                    End With
                    With tVA(1 + (6 * Pos))   'Bottom-left corner
                        .Color = TempColor
                        .X = X + Count
                        .Y = Y2 + Font_Default.HeaderInfo.CellHeight
                        .tu = u
                        .tv = V + Font_Default.RowFactor
                        .Rhw = 1
                    End With
                    With tVA(2 + (6 * Pos))   'Bottom-right corner
                        .Color = TempColor
                        .X = X + Count + Font_Default.HeaderInfo.CellWidth
                        .Y = Y2 + Font_Default.HeaderInfo.CellHeight
                        .tu = u + Font_Default.ColFactor
                        .tv = V + Font_Default.RowFactor
                        .Rhw = 1
                    End With
                    
                    'Triangle 2 (only one new verticy is needed)
                    tVA(3 + (6 * Pos)) = tVA(0 + (6 * Pos)) 'Top-left corner
                    With tVA(4 + (6 * Pos))   'Top-right corner
                        .Color = TempColor
                        .X = X + Count + Font_Default.HeaderInfo.CellWidth
                        .Y = Y2
                        .tu = u + Font_Default.ColFactor
                        .tv = V
                        .Rhw = 1
                    End With
                    tVA(5 + (6 * Pos)) = tVA(2 + (6 * Pos))

                    'Update the character we are on
                    Pos = Pos + 1
    
                    'Shift over the the position to render the next character
                    Count = Count + Font_Default.HeaderInfo.CharWidth(Ascii)

                End If
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = ChatTextBuffer(LoopC).Color
                End If
                
            Next j
            
        End If

    Next LoopC

    'Set the vertex array to the vertex buffer
    If Pos <= 0 Then Pos = 1
    If ObjPtr(D3DDevice) Then   'Make sure the D3DDevice exists - this will only return false if we received messages before it had time to load
        Set ChatVB = D3DDevice.CreateVertexBuffer(FVF_Size * Pos * 6, 0, FVF, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVB, 0, FVF_Size * Pos * 6, 0, tVA(0)
    End If
    
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
        
    On Error GoTo 0

End Function

Public Sub Engine_Blood_Create(ByVal X As Integer, ByVal Y As Integer)

'*****************************************************************
'Create a blood splatter
'*****************************************************************

Dim BloodIndex As Integer

'Get the next open blood slot

    Do
        BloodIndex = BloodIndex + 1

        'Update LastBlood if we go over the size of the current array
        If BloodIndex > LastBlood Then
            LastBlood = BloodIndex
            ReDim Preserve BloodList(1 To LastBlood)
            Exit Do
        End If

    Loop While BloodList(BloodIndex).Grh.GrhIndex > 0

    'Fill in the values
    BloodList(BloodIndex).Pos.X = X
    BloodList(BloodIndex).Pos.Y = Y
    Engine_Init_Grh BloodList(BloodIndex).Grh, 21

End Sub

Public Sub Engine_Blood_Erase(ByVal BloodIndex As Integer)

'*****************************************************************
'Erase a blood splatter
'*****************************************************************

Dim j As Integer

'Clear the selected index

    BloodList(BloodIndex).Grh.FrameCounter = 0
    BloodList(BloodIndex).Grh.GrhIndex = 0
    BloodList(BloodIndex).Pos.X = 0
    BloodList(BloodIndex).Pos.Y = 0

    'Update LastBlood
    If j = LastBlood Then
        Do Until BloodList(LastBlood).Grh.GrhIndex > 1

            'Move down one splatter
            LastBlood = LastBlood - 1

            If LastBlood = 0 Then
                Exit Sub
            Else
                'We still have blood, resize the array to end at the last used slot
                ReDim Preserve BloodList(1 To LastBlood)
            End If

        Loop
    End If

End Sub

Sub Engine_ChangeHeading(ByVal Direction As Byte)

'*****************************************************************
'Face user in appropriate direction
'*****************************************************************

    'Check for a valid UserCharIndex
    If UserCharIndex <= 0 Or UserCharIndex > LastChar Then
    
        'We have an invalid user char index, so we must have the wrong one - request an update on the right one
        sndBuf.Put_Byte DataCode.User_RequestUserCharIndex
        Exit Sub
        
    End If
    
    'Only rotate if the user is not already facing that direction
    If CharList(UserCharIndex).Heading <> Direction Then
        sndBuf.Allocate 2
        sndBuf.Put_Byte DataCode.User_Rotate
        sndBuf.Put_Byte Direction
    End If

End Sub

Sub Engine_Char_Erase(ByVal CharIndex As Integer)

'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************

    'Check for targeted character
    If TargetCharIndex = CharIndex Then TargetCharIndex = 0
    If CharIndex = 0 Then Exit Sub
    
    'Make inactive
    CharList(CharIndex).Active = 0

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

Sub Engine_Char_Make(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Speed As Long, ByVal name As String, ByVal Weapon As Integer, ByVal Hair As Integer, ByVal Wings As Integer, ByVal ChatID As Byte, Optional ByVal HP As Byte = 100, Optional ByVal MP As Byte = 100)

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
    CharList(CharIndex).Wings = WingData(Wings)
    CharList(CharIndex).Heading = Heading
    CharList(CharIndex).HeadHeading = Heading
    CharList(CharIndex).HealthPercent = HP
    CharList(CharIndex).ManaPercent = MP
    CharList(CharIndex).Speed = Speed
    CharList(CharIndex).NPCChatIndex = ChatID
    
    'Update position
    CharList(CharIndex).Pos.X = X
    CharList(CharIndex).Pos.Y = Y

    'Make active
    CharList(CharIndex).Active = 1
    
    'Calculate the name length so we can center the name above the head
    CharList(CharIndex).name = name
    CharList(CharIndex).NameOffset = Engine_GetTextWidth(name) * 0.5

    'Set action index
    CharList(CharIndex).ActionIndex = 0

End Sub

Sub Engine_Char_Move_ByHead(ByVal CharIndex As Integer, ByVal nHeading As Byte, ByVal Running As Byte)

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
    CharList(CharIndex).Running = Running

End Sub

Sub Engine_Char_Move_ByPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer, ByVal Running As Byte)

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
    CharList(CharIndex).Running = Running
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
    
    'If the targeted character move, re-check if the path is blocked
    If TargetCharIndex > 0 Then
        If CharIndex = UserCharIndex Or CharIndex = TargetCharIndex Then
            ClearPathToTarget = Engine_ClearPath(CharList(UserCharIndex).Pos.X, CharList(UserCharIndex).Pos.Y, CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y)
        End If
    End If

End Sub

Sub Engine_ConvertCPtoTP(ByVal StartPixelLeft As Integer, ByVal StartPixelTop As Integer, ByVal cx As Integer, ByVal cy As Integer, ByRef tX As Integer, ByRef tY As Integer)

'******************************************
'Converts where the user clicks in the main window
'to a tile position
'******************************************

    tX = UserPos.X + (cx - StartPixelLeft) \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.Y + (cy - StartPixelTop) \ TilePixelHeight - WindowTileHeight \ 2

End Sub

Public Sub Engine_Damage_Create(ByVal X As Integer, ByVal Y As Integer, ByVal Value As Integer)

'*****************************************************************
'Create damage text
'*****************************************************************

Dim DamageIndex As Integer

'Get the next open damage slot

    Do
        DamageIndex = DamageIndex + 1

        'Update LastDamage if we go over the size of the current array
        If DamageIndex > LastDamage Then
            LastDamage = DamageIndex
            ReDim Preserve DamageList(1 To LastDamage)
            Exit Do
        End If

    Loop While DamageList(DamageIndex).Counter > 0

    'Set the values
    If Value = -1 Then DamageList(DamageIndex).Value = "Miss" Else DamageList(DamageIndex).Value = Value
    DamageList(DamageIndex).Counter = DamageDisplayTime
    DamageList(DamageIndex).Width = Engine_GetTextWidth(DamageList(DamageIndex).Value)
    DamageList(DamageIndex).Pos.X = X
    DamageList(DamageIndex).Pos.Y = Y

End Sub

Public Sub Engine_Damage_Erase(ByVal DamageIndex As Integer)

'*****************************************************************
'Erase damage text
'*****************************************************************

Dim j As Integer

'Clear the selected index

    DamageList(DamageIndex).Counter = 0
    DamageList(DamageIndex).Value = vbNullString
    DamageList(DamageIndex).Width = 0

    'Update LastDamage
    If j = LastDamage Then
        Do Until DamageList(LastDamage).Counter > 0

            'Move down one splatter
            LastDamage = LastDamage - 1

            If LastDamage = 0 Then
                Exit Sub
            Else
                'We still have damage text, resize the array to end at the last used slot
                ReDim Preserve DamageList(1 To LastDamage)
            End If

        Loop
    End If

End Sub

Public Sub Engine_Projectile_Create(ByVal AttackerIndex As Integer, ByVal TargetIndex As Integer, ByVal GrhIndex As Long, ByVal Rotation As Byte)

'*****************************************************************
'Creates a projectile for a ranged weapon
'*****************************************************************

Dim ProjectileIndex As Integer

    If AttackerIndex = 0 Then Exit Sub
    If TargetIndex = 0 Then Exit Sub
    If AttackerIndex > UBound(CharList) Then Exit Sub
    If TargetIndex > UBound(CharList) Then Exit Sub

    'Get the next open projectile slot
    Do
        ProjectileIndex = ProjectileIndex + 1
        
        'Update LastProjectile if we go over the size of the current array
        If ProjectileIndex > LastProjectile Then
            LastProjectile = ProjectileIndex
            ReDim Preserve ProjectileList(1 To LastProjectile)
            Exit Do
        End If
        
    Loop While ProjectileList(ProjectileIndex).Grh.GrhIndex > 0
    
    'Figure out the initial rotation value
    ProjectileList(ProjectileIndex).Rotate = Engine_GetAngle(CharList(AttackerIndex).Pos.X, CharList(AttackerIndex).Pos.Y, CharList(TargetIndex).Pos.X, CharList(TargetIndex).Pos.Y)
    
    'Fill in the values
    ProjectileList(ProjectileIndex).tX = CharList(TargetIndex).Pos.X * 32
    ProjectileList(ProjectileIndex).tY = CharList(TargetIndex).Pos.Y * 32
    ProjectileList(ProjectileIndex).RotateSpeed = Rotation
    ProjectileList(ProjectileIndex).X = CharList(AttackerIndex).Pos.X * 32
    ProjectileList(ProjectileIndex).Y = CharList(AttackerIndex).Pos.Y * 32
    Engine_Init_Grh ProjectileList(ProjectileIndex).Grh, GrhIndex
    
End Sub

Public Sub Engine_Effect_Create(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Long, Optional ByVal Angle As Single = 0, Optional ByVal Time As Long = 0, Optional ByVal Animated As Byte = 1)

'*****************************************************************
'Creates an effect layer for spells and such
'Life is only used if the effect is looped
'*****************************************************************

Dim EffectIndex As Integer

    'Get the next open effect slot
    Do
        EffectIndex = EffectIndex + 1

        'Update LastEffect if we go over the size of the current array
        If EffectIndex > LastEffect Then
            LastEffect = EffectIndex
            ReDim Preserve EffectList(1 To LastEffect)
            Exit Do
        End If

    Loop While EffectList(EffectIndex).Grh.GrhIndex > 0

    'Fill in the values
    If Time > 0 Then EffectList(EffectIndex).Time = timeGetTime + Time Else EffectList(EffectIndex).Time = 0
    EffectList(EffectIndex).Animated = Animated
    EffectList(EffectIndex).Angle = Angle
    EffectList(EffectIndex).Pos.X = X
    EffectList(EffectIndex).Pos.Y = Y
    Engine_Init_Grh EffectList(EffectIndex).Grh, GrhIndex

End Sub

Public Sub Engine_Projectile_Erase(ByVal ProjectileIndex As Integer)

'*****************************************************************
'Erase a projectile by the projectile index
'*****************************************************************

Dim j As Integer

    'Clear the selected index
    ProjectileList(ProjectileIndex).Grh.FrameCounter = 0
    ProjectileList(ProjectileIndex).Grh.GrhIndex = 0
    ProjectileList(ProjectileIndex).X = 0
    ProjectileList(ProjectileIndex).Y = 0
    ProjectileList(ProjectileIndex).tX = 0
    ProjectileList(ProjectileIndex).tY = 0
    ProjectileList(ProjectileIndex).Rotate = 0
    ProjectileList(ProjectileIndex).RotateSpeed = 0
    
    'Update LastProjectile
    If j = LastProjectile Then
        Do Until ProjectileList(ProjectileIndex).Grh.GrhIndex > 1
            
            'Move down one index
            LastProjectile = LastProjectile - 1
            
            If LastProjectile = 0 Then
                Exit Sub
            Else
                'We still have projectiles left, resize the array to end at the last used slot
                ReDim Preserve ProjectileList(1 To LastProjectile)
            End If
            
        Loop
    End If

End Sub

Public Sub Engine_Effect_Erase(ByVal EffectIndex As Integer)

'*****************************************************************
'Erase an effect by the effect index
'*****************************************************************

Dim j As Integer

    'Clear the selected index
    ZeroMemory EffectList(EffectIndex), Len(EffectList(EffectIndex))

    'Update LastEffect
    If j = LastEffect Then
        Do Until EffectList(LastEffect).Grh.GrhIndex > 1

            'Move down one effect
            LastEffect = LastEffect - 1

            If LastEffect = 0 Then
                Erase EffectList
                Exit Sub
            Else
                'We still have effects, resize the array to end at the last used slot
                ReDim Preserve EffectList(1 To LastEffect)
            End If

        Loop
    End If

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
On Error GoTo ErrOut

    If Dir$(File, FileType) <> "" Then Engine_FileExist = True

Exit Function

'An error will most likely be caused by invalid filenames (those that do not follow the file name rules)
ErrOut:

    Engine_FileExist = False
    
End Function

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single

'************************************************************
'Gets the angle between two points in a 2d plane
'************************************************************
Dim SideA As Single
Dim SideC As Single

    On Error GoTo ErrOut

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

Public Function Engine_GetTextWidth(ByVal Text As String) As Integer

'***************************************************
'Returns the width of text
'***************************************************
Dim i As Integer

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
    
    'Loop through the text
    For i = 1 To Len(Text)
        
        'Add up the stored character widths
        Engine_GetTextWidth = Engine_GetTextWidth + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(Text, i, 1)))
        
    Next i

End Function

Sub Engine_Init_Signs()

'*****************************************************************
'Loads the sign messages
'*****************************************************************
Dim NumSigns As Integer
Dim LoopC As Integer
Dim s As String

    'Get the number of signs
    NumSigns = Val(Engine_Var_Get(DataPath & "Signs.dat", "SIGNS", "NumSigns"))
    If NumSigns = 0 Then Exit Sub
    ReDim Signs(1 To NumSigns)
    
    'Loop through the signs, and get the values
    For LoopC = 1 To NumSigns
        Signs(LoopC) = Engine_Var_Get(DataPath & "Signs.dat", "SIGNS", CStr(LoopC))
    Next LoopC
    
End Sub

Function Engine_Init_Messages(ByVal Language As String) As String

'*****************************************************************
'Loads the game messages
'*****************************************************************
Dim LoopC As Byte
Dim s As String

    'Make sure we are working in lowercase (since all our files are in lowercase)
    Language = LCase$(Language)
    
    'Check for a redirection flag (will return nothing if the flag doesn't exist)
    Do  'This "Do" will allow us to do redirections to redirections, even though we shouldn't even do that
        s = Engine_Var_Get(MessagePath & Language & ".ini", "REDIRECT", "TO")
        If s <> "" Then
            If Engine_FileExist(MessagePath & LCase$(s) & ".ini", vbNormal) = False Then
                MsgBox "Invalid language redirection! Could not load system messages!" & vbCrLf & _
                        "Language '" & Language & "' redirected to '" & LCase$(s) & "', which could not be found!", vbOKOnly
                Exit Function
            End If
            Language = LCase$(s)
        Else
        
            'No redirection was found, so move on
            Exit Do
            
        End If
    Loop
    
    Engine_Init_Messages = Language

    'Get the number of messages
    NumMessages = CByte(Engine_Var_Get(MessagePath & "_nummessages.ini", "MAIN", "NumMessages"))
    
    'Check for a valid number of messages
    If NumMessages = 0 Then
        MsgBox "Error loading message count!", vbOKOnly
        Exit Function
    End If
    
    'Resize our message array to hold all the messages
    ReDim Message(1 To NumMessages)
    
    'Loop through every message and find the message string
    For LoopC = 1 To NumMessages
        Message(LoopC) = Engine_Var_Get(MessagePath & Language & ".ini", "MAIN", CStr(LoopC))
        
        'If the message wasn't found, resort to the primary language, English, since that should hold all messages
        If LCase$(Language) <> "english" Then   'Make sure we're not already using English
            If LenB(Trim$(Message(LoopC))) = 0 Then
                Message(LoopC) = Engine_Var_Get(MessagePath & "english.ini", "MAIN", CStr(LoopC))
            End If
        End If
        
    Next LoopC
    
    'Load the NPC chat messages
    Engine_Init_NPCChat Language
    
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
            Engine_Init_Grh BodyData(LoopC).Walk(j), CInt(Engine_Var_Get(DataPath & "Body.dat", Str(LoopC), Str(j))), 0
            Engine_Init_Grh BodyData(LoopC).Attack(j), CInt(Engine_Var_Get(DataPath & "Body.dat", Str(LoopC), "a" & j)), 1
        Next j
        BodyData(LoopC).HeadOffset.X = CLng(Engine_Var_Get(DataPath & "Body.dat", Str(LoopC), "HeadOffsetX"))
        BodyData(LoopC).HeadOffset.Y = CLng(Engine_Var_Get(DataPath & "Body.dat", Str(LoopC), "HeadOffsetY"))
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

Private Sub Engine_Init_Sound()

'************************************************************
'Initialize the 3D sound device
'************************************************************

    'Make sure we try not to load a file while the engine is unloading
    If IsUnloading Then Exit Sub
    
    On Error GoTo ErrOut

    'Create the DirectSound device (with the default device)
    Set DS = DX.DirectSoundCreate("")
    DS.SetCooperativeLevel frmMain.hWnd, DSSCL_PRIORITY
    
    'Set up the buffer description for later use
    'We are only using panning and volume - combined, we will use this to create a custom 3D effect
    DSBDesc.lFlags = DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    
    'Check if the texture exists
    If Engine_FileExist(SfxPath & "Sfx.ini", vbNormal) = False Then
        MsgBox "Error! Could not find the following data file:" & vbCrLf & SfxPath & "Sfx.ini", vbOKOnly
        IsUnloading = 1
        Exit Sub
    End If

    'Get the number of sound effects
    NumSfx = Val(Engine_Var_Get(SfxPath & "Sfx.ini", "INIT", "NumSfx"))
    
    'Resize the sound buffer array
    If NumSfx > 0 Then
        ReDim DSBuffer(1 To NumSfx)
        ReDim SoundBufferTimer(1 To NumSfx)
    End If
    
    On Error GoTo 0
    
    'All successful, use sounds
    UseSounds = 1
    
    Exit Sub
    
ErrOut:

    'Failure loading sounds, so we won't use them
    UseSounds = 0

End Sub

Public Sub Engine_Sound_SetToMap(ByVal SoundID As Integer, ByVal TileX As Byte, ByVal TileY As Byte)

'************************************************************
'Create a looping sound on the tile
'************************************************************

    If UseSounds = 0 Then Exit Sub
    
    'Make sure the sound isn't already going
    If Not MapData(TileX, TileY).Sfx Is Nothing Then
        MapData(TileX, TileY).Sfx.Stop
        Set MapData(TileX, TileY).Sfx = Nothing
    End If
    
    'Create the buffer
    Engine_Sound_Set MapData(TileX, TileY).Sfx, SoundID
    
    'Exit if theres an error
    If MapData(TileX, TileY).Sfx Is Nothing Then Exit Sub

    'Start the loop
    MapData(TileX, TileY).Sfx.Play DSBPLAY_LOOPING
    
    'Since we dont want to start hearing the sound until we have calculated the panning/volume, we set the volume to off for now
    MapData(TileX, TileY).Sfx.SetVolume -10000

End Sub

Public Sub Engine_Sound_UpdateMap()

'************************************************************
'Update the panning and volume on the map's sfx
'************************************************************
Dim SX As Integer
Dim SY As Integer
Dim X As Byte
Dim Y As Byte
Dim L As Long

    If UseSounds = 0 Then Exit Sub

    'Set the user's position to sX/sY
    SX = CharList(UserCharIndex).Pos.X
    SY = CharList(UserCharIndex).Pos.Y
    
    'Loop through all the map tiles
    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            
            'Only update used tiles
            If Not MapData(X, Y).Sfx Is Nothing Then
                
                'Calculate the volume and check for valid range
                L = Engine_Sound_CalcVolume(SX, SY, X, Y)
                If L < -5000 Then
                    MapData(X, Y).Sfx.Stop
                Else
                    If L > 0 Then L = 0
                    If MapData(X, Y).Sfx.GetStatus <> DSBSTATUS_LOOPING Then MapData(X, Y).Sfx.Play DSBPLAY_LOOPING
                    MapData(X, Y).Sfx.SetVolume L
                End If
                
                'Calculate the panning and check for a valid range
                L = Engine_Sound_CalcPan(SX, X)
                If L > 10000 Then L = 10000
                If L < -10000 Then L = -10000
                MapData(X, Y).Sfx.SetPan L
                
            End If
            
        Next Y
    Next X

End Sub

Public Sub Engine_Sound_Play(ByRef SoundBuffer As DirectSoundSecondaryBuffer8, Optional ByVal flags As CONST_DSBPLAYFLAGS = DSBPLAY_DEFAULT)
'************************************************************
'Used for non area-specific sound effects, such as weather
'************************************************************

    'Play the sound
    SoundBuffer.Play flags
    
End Sub

Public Sub Engine_Sound_Erase(ByRef SoundBuffer As DirectSoundSecondaryBuffer8)

'************************************************************
'Erase the sound buffer
'************************************************************
    
    'Make sure the object exists
    If ObjPtr(SoundBuffer) Then
    
        'If it is playing, we have to stop it first
        If SoundBuffer.GetStatus > 0 Then SoundBuffer.Stop
        
        'Clear the object
        Set SoundBuffer = Nothing
        
    End If

End Sub

Public Sub Engine_Sound_Set(ByRef SoundBuffer As DirectSoundSecondaryBuffer8, ByVal SoundID As Integer)

'************************************************************
'Set the SoundID to the sound buffer
'************************************************************

    If UseSounds = 0 Then Exit Sub

    'Check if the sound buffer is in use
    Engine_Sound_Erase SoundBuffer
    
    'Set the buffer
    If Engine_FileExist(SfxPath & SoundID & ".wav", vbNormal) Then Set SoundBuffer = DS.CreateSoundBufferFromFile(SfxPath & SoundID & ".wav", DSBDesc)

End Sub

Public Sub Engine_Sound_Play3D(ByVal SoundID As Integer, TileX As Integer, TileY As Integer)

'************************************************************
'Play a pseudo-3D sound by the sound buffer ID
'************************************************************
Dim SX As Integer
Dim SY As Integer

    If UseSounds = 0 Then Exit Sub

    'Make sure we have the UserCharIndex, or else we cant play the sound! :o
    If UserCharIndex = 0 Then Exit Sub

    'Check for a valid sound
    If SoundID <= 0 Then Exit Sub

    'Create the buffer if needed
    If SoundBufferTimer(SoundID) <= 0 Then
        If DSBuffer(SoundID) Is Nothing Then Engine_Sound_Set DSBuffer(SoundID), SoundID
    End If
    
    'Update the timer
    SoundBufferTimer(SoundID) = SoundBufferTimerMax
    
    'Clear the position (used in case the sound was already playing - we can only have one of each sound play at a time)
    DSBuffer(SoundID).SetCurrentPosition 0
    
    'Set the user's position to sX/sY
    SX = CharList(UserCharIndex).Pos.X
    SY = CharList(UserCharIndex).Pos.Y
    
    'Calculate the panning
    Engine_Sound_Pan DSBuffer(SoundID), Engine_Sound_CalcPan(SX, TileX)
    
    'Calculate the volume
    Engine_Sound_Volume DSBuffer(SoundID), Engine_Sound_CalcVolume(SX, SY, TileX, TileY)
    
    'Play the sound
    DSBuffer(SoundID).Play DSBPLAY_DEFAULT
    
End Sub

Public Function Engine_Sound_CalcPan(ByVal x1 As Integer, ByVal x2 As Integer) As Long

'************************************************************
'Calculate the panning for 3D sound based on the user's position and the sound's position
'************************************************************

    Engine_Sound_CalcPan = (x1 - x2) * 75
    
End Function

Public Function Engine_Sound_CalcVolume(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer) As Long

'************************************************************
'Calculate the volume for 3D sound based on the user's position and the sound's position
'the (Abs(sX - TileX) * 25) is put on the end to make up for the simulated
' volume loss during panning (since one speaker gets muted to create the panning)
'************************************************************
Dim Dist As Single

    'Store the distance
    Dist = Sqr(((Y1 - Y2) * (Y1 - Y2)) + ((x1 - x2) * (x1 - x2)))
    
    'Apply the initial value
    Engine_Sound_CalcVolume = -(Dist * 80) + (Abs(x1 - x2) * 25)
    
    'Once we get out of the screen (>= 13 tiles away) then we want to fade fast
    If Dist > 13 Then Engine_Sound_CalcVolume = Engine_Sound_CalcVolume - ((Dist - 13) * 180)
    
End Function

Private Sub Engine_Sound_Pan(ByRef SoundBuffer As DirectSoundSecondaryBuffer8, ByVal Value As Long)

'************************************************************
'Pan the selected SoundID (-10,000 to 10,000)
'************************************************************

    If SoundBuffer Is Nothing Then Exit Sub
    SoundBuffer.SetPan Value

End Sub

Private Sub Engine_Sound_Volume(ByRef SoundBuffer As DirectSoundSecondaryBuffer8, ByVal Value As Long)

'************************************************************
'Pan the selected SoundID (-10,000 to 0)
'************************************************************

    If UseSounds = 0 Then Exit Sub

    If SoundBuffer Is Nothing Then Exit Sub
    If Value > 0 Then Value = 0
    If Value < -9000 Then Exit Sub  'Too quiet to care about
    SoundBuffer.SetVolume Value

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
    If Windowed Then
        D3DWindow.Windowed = 1  'State that using windowed mode
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
        D3DWindow.BackBufferFormat = DispMode.Format    'Use format just retrieved
    Else
        DispMode.Format = DispMode.Format
        DispMode.Width = 800
        DispMode.Height = 600
        DispMode.Format = D3DFMT_R5G6B5
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
        D3DWindow.BackBufferCount = 1
        D3DWindow.BackBufferFormat = DispMode.Format
        D3DWindow.BackBufferWidth = 800
        D3DWindow.BackBufferHeight = 600
        D3DWindow.hDeviceWindow = frmMain.hWnd
    End If

    'Set the D3DDevices
    If ObjPtr(D3DDevice) Then Set D3DDevice = Nothing
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATEFLAGS, D3DWindow)

    'Store the create flags
    UsedCreateFlags = D3DCREATEFLAGS

    'The Rhw will always be 1, so set it now instead of every call
    For i = 0 To 3
        VertexArray(i).Rhw = 1
    Next i
    
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
    Grh.LastCount = timeGetTime
    Grh.FrameCounter = 1
    Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed

End Sub

Sub Engine_Init_GrhData()

'*****************************************************************
'Loads Grh.dat
'*****************************************************************

Dim Grh As Long
Dim Frame As Long

    'Get Number of Graphics
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
                If GrhData(Grh).Frames(Frame) <= 0 Then
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
            Get #1, , GrhData(Grh).SX
            If GrhData(Grh).SX < 0 Then GoTo ErrorHandler
            Get #1, , GrhData(Grh).SY
            If GrhData(Grh).SY < 0 Then GoTo ErrorHandler
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
    IsUnloading = 1

End Sub

Public Sub Engine_Init_GUI(Optional ByVal LoadCustomPos As Byte = 1)

'************************************************************
'Load skin GUI data
'************************************************************
Dim ImageOffsetX As Long
Dim ImageOffsetY As Long
Dim ImageSpaceX As Long
Dim ImageSpaceY As Long
Dim LoopC As Long
Dim s As String 'Stores the path to our master skins file (.ini)
Dim t As String 'Stores the path to our custom window positions file (.dat)

    s = DataPath & "Skins\" & CurrentSkin & ".ini"
    t = DataPath & "Skins\" & CurrentSkin & ".dat"
    
    'Load Quickbar
    With GameWindow.QuickBar
        If LoadCustomPos Then
            .Screen.X = Val(Engine_Var_Get(t, "QUICKBAR", "ScreenX"))
            .Screen.Y = Val(Engine_Var_Get(t, "QUICKBAR", "ScreenY"))
        Else
            .Screen.X = Val(Engine_Var_Get(s, "QUICKBAR", "ScreenX"))
            .Screen.Y = Val(Engine_Var_Get(s, "QUICKBAR", "ScreenY"))
        End If
        .Screen.Width = Val(Engine_Var_Get(s, "QUICKBAR", "ScreenWidth"))
        .Screen.Height = Val(Engine_Var_Get(s, "QUICKBAR", "ScreenHeight"))
        Engine_Init_Grh .SkinGrh, Val(Engine_Var_Get(s, "QUICKBAR", "Grh"))
    End With
    For LoopC = 1 To 12
        With GameWindow.QuickBar.Image(LoopC)
            .X = Val(Engine_Var_Get(s, "QUICKBAR", "Image" & LoopC & "X"))
            .Y = Val(Engine_Var_Get(s, "QUICKBAR", "Image" & LoopC & "Y"))
            .Width = 32
            .Height = 32
        End With
    Next LoopC
    
    'Load stats window
    With GameWindow.StatWindow
        If LoadCustomPos Then
            .Screen.X = Val(Engine_Var_Get(t, "STATWINDOW", "ScreenX"))
            .Screen.Y = Val(Engine_Var_Get(t, "STATWINDOW", "ScreenY"))
        Else
            .Screen.X = Val(Engine_Var_Get(s, "STATWINDOW", "ScreenX"))
            .Screen.Y = Val(Engine_Var_Get(s, "STATWINDOW", "ScreenY"))
        End If
        .Screen.Width = Val(Engine_Var_Get(s, "STATWINDOW", "ScreenWidth"))
        .Screen.Height = Val(Engine_Var_Get(s, "STATWINDOW", "ScreenHeight"))
        .AddStr.X = Val(Engine_Var_Get(s, "STATWINDOW", "AddStrX"))
        .AddStr.Y = Val(Engine_Var_Get(s, "STATWINDOW", "AddStrY"))
        .AddStr.Width = Val(Engine_Var_Get(s, "STATWINDOW", "AddStrWidth"))
        .AddStr.Height = Val(Engine_Var_Get(s, "STATWINDOW", "AddStrHeight"))
        .AddAgi.X = Val(Engine_Var_Get(s, "STATWINDOW", "AddAgiX"))
        .AddAgi.Y = Val(Engine_Var_Get(s, "STATWINDOW", "AddAgiY"))
        .AddAgi.Width = Val(Engine_Var_Get(s, "STATWINDOW", "AddAgiWidth"))
        .AddAgi.Height = Val(Engine_Var_Get(s, "STATWINDOW", "AddAgiHeight"))
        .AddMag.X = Val(Engine_Var_Get(s, "STATWINDOW", "AddMagX"))
        .AddMag.Y = Val(Engine_Var_Get(s, "STATWINDOW", "AddMagY"))
        .AddMag.Width = Val(Engine_Var_Get(s, "STATWINDOW", "AddMagWidth"))
        .AddMag.Height = Val(Engine_Var_Get(s, "STATWINDOW", "AddMagHeight"))
        .Str.X = Val(Engine_Var_Get(s, "STATWINDOW", "StrX"))
        .Str.Y = Val(Engine_Var_Get(s, "STATWINDOW", "StrY"))
        .Agi.X = Val(Engine_Var_Get(s, "STATWINDOW", "AgiX"))
        .Agi.Y = Val(Engine_Var_Get(s, "STATWINDOW", "AgiY"))
        .Mag.X = Val(Engine_Var_Get(s, "STATWINDOW", "MagX"))
        .Mag.Y = Val(Engine_Var_Get(s, "STATWINDOW", "MagY"))
        .Gold.X = Val(Engine_Var_Get(s, "STATWINDOW", "GoldX"))
        .Gold.Y = Val(Engine_Var_Get(s, "STATWINDOW", "GoldY"))
        .DEF.X = Val(Engine_Var_Get(s, "STATWINDOW", "DefX"))
        .DEF.Y = Val(Engine_Var_Get(s, "STATWINDOW", "DefY"))
        .Dmg.X = Val(Engine_Var_Get(s, "STATWINDOW", "DmgX"))
        .Dmg.Y = Val(Engine_Var_Get(s, "STATWINDOW", "DmgY"))
        .Points.X = Val(Engine_Var_Get(s, "STATWINDOW", "PointsX"))
        .Points.Y = Val(Engine_Var_Get(s, "STATWINDOW", "PointsY"))
        Engine_Init_Grh .AddGrh, Val(Engine_Var_Get(s, "STATWINDOW", "AddGrh"))
        Engine_Init_Grh .SkinGrh, Val(Engine_Var_Get(s, "STATWINDOW", "Grh"))
    End With
    
    'Load chat window
    With GameWindow.ChatWindow
        If LoadCustomPos Then
            .Screen.X = Val(Engine_Var_Get(t, "CHATWINDOW", "ScreenX"))
            .Screen.Y = Val(Engine_Var_Get(t, "CHATWINDOW", "ScreenY"))
        Else
            .Screen.X = Val(Engine_Var_Get(s, "CHATWINDOW", "ScreenX"))
            .Screen.Y = Val(Engine_Var_Get(s, "CHATWINDOW", "ScreenY"))
        End If
        .Screen.Width = Val(Engine_Var_Get(s, "CHATWINDOW", "ScreenWidth"))
        .Screen.Height = Val(Engine_Var_Get(s, "CHATWINDOW", "ScreenHeight"))
        .Text.X = Val(Engine_Var_Get(s, "CHATWINDOW", "ChatX"))
        .Text.Y = Val(Engine_Var_Get(s, "CHATWINDOW", "ChatY"))
        .Text.Width = Val(Engine_Var_Get(s, "CHATWINDOW", "ChatWidth"))
        .Text.Height = Val(Engine_Var_Get(s, "CHATWINDOW", "ChatHeight"))
        Engine_Init_Grh .SkinGrh, Val(Engine_Var_Get(s, "CHATWINDOW", "Grh"))
    End With

    'Load Inventory
    With GameWindow.Inventory
        If LoadCustomPos Then
            .Screen.X = Val(Engine_Var_Get(t, "INVENTORY", "ScreenX"))
            .Screen.Y = Val(Engine_Var_Get(t, "INVENTORY", "ScreenY"))
        Else
            .Screen.X = Val(Engine_Var_Get(s, "INVENTORY", "ScreenX"))
            .Screen.Y = Val(Engine_Var_Get(s, "INVENTORY", "ScreenY"))
        End If
        .Screen.Width = Val(Engine_Var_Get(s, "INVENTORY", "ScreenWidth"))
        .Screen.Height = Val(Engine_Var_Get(s, "INVENTORY", "ScreenHeight"))
        Engine_Init_Grh .SkinGrh, Val(Engine_Var_Get(s, "INVENTORY", "Grh"))
    End With
    ImageOffsetX = Val(Engine_Var_Get(s, "INVENTORY", "ImageOffsetX"))
    ImageOffsetY = Val(Engine_Var_Get(s, "INVENTORY", "ImageOffsetY"))
    ImageSpaceX = Val(Engine_Var_Get(s, "INVENTORY", "ImageSpaceX"))
    ImageSpaceY = Val(Engine_Var_Get(s, "INVENTORY", "ImageSpaceY"))
    For LoopC = 1 To 49
        With GameWindow.Inventory.Image(LoopC)
            .X = ImageOffsetX + ((ImageSpaceX + 32) * (((LoopC - 1) Mod 7)))
            .Y = ImageOffsetY + ((ImageSpaceY + 32) * ((LoopC - 1) \ 7))
            .Width = 32
            .Height = 32
        End With
    Next LoopC

    'Load Shop window
    GameWindow.Shop = GameWindow.Inventory
    With GameWindow.Shop
        If LoadCustomPos Then
            .Screen.X = Val(Engine_Var_Get(t, "SHOP", "ScreenX"))
            .Screen.Y = Val(Engine_Var_Get(t, "SHOP", "ScreenY"))
        Else
            .Screen.X = Val(Engine_Var_Get(s, "SHOP", "ScreenX"))
            .Screen.Y = Val(Engine_Var_Get(s, "SHOP", "ScreenY"))
        End If
        Engine_Init_Grh .SkinGrh, Val(Engine_Var_Get(s, "SHOP", "Grh"))
    End With
    
    'Load bank window
    GameWindow.Bank = GameWindow.Inventory
    With GameWindow.Bank
        If LoadCustomPos Then
            .Screen.X = Val(Engine_Var_Get(t, "BANK", "ScreenX"))
            .Screen.Y = Val(Engine_Var_Get(t, "BANK", "ScreenY"))
        Else
            .Screen.X = Val(Engine_Var_Get(s, "BANK", "ScreenX"))
            .Screen.Y = Val(Engine_Var_Get(s, "BANK", "ScreenY"))
        End If
        Engine_Init_Grh .SkinGrh, Val(Engine_Var_Get(s, "BANK", "Grh"))
    End With

    'Load Mailbox window
    With GameWindow.Mailbox.Screen
        If LoadCustomPos Then
            .X = Val(Engine_Var_Get(t, "MAILBOX", "ScreenX"))
            .Y = Val(Engine_Var_Get(t, "MAILBOX", "ScreenY"))
        Else
            .X = Val(Engine_Var_Get(s, "MAILBOX", "ScreenX"))
            .Y = Val(Engine_Var_Get(s, "MAILBOX", "ScreenY"))
        End If
        .Width = Val(Engine_Var_Get(s, "MAILBOX", "ScreenWidth"))
        .Height = Val(Engine_Var_Get(s, "MAILBOX", "ScreenHeight"))
    End With
    Engine_Init_Grh GameWindow.Mailbox.SkinGrh, Val(Engine_Var_Get(s, "MAILBOX", "Grh"))
    With GameWindow.Mailbox.WriteLbl
        .X = Val(Engine_Var_Get(s, "MAILBOX", "WriteMessageX"))
        .Y = Val(Engine_Var_Get(s, "MAILBOX", "WriteMessageY"))
        .Width = Val(Engine_Var_Get(s, "MAILBOX", "WriteMessageWidth"))
        .Height = Val(Engine_Var_Get(s, "MAILBOX", "WriteMessageHeight"))
    End With
    With GameWindow.Mailbox.DeleteLbl
        .X = Val(Engine_Var_Get(s, "MAILBOX", "DeleteMessageX"))
        .Y = Val(Engine_Var_Get(s, "MAILBOX", "DeleteMessageY"))
        .Width = Val(Engine_Var_Get(s, "MAILBOX", "DeleteMessageWidth"))
        .Height = Val(Engine_Var_Get(s, "MAILBOX", "DeleteMessageHeight"))
    End With
    With GameWindow.Mailbox.ReadLbl
        .X = Val(Engine_Var_Get(s, "MAILBOX", "ReadMessageX"))
        .Y = Val(Engine_Var_Get(s, "MAILBOX", "ReadMessageY"))
        .Width = Val(Engine_Var_Get(s, "MAILBOX", "ReadMessageWidth"))
        .Height = Val(Engine_Var_Get(s, "MAILBOX", "ReadMessageHeight"))
    End With
    With GameWindow.Mailbox.List
        .X = Val(Engine_Var_Get(s, "MAILBOX", "ListX"))
        .Y = Val(Engine_Var_Get(s, "MAILBOX", "ListY"))
        .Width = Val(Engine_Var_Get(s, "MAILBOX", "ListWidth"))
        .Height = Val(Engine_Var_Get(s, "MAILBOX", "ListHeight"))
    End With

    'Load View Message window
    With GameWindow.ViewMessage.Screen
        If LoadCustomPos Then
            .X = Val(Engine_Var_Get(t, "VIEWMESSAGE", "ScreenX"))
            .Y = Val(Engine_Var_Get(t, "VIEWMESSAGE", "ScreenY"))
        Else
            .X = Val(Engine_Var_Get(s, "VIEWMESSAGE", "ScreenX"))
            .Y = Val(Engine_Var_Get(s, "VIEWMESSAGE", "ScreenY"))
        End If
        .Width = Val(Engine_Var_Get(s, "VIEWMESSAGE", "ScreenWidth"))
        .Height = Val(Engine_Var_Get(s, "VIEWMESSAGE", "ScreenHeight"))
    End With
    Engine_Init_Grh GameWindow.ViewMessage.SkinGrh, Val(Engine_Var_Get(s, "VIEWMESSAGE", "Grh"))
    With GameWindow.ViewMessage.From
        .X = Val(Engine_Var_Get(s, "VIEWMESSAGE", "FromX"))
        .Y = Val(Engine_Var_Get(s, "VIEWMESSAGE", "FromY"))
        .Width = Val(Engine_Var_Get(s, "VIEWMESSAGE", "FromWidth"))
        .Height = Val(Engine_Var_Get(s, "VIEWMESSAGE", "FromHeight"))
    End With
    With GameWindow.ViewMessage.Subject
        .X = Val(Engine_Var_Get(s, "VIEWMESSAGE", "SubjectX"))
        .Y = Val(Engine_Var_Get(s, "VIEWMESSAGE", "SubjectY"))
        .Width = Val(Engine_Var_Get(s, "VIEWMESSAGE", "SubjectWidth"))
        .Height = Val(Engine_Var_Get(s, "VIEWMESSAGE", "SubjectHeight"))
    End With
    With GameWindow.ViewMessage.Message
        .X = Val(Engine_Var_Get(s, "VIEWMESSAGE", "MessageX"))
        .Y = Val(Engine_Var_Get(s, "VIEWMESSAGE", "MessageY"))
        .Width = Val(Engine_Var_Get(s, "VIEWMESSAGE", "MessageWidth"))
        .Height = Val(Engine_Var_Get(s, "VIEWMESSAGE", "MessageHeight"))
    End With
    ImageOffsetX = Val(Engine_Var_Get(s, "VIEWMESSAGE", "ImageOffsetX"))
    ImageOffsetY = Val(Engine_Var_Get(s, "VIEWMESSAGE", "ImageOffsetY"))
    ImageSpaceX = Val(Engine_Var_Get(s, "VIEWMESSAGE", "ImageSpaceX"))
    For LoopC = 1 To MaxMailObjs
        With GameWindow.ViewMessage.Image(LoopC)
            .X = ImageOffsetX + ((LoopC - 1) * (ImageSpaceX + 32))
            .Y = ImageOffsetY
            .Width = 32
            .Height = 32
        End With
    Next LoopC

    'Load Write Message window
    GameWindow.WriteMessage = GameWindow.ViewMessage
    With GameWindow.ViewMessage.Screen
        If LoadCustomPos Then
            .X = Val(Engine_Var_Get(t, "WRITEMESSAGE", "ScreenX"))
            .Y = Val(Engine_Var_Get(t, "WRITEMESSAGE", "ScreenY"))
        Else
            .X = Val(Engine_Var_Get(s, "WRITEMESSAGE", "ScreenX"))
            .Y = Val(Engine_Var_Get(s, "WRITEMESSAGE", "ScreenY"))
        End If
    End With
    Engine_Init_Grh GameWindow.ViewMessage.SkinGrh, Val(Engine_Var_Get(s, "WRITEMESSAGE", "Grh"))

    'Load Amount window
    With GameWindow.Amount.Screen
        If LoadCustomPos Then
            .X = Val(Engine_Var_Get(t, "AMOUNT", "ScreenX"))
            .Y = Val(Engine_Var_Get(t, "AMOUNT", "ScreenY"))
        Else
            .X = Val(Engine_Var_Get(s, "AMOUNT", "ScreenX"))
            .Y = Val(Engine_Var_Get(s, "AMOUNT", "ScreenY"))
        End If
        .Width = Val(Engine_Var_Get(s, "AMOUNT", "ScreenWidth"))
        .Height = Val(Engine_Var_Get(s, "AMOUNT", "ScreenHeight"))
    End With
    Engine_Init_Grh GameWindow.Amount.SkinGrh, Val(Engine_Var_Get(s, "AMOUNT", "Grh"))
    With GameWindow.Amount.Value
        .X = Val(Engine_Var_Get(s, "AMOUNT", "ValueX"))
        .Y = Val(Engine_Var_Get(s, "AMOUNT", "ValueY"))
        .Width = Val(Engine_Var_Get(s, "AMOUNT", "ValueWidth"))
        .Height = Val(Engine_Var_Get(s, "AMOUNT", "ValueHeight"))
    End With

    'Load Menu Window
    With GameWindow.Menu.Screen
        If LoadCustomPos Then
            .X = Val(Engine_Var_Get(t, "MENU", "ScreenX"))
            .Y = Val(Engine_Var_Get(t, "MENU", "ScreenY"))
        Else
            .X = Val(Engine_Var_Get(s, "MENU", "ScreenX"))
            .Y = Val(Engine_Var_Get(s, "MENU", "ScreenY"))
        End If
        .Width = Val(Engine_Var_Get(s, "MENU", "ScreenWidth"))
        .Height = Val(Engine_Var_Get(s, "MENU", "ScreenHeight"))
    End With
    Engine_Init_Grh GameWindow.Menu.SkinGrh, Val(Engine_Var_Get(s, "MENU", "Grh"))
    With GameWindow.Menu.QuitLbl
        .X = Val(Engine_Var_Get(s, "MENU", "QuitX"))
        .Y = Val(Engine_Var_Get(s, "MENU", "QuitY"))
        .Width = Val(Engine_Var_Get(s, "MENU", "QuitWidth"))
        .Height = Val(Engine_Var_Get(s, "MENU", "QuitHeight"))
    End With
    
    'Reset text position
    If CurMap > 0 Then Engine_UpdateChatArray

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

Public Sub Engine_Init_NPCChat(ByVal Language As String)

'*****************************************************************
'Loads the NPC messages according to the language
'*****************************************************************
Dim Conditions() As NPCChatLineCondition
Dim NumConditions As Byte   'The number of conditions
Dim ConditionFlags As Long  'States what conditions are currently used (so we don't have to loop through the Conditions() array)
Dim ChatLine As Byte    'The chat line for the current index
Dim ErrTxt As String    'If there is an error, this extra text is added
Dim ChatIndex As Byte   'Index of the chat line
Dim HighIndex As Long   'Highest index retrieved
Dim Index As Long       'Current index
Dim FileNum As Byte
Dim ln As String        'Used to grab our lines
Dim Style As Byte       'Style used for the current index
Dim TempSplit() As String
Dim i As Long

    On Error GoTo ErrOut

    'Make sure the file exists
    If Not Engine_FileExist(DataPath & "NPC Chat\" & LCase$(Language) & ".ini", vbNormal) Then
        
        'Error! Change to English before we die!
        Language = "english"
    
    End If
    
    'Open the file
    FileNum = FreeFile
    Open DataPath & "NPC Chat\" & LCase$(Language) & ".ini" For Input Access Read As FileNum
        
        'Loop until we reach the BEGINFILE line, stating the data is going to start coming in
        Do
            Line Input #FileNum, ln
        Loop While UCase$(ln) <> "BEGINFILE"
        
        'Loop through the data
        Do
        
            'Get the line
            Line Input #FileNum, ln
            ln = Trim$(ln)
            
            'Look for empty lines
            If LenB(ln) = 0 Then GoTo NextLine
            
            '*** Look for a new index ***
            If Left$(ln, 1) = "[" Then
                
                'Grab the index
                Index = Mid$(ln, 2, Len(ln) - 2)
                
                'Clear the variables from the last line
                Style = 0
                ChatLine = 0
                Erase Conditions
                NumConditions = 0
                ConditionFlags = 0

                'Resize the chat array according to the index if needed
                If Index > HighIndex Then
                    ReDim Preserve NPCChat(1 To Index)
                    HighIndex = Index
                End If
                
                'Grab the format - this little loop will help us ignore blank lines
                Do
                    Line Input #FileNum, ln
                Loop While LenB(Trim$(ln)) = 0
                
                'Format text not found!
                If UCase$(Left$(ln, 6)) <> "FORMAT" Then
                    ErrTxt = "FORMAT not found immediately after index ([x]) tag!"
                    GoTo ErrOut
                End If
                
                'Figure out what format it is
                ln = Trim$(ln)
                Select Case UCase$(Right$(ln, Len(ln) - 7))
                    Case "RANDOM"
                        NPCChat(Index).Format = NPCCHAT_FORMAT_RANDOM
                    Case "LINEAR"
                        NPCChat(Index).Format = NPCCHAT_FORMAT_LINEAR
                    Case Else
                        ErrTxt = "Unknown format " & UCase$(Right$(ln, Len(ln) - 7)) & " retrieved!"
                        GoTo ErrOut
                End Select
                GoTo NextLine
                
            End If
            
            '*** Look for a new style ***
            If UCase$(Left$(ln, 6)) = "STYLE " Then
            
                'Figure out what style it is
                ln = Trim$(ln)
                Select Case UCase$(Right$(ln, Len(ln) - 6))
                    Case "BUBBLE"
                        Style = NPCCHAT_STYLE_BUBBLE
                    Case "BOX"
                        Style = NPCCHAT_STYLE_BOX
                    Case "BOTH"
                        Style = NPCCHAT_STYLE_BOTH
                    Case Else
                        ErrTxt = "Unknown style " & UCase$(Right$(ln, Len(ln) - 6)) & " retrieved!"
                        GoTo ErrOut
                End Select
                
            End If
            
            '*** Look for a new condition ***
            If Left$(ln, 1) = "!" Then
                
                'Figure out what condition it is
                ln = Trim$(ln)  'Trim off spaces
                TempSplit = Split(UCase$(Right$(ln, Len(ln) - 1)), " ") 'Remove the ! and turn to uppercase
                Select Case TempSplit(0)
                    Case "CLEAR"
                        Erase Conditions
                        NumConditions = 0
                        ConditionFlags = 0
                    Case "LEVELLESSTHAN"
                        If Not ConditionFlags And NPCCHAT_COND_LEVELLESSTHAN Then
                            NumConditions = NumConditions + 1
                            ReDim Preserve Conditions(1 To NumConditions)
                            Conditions(NumConditions).Condition = NPCCHAT_COND_LEVELLESSTHAN
                            Conditions(NumConditions).Value = Val(TempSplit(1))
                            ConditionFlags = ConditionFlags Or NPCCHAT_COND_LEVELLESSTHAN
                        End If
                    Case "LEVELMORETHAN"
                        If Not ConditionFlags And NPCCHAT_COND_LEVELMORETHAN Then
                            NumConditions = NumConditions + 1
                            ReDim Preserve Conditions(1 To NumConditions)
                            Conditions(NumConditions).Condition = NPCCHAT_COND_LEVELMORETHAN
                            Conditions(NumConditions).Value = Val(TempSplit(1))
                            ConditionFlags = ConditionFlags Or NPCCHAT_COND_LEVELMORETHAN
                        End If
                    Case "HPLESSTHAN"
                        If Not ConditionFlags And NPCCHAT_COND_HPLESSTHAN Then
                            NumConditions = NumConditions + 1
                            ReDim Preserve Conditions(1 To NumConditions)
                            Conditions(NumConditions).Condition = NPCCHAT_COND_HPLESSTHAN
                            Conditions(NumConditions).Value = Val(TempSplit(1))
                            ConditionFlags = ConditionFlags Or NPCCHAT_COND_HPLESSTHAN
                        End If
                    Case "HPMORETHAN"
                        If Not ConditionFlags And NPCCHAT_COND_HPMORETHAN Then
                            NumConditions = NumConditions + 1
                            ReDim Preserve Conditions(1 To NumConditions)
                            Conditions(NumConditions).Condition = NPCCHAT_COND_HPMORETHAN
                            Conditions(NumConditions).Value = Val(TempSplit(1))
                            ConditionFlags = ConditionFlags Or NPCCHAT_COND_HPMORETHAN
                        End If
                    Case "KNOWSKILL"
                        If Not ConditionFlags And NPCCHAT_COND_KNOWSKILL Then
                            NumConditions = NumConditions + 1
                            ReDim Preserve Conditions(1 To NumConditions)
                            Conditions(NumConditions).Condition = NPCCHAT_COND_KNOWSKILL
                            Conditions(NumConditions).Value = Val(TempSplit(1))
                            ConditionFlags = ConditionFlags Or NPCCHAT_COND_KNOWSKILL
                        End If
                    Case "DONTKNOWSKILL"
                        If Not ConditionFlags And NPCCHAT_COND_DONTKNOWSKILL Then
                            NumConditions = NumConditions + 1
                            ReDim Preserve Conditions(1 To NumConditions)
                            Conditions(NumConditions).Condition = NPCCHAT_COND_DONTKNOWSKILL
                            Conditions(NumConditions).Value = Val(TempSplit(1))
                            ConditionFlags = ConditionFlags Or NPCCHAT_COND_DONTKNOWSKILL
                        End If
                    Case "SAY"
                        If Not ConditionFlags And NPCCHAT_COND_SAY Then
                            NumConditions = NumConditions + 1
                            ReDim Preserve Conditions(1 To NumConditions)
                            Conditions(NumConditions).Condition = NPCCHAT_COND_SAY  'Notice we UCase$() the next line - this is so we can ignore the case
                            Conditions(NumConditions).ValueStr = UCase$(Replace$(TempSplit(1), "_", " "))   'Replace underscores with spaces
                            ConditionFlags = ConditionFlags Or NPCCHAT_COND_SAY
                        End If
                    Case Else
                        ErrTxt = "Unknown condition " & TempSplit(0) & " retrieved!"
                        GoTo ErrOut
                End Select
                
            End If
            
            '*** Look for a chat line ***
            If UCase$(Left$(ln, 4)) = "SAY " Then
                
                'Split up the information (0 = "SAY", 1 = Delay, 2 = Chat text)
                TempSplit() = Split(ln, " ", 3)
                
                'Raise the lines count
                ChatLine = ChatLine + 1
                ReDim Preserve NPCChat(Index).ChatLine(1 To ChatLine)
                NPCChat(Index).NumLines = ChatLine
                
                'Set the delay, style and text
                NPCChat(Index).ChatLine(ChatLine).Delay = Val(TempSplit(1))
                NPCChat(Index).ChatLine(ChatLine).Text = Trim$(TempSplit(2))
                NPCChat(Index).ChatLine(ChatLine).Style = Style
                
                'Check for empty text lines
                If UCase$(NPCChat(Index).ChatLine(ChatLine).Text) = "[EMPTY]" Then
                    NPCChat(Index).ChatLine(ChatLine).Text = vbNullString
                End If
                
                'Set the conditions
                NPCChat(Index).ChatLine(ChatLine).NumConditions = NumConditions
                If NumConditions > 0 Then
                    ReDim NPCChat(Index).ChatLine(ChatLine).Conditions(1 To NumConditions)
                    For i = 1 To NumConditions
                        NPCChat(Index).ChatLine(ChatLine).Conditions(i) = Conditions(i)
                    Next i
                End If
                
            End If

NextLine:
        
        Loop While Not EOF(FileNum)
    
    Close #FileNum
    
    Exit Sub
    
ErrOut:

    MsgBox "Error in NPCChat routine! Stopped on line " & Loc(FileNum) & "!" & vbNewLine & _
            "The remainder of the line text is: " & vbNewLine & ln & vbNewLine & vbNewLine & _
            "The following message has been added:" & vbNewLine & ErrTxt, vbOKOnly Or vbCritical
            
    If FileNum > 0 Then Close #FileNum
    
End Sub

Public Sub Engine_Init_Input()

'*****************************************************************
'Init Input Devices
'*****************************************************************

Dim diProp As DIPROPLONG
'Load the mouse input

    Set DI = DX.DirectInputCreate
    Set DIDevice = DI.CreateDevice("guid_SysMouse")
    Call DIDevice.SetCommonDataFormat(DIFORMAT_MOUSE)
    Call DIDevice.SetCooperativeLevel(frmMain.hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE)
    diProp.lHow = DIPH_DEVICE
    diProp.lObj = 0
    diProp.lData = BufferSize
    Call DIDevice.SetProperty("DIPROP_BUFFERSIZE", diProp)
    MouseEvent = DX.CreateEvent(frmMain)
    DIDevice.SetEventNotification MouseEvent

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
    D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
    'Particle engine settings
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0

    'Set the texture stage stats (filters)
    'D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    'D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR

End Sub

Sub Engine_Init_Texture(ByVal TextureNum As Integer)

'*****************************************************************
'Loads a texture into memory
'*****************************************************************
Dim TexInfo As D3DXIMAGE_INFO_A
Dim FilePath As String

    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Make sure we try not to load a file while the engine is unloading
    If IsUnloading Then Exit Sub

    'Get the path
    FilePath = GrhPath & TextureNum & ".png"
    
    'Check if the texture exists
    If Engine_FileExist(FilePath, vbNormal) = False Then
        MsgBox "Error! Could not find the following texture file:" & vbNewLine & FilePath, vbOKOnly
        IsUnloading = 1
        Exit Sub
    End If

    If SurfaceSize(TextureNum).X = 0 Then   'We need to get the size
    
        'Set the texture (and get the dimensions)
        Set SurfaceDB(TextureNum) = D3DX.CreateTextureFromFileEx(D3DDevice, FilePath, D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFF000000, TexInfo, ByVal 0)
        SurfaceSize(TextureNum).X = TexInfo.Width
        SurfaceSize(TextureNum).Y = TexInfo.Height
        SurfaceSize(TextureNum).MipLevels = TexInfo.MipLevels
        SurfaceSize(TextureNum).BmpFormat = TexInfo.Format
        
    Else
        
        'Set the texture (without getting the dimensions)
        Set SurfaceDB(TextureNum) = D3DX.CreateTextureFromFileEx(D3DDevice, FilePath, SurfaceSize(TextureNum).X, SurfaceSize(TextureNum).Y, SurfaceSize(TextureNum).MipLevels, 0, SurfaceSize(TextureNum).BmpFormat, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)
    
    End If

    'Set the texture timer
    SurfaceTimer(TextureNum) = SurfaceTimerMax

End Sub

Sub Engine_Init_FontTextures()

'*****************************************************************
'Init the custom font textures
'*****************************************************************
Dim FileNum As Byte

    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    '*** Default font ***
    Set Font_Default.Texture = D3DX.CreateTextureFromFileEx(D3DDevice, DataPath & "texdefault.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)

End Sub

Sub Engine_Init_FontSettings()

'*****************************************************************
'Init the custom font settings
'*****************************************************************
Dim FileNum As Byte
    
    '*** Default font ***

    'Load the header information
    FileNum = FreeFile
    Open DataPath & "texdefault.dat" For Binary As #FileNum
        Get #FileNum, , Font_Default.HeaderInfo
    Close #FileNum
    
    'Calculate some common values
    Font_Default.CharHeight = Font_Default.HeaderInfo.CellHeight - 4
    Font_Default.RowPitch = Font_Default.HeaderInfo.BitmapWidth \ Font_Default.HeaderInfo.CellWidth
    Font_Default.ColFactor = Font_Default.HeaderInfo.CellWidth / Font_Default.HeaderInfo.BitmapWidth
    Font_Default.RowFactor = Font_Default.HeaderInfo.CellHeight / Font_Default.HeaderInfo.BitmapHeight

End Sub

Function Engine_Init_TileEngine() As Boolean

'*****************************************************************
'Init Tile Engine
'*****************************************************************

    '****** INIT DirectX ******
    ' Create the root D3D objects
    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate()
    Set D3DX = New D3DX8
    Engine_Init_Input
    Engine_Init_Sound

    'Create the D3D Device
    If Engine_Init_D3DDevice(D3DCREATE_PUREDEVICE) = 0 Then
        If Engine_Init_D3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) = 0 Then
            If Engine_Init_D3DDevice(D3DCREATE_MIXED_VERTEXPROCESSING) = 0 Then
                If Engine_Init_D3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) = 0 Then
                    If Engine_Init_D3DDevice(D3DCREATE_FPU_PRESERVE) = 0 Then
                        If Engine_Init_D3DDevice(D3DCREATE_MULTITHREADED) = 0 Then
                            MsgBox "Could not init D3DDevice. Exiting..."
                            Engine_Init_UnloadTileEngine
                            Engine_UnloadAllForms
                            End
                        End If
                    End If
                End If
            End If
        End If
    End If
    Engine_Init_RenderStates
    
    'Load the rest of the tile engine stuff
    Engine_Init_FontTextures
    Engine_Init_ParticleEngine

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
Dim X As Long
Dim Y As Long

    EngineRun = False
    
    '****** Clear DirectX objects ******
    If Not DIDevice Is Nothing Then DIDevice.Unacquire
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    If Not DIDevice Is Nothing Then Set DIDevice = Nothing
    If Not D3DX Is Nothing Then Set D3DX = Nothing
    If Not DI Is Nothing Then Set DI = Nothing

    'Clear particles
    For LoopC = 1 To UBound(ParticleTexture)
        If Not ParticleTexture(LoopC) Is Nothing Then Set ParticleTexture(LoopC) = Nothing
    Next LoopC

    'Clear GRH memory
    For LoopC = 1 To NumGrhFiles
        If Not SurfaceDB(LoopC) Is Nothing Then Set SurfaceDB(LoopC) = Nothing
    Next LoopC
    
    'Clear sound buffers
    For LoopC = 1 To NumSfx
        If Not DSBuffer(LoopC) Is Nothing Then Set DSBuffer(LoopC) = Nothing
    Next LoopC
    
    'Clear map sound buffers
    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            If Not MapData(X, Y).Sfx Is Nothing Then Set MapData(X, Y).Sfx = Nothing
        Next Y
    Next X

    'Clear music objects
    For LoopC = 1 To NumMusicBuffers
        If Not DirectShow_Position(LoopC) Is Nothing Then Set DirectShow_Position(LoopC) = Nothing
        If Not DirectShow_Control(LoopC) Is Nothing Then Set DirectShow_Control(LoopC) = Nothing
        If Not DirectShow_Event(LoopC) Is Nothing Then Set DirectShow_Event(LoopC) = Nothing
        If Not DirectShow_Audio(LoopC) Is Nothing Then Set DirectShow_Audio(LoopC) = Nothing
    Next LoopC
    
    'Clear arrays
    Erase SurfaceTimer
    Erase SoundBufferTimer
    Erase VertexArray
    Erase MapData
    Erase GrhData
    Erase GrhData
    Erase SurfaceSize
    Erase BodyData
    Erase HeadData
    Erase WeaponData
    Erase MapData
    Erase CharList
    Erase OBJList
    Erase BloodList
    Erase EffectList
    Erase DamageList
    Erase SkillList
    Erase QuickBarID
    Erase ShowGameWindow
    Erase SaveLightBuffer
    
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
Static LastWeather As Byte
Dim X As Byte
Dim Y As Byte
Dim i As Byte

    'Only update the weather settings if it has changed!
    If LastWeather <> MapInfo.Weather Then
    
        'Set the lastweather to the current weather
        LastWeather = MapInfo.Weather
        
        'Erase sounds
        Engine_Sound_Erase WeatherSfx1
        Engine_Sound_Erase WeatherSfx2
    
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
            DoLightning = 0
            
        Case 2  'Rain Storm (heavy rain + lightning)
            If WeatherEffectIndex <= 0 Then
                WeatherEffectIndex = Effect_Rain_Begin(9, 300)
            ElseIf Effect(WeatherEffectIndex).EffectNum <> EffectNum_Rain Then
                Effect_Kill WeatherEffectIndex
                WeatherEffectIndex = Effect_Rain_Begin(9, 300)
            ElseIf Not Effect(WeatherEffectIndex).Used Then
                WeatherEffectIndex = Effect_Rain_Begin(9, 300)
            End If
            DoLightning = 1 'We take our rain with a bit of lightning on top >:D
            Engine_Sound_Set WeatherSfx1, 3
            Engine_Sound_Set WeatherSfx2, 2
            Engine_Sound_Play WeatherSfx1, DSBPLAY_LOOPING
            
        Case 3  'Inside of a cave/house in a storm (lightning + cave rain sound)
            If WeatherEffectIndex > 0 Then  'Kill the weather effect if used
                If Effect(WeatherEffectIndex).Used Then Effect_Kill WeatherEffectIndex
            End If
            DoLightning = 1
            Engine_Sound_Set WeatherSfx1, 4
            Engine_Sound_Set WeatherSfx2, 6
            Engine_Sound_Play WeatherSfx1, DSBPLAY_LOOPING
            
        Case Else   'None
            If WeatherEffectIndex > 0 Then  'Kill the weather effect if used
                If Effect(WeatherEffectIndex).Used Then Effect_Kill WeatherEffectIndex
                Engine_Sound_Erase WeatherSfx1  'Remove the sounds
                Engine_Sound_Erase WeatherSfx2
            End If
            
        End Select
        
    End If
    
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
                        For i = 1 To 24
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
                
                'Randomly place the lightning
                LightningX = 50 + Rnd * 700
                LightningY = Rnd * -200
                Engine_Sound_Play WeatherSfx2, DSBPLAY_DEFAULT  'BAM!
                
                'Change the light of all the tiles to white
                For X = XMinMapSize To XMaxMapSize
                    For Y = YMinMapSize To YMaxMapSize
                        For i = 1 To 24
                            MapData(X, Y).Light(i) = -1
                        Next i
                    Next Y
                Next X
                
            End If
            
        End If
        
    End If

End Sub

Sub Engine_Input_CheckKeys()

'*****************************************************************
'Checks keys and respond
'*****************************************************************
    
    If DisableInput = 1 Then Exit Sub
    
    'Dont move when Control is pressed
    If GetAsyncKeyState(vbKeyControl) Then Exit Sub

    'Check if certain screens are open that require ASDW keys
    If ShowGameWindow(WriteMessageWindow) Then
        If WMSelCon <> 0 Then Exit Sub
    End If

    'Zoom in / out
    If GetAsyncKeyState(vbKeyNumpad8) Then      'In
        ZoomLevel = ZoomLevel + (ElapsedTime * 0.0003)
        If ZoomLevel > MaxZoomLevel Then ZoomLevel = MaxZoomLevel
    ElseIf GetAsyncKeyState(vbKeyNumpad2) Then  'Out
        ZoomLevel = ZoomLevel - (ElapsedTime * 0.0003)
        If ZoomLevel < 0 Then ZoomLevel = 0
    End If

    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If GetAsyncKeyState(vbKeyTab) Then
            'Move Up-Right
            If GetKeyState(vbKeyUp) < 0 And GetKeyState(vbKeyRight) < 0 Then
                Engine_ChangeHeading NORTHEAST
                Exit Sub
            End If
            'Move Up-Left
            If GetKeyState(vbKeyUp) < 0 And GetKeyState(vbKeyLeft) < 0 Then
                Engine_ChangeHeading NORTHWEST
                Exit Sub
            End If
            'Move Down-Right
            If GetKeyState(vbKeyDown) < 0 And GetKeyState(vbKeyRight) < 0 Then
                Engine_ChangeHeading SOUTHEAST
                Exit Sub
            End If
            'Move Down-Left
            If GetKeyState(vbKeyDown) < 0 And GetKeyState(vbKeyLeft) < 0 Then
                Engine_ChangeHeading SOUTHWEST
                Exit Sub
            End If
            'Move Up
            If GetKeyState(vbKeyUp) < 0 Then
                Engine_ChangeHeading NORTH
                Exit Sub
            End If
            'Move Right
            If GetKeyState(vbKeyRight) < 0 Then
                Engine_ChangeHeading EAST
                Exit Sub
            End If
            'Move down
            If GetKeyState(vbKeyDown) < 0 Then
                Engine_ChangeHeading SOUTH
                Exit Sub
            End If
            'Move left
            If GetKeyState(vbKeyLeft) < 0 Then
                Engine_ChangeHeading WEST
                Exit Sub
            End If
            If EnterText = False Then
                If GetKeyState(vbKeyW) < 0 And GetKeyState(vbKeyD) < 0 Then
                    Engine_ChangeHeading NORTHEAST
                    Exit Sub
                End If
                If GetKeyState(vbKeyW) < 0 And GetKeyState(vbKeyA) < 0 Then
                    Engine_ChangeHeading NORTHWEST
                    Exit Sub
                End If
                If GetKeyState(vbKeyS) < 0 And GetKeyState(vbKeyD) < 0 Then
                    Engine_ChangeHeading SOUTHEAST
                    Exit Sub
                End If
                If GetKeyState(vbKeyS) < 0 And GetKeyState(vbKeyA) < 0 Then
                    Engine_ChangeHeading SOUTHWEST
                    Exit Sub
                End If
                If GetKeyState(vbKeyW) < 0 Then
                    Engine_ChangeHeading NORTH
                    Exit Sub
                End If
                If GetKeyState(vbKeyD) < 0 Then
                    Engine_ChangeHeading EAST
                    Exit Sub
                End If
                If GetKeyState(vbKeyS) < 0 Then
                    Engine_ChangeHeading SOUTH
                    Exit Sub
                End If
                If GetKeyState(vbKeyA) < 0 Then
                    Engine_ChangeHeading WEST
                    Exit Sub
                End If
            End If
        Else
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
            If EnterText = False Then
                If GetKeyState(vbKeyW) < 0 And GetKeyState(vbKeyD) < 0 Then
                    Engine_MoveUser NORTHEAST
                    Exit Sub
                End If
                If GetKeyState(vbKeyW) < 0 And GetKeyState(vbKeyA) < 0 Then
                    Engine_MoveUser NORTHWEST
                    Exit Sub
                End If
                If GetKeyState(vbKeyS) < 0 And GetKeyState(vbKeyD) < 0 Then
                    Engine_MoveUser SOUTHEAST
                    Exit Sub
                End If
                If GetKeyState(vbKeyS) < 0 And GetKeyState(vbKeyA) < 0 Then
                    Engine_MoveUser SOUTHWEST
                    Exit Sub
                End If
                If GetKeyState(vbKeyW) < 0 Then
                    Engine_MoveUser NORTH
                    Exit Sub
                End If
                If GetKeyState(vbKeyD) < 0 Then
                    Engine_MoveUser EAST
                    Exit Sub
                End If
                If GetKeyState(vbKeyS) < 0 Then
                    Engine_MoveUser SOUTH
                    Exit Sub
                End If
                If GetKeyState(vbKeyA) < 0 Then
                    Engine_MoveUser WEST
                    Exit Sub
                End If
            End If
        End If
    End If

End Sub

Sub Engine_Input_Mouse_LeftClick()

'******************************************
'Left click mouse
'******************************************

Dim tX As Integer
Dim tY As Integer

'Make sure engine is running

    If Not EngineRun Then Exit Sub

    '***Check for skill list click***
    'Skill lists, because it is not actually a window, must be handled differently
    If QuickBarSetSlot <= 0 Then DrawSkillList = 0
    If DrawSkillList Then
        If SkillListSize Then
            For tX = 1 To SkillListSize
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, SkillList(tX).X, SkillList(tX).Y, 32, 32) Then
                    QuickBarID(QuickBarSetSlot).ID = SkillList(tX).SkillID
                    QuickBarID(QuickBarSetSlot).Type = QuickBarType_Skill
                    DrawSkillList = 0
                    QuickBarSetSlot = 0
                    Exit Sub
                End If
            Next tX
        End If
    End If

    '***Check for a window click***
    WMSelCon = 0

    'Start with the last clicked window, then move in order of importance
    If Engine_Input_Mouse_LeftClick_Window(LastClickedWindow) = 0 Then
        If Engine_Input_Mouse_LeftClick_Window(ChatWindow) = 0 Then
            If Engine_Input_Mouse_LeftClick_Window(QuickBarWindow) = 0 Then
                If Engine_Input_Mouse_LeftClick_Window(MenuWindow) = 0 Then
                    If Engine_Input_Mouse_LeftClick_Window(InventoryWindow) = 0 Then
                        If Engine_Input_Mouse_LeftClick_Window(ShopWindow) = 0 Then
                            If Engine_Input_Mouse_LeftClick_Window(BankWindow) = 0 Then
                                If Engine_Input_Mouse_LeftClick_Window(StatWindow) = 0 Then
                                    If Engine_Input_Mouse_LeftClick_Window(MailboxWindow) = 0 Then
                                        If Engine_Input_Mouse_LeftClick_Window(ViewMessageWindow) = 0 Then
                                            If Engine_Input_Mouse_LeftClick_Window(WriteMessageWindow) = 0 Then
                                                If Engine_Input_Mouse_LeftClick_Window(AmountWindow) = 0 Then

                                                    'No windows clicked, so a tile click will take place
                                                    'Get the tile positions
                                                    Engine_ConvertCPtoTP 0, 0, MousePos.X, MousePos.Y, tX, tY
            
                                                    'Send left click
                                                    sndBuf.Allocate 3
                                                    sndBuf.Put_Byte DataCode.User_LeftClick
                                                    sndBuf.Put_Byte CByte(tX)
                                                    sndBuf.Put_Byte CByte(tY)
            
                                                    'If there was a click on the game screen and the
                                                    ' skill list is up, but no window clicked, set to 0
                                                    If DrawSkillList Then
                                                        If QuickBarSetSlot Then
                                                            QuickBarID(QuickBarSetSlot).ID = 0
                                                            QuickBarID(QuickBarSetSlot).Type = 0
                                                            DrawSkillList = 0
                                                            QuickBarSetSlot = 0
                                                        End If
                                                    End If
                                                    
                                                    'Last clicked window was nothing, so set to nothing :)
                                                    LastClickedWindow = 0
                                                
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

End Sub

Function Engine_Input_Mouse_LeftClick_Window(ByVal WindowIndex As Byte) As Byte

'******************************************
'Left click a game window
'******************************************

Dim i As Byte
Dim j As Byte

    Select Case WindowIndex
    Case MenuWindow
        If ShowGameWindow(MenuWindow) Then
            With GameWindow.Menu
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_LeftClick_Window = 1
                    LastClickedWindow = MenuWindow
                    'Quit button
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .QuitLbl.X, .Screen.Y + .QuitLbl.Y, .QuitLbl.Width, .QuitLbl.Height) Then
                        IsUnloading = 1
                        Exit Function
                    End If
                    SelGameWindow = MenuWindow
                End If
            End With
        End If
        
    Case StatWindow
        If ShowGameWindow(StatWindow) Then
            With GameWindow.StatWindow
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_LeftClick_Window = 1
                    LastClickedWindow = StatWindow
                    'Raise str
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .AddStr.X, .Screen.Y + .AddStr.Y, .AddStr.Width, .AddStr.Height) Then
                        sndBuf.Allocate 2
                        sndBuf.Put_Byte DataCode.User_BaseStat
                        sndBuf.Put_Byte SID.Str
                    End If
                    'Raise agi
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .AddAgi.X, .Screen.Y + .AddAgi.Y, .AddAgi.Width, .AddAgi.Height) Then
                        sndBuf.Allocate 2
                        sndBuf.Put_Byte DataCode.User_BaseStat
                        sndBuf.Put_Byte SID.Agi
                    End If
                    'Raise mag
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .AddMag.X, .Screen.Y + .AddMag.Y, .AddMag.Width, .AddMag.Height) Then
                        sndBuf.Allocate 2
                        sndBuf.Put_Byte DataCode.User_BaseStat
                        sndBuf.Put_Byte SID.Mag
                    End If
                    SelGameWindow = StatWindow
                End If
            End With
        End If
        
    Case ChatWindow
        If ShowGameWindow(ChatWindow) Then
            With GameWindow.ChatWindow
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Text.X, .Screen.Y + .Text.Y, .Text.Width, .Text.Height) Then
                        EnterText = True
                    End If
                    Engine_Input_Mouse_LeftClick_Window = 1
                    LastClickedWindow = ChatWindow
                    SelGameWindow = ChatWindow
                    Exit Function
                End If
            End With
        End If
    
    Case QuickBarWindow
        If ShowGameWindow(QuickBarWindow) Then
            With GameWindow.QuickBar
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_LeftClick_Window = 1
                    LastClickedWindow = QuickBarWindow
                    'Cancel changes to quick bar items
                    DrawSkillList = 0
                    QuickBarSetSlot = 0
                    'Check if an item was clicked
                    For i = 1 To 12
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                            If GetAsyncKeyState(vbKeyShift) Then
                                QuickBarSetSlot = i
                                DrawSkillList = 1
                            Else
                                Engine_UseQuickBar i
                            End If
                            Exit Function
                        End If
                    Next i
                    'Item was not clicked
                    SelGameWindow = QuickBarWindow
                    Exit Function
                End If
            End With
        End If
        
    Case InventoryWindow
        If ShowGameWindow(InventoryWindow) Then
            With GameWindow.Inventory
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_LeftClick_Window = 1
                    LastClickedWindow = InventoryWindow
                    'Check if an item was clicked
                    For i = 1 To 49
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                            If GetAsyncKeyState(vbKeyShift) Then
                                If Game_ClickItem(i) Then
                                    If UserInventory(i).Amount = 1 Then
                                        'Drop item into mailbox
                                        If ShowGameWindow(WriteMessageWindow) Then
                                            'Check for duplicate entries
                                            For j = 1 To MaxMailObjs
                                                If WriteMailData.ObjIndex(j) = i Then Exit Function
                                            Next j
                                            'Place item in next free slot (if any)
                                            j = 0
                                            Do
                                                j = j + 1
                                                If j > MaxMailObjs Then Exit Function
                                            Loop While WriteMailData.ObjIndex(j) > 0
                                            WriteMailData.ObjIndex(j) = i
                                            WriteMailData.ObjAmount(j) = 1
                                        'Sell item to shopkeeper
                                        ElseIf ShowGameWindow(ShopWindow) Then
                                            sndBuf.Allocate 4
                                            sndBuf.Put_Byte DataCode.User_Trade_SellToNPC
                                            sndBuf.Put_Byte i
                                            sndBuf.Put_Integer 1
                                        'Put item in the bank
                                        ElseIf ShowGameWindow(BankWindow) Then
                                            sndBuf.Allocate 4
                                            sndBuf.Put_Byte DataCode.User_Bank_PutItem
                                            sndBuf.Put_Byte i
                                            sndBuf.Put_Integer 1
                                        'Drop item on ground
                                        Else
                                            sndBuf.Allocate 4
                                            sndBuf.Put_Byte DataCode.User_Drop
                                            sndBuf.Put_Byte i
                                            sndBuf.Put_Integer 1
                                        End If
                                    Else
                                        'Drop item into mailbox
                                        If ShowGameWindow(WriteMessageWindow) Then
                                            'Check for duplicate entries
                                            For j = 1 To MaxMailObjs
                                                If WriteMailData.ObjIndex(j) = i Then Exit Function
                                            Next j
                                            'Check for free slots
                                            j = 0
                                            Do
                                                j = j + 1
                                                If j > MaxMailObjs Then Exit Function
                                            Loop While WriteMailData.ObjIndex(j) > 0
                                            'Open the amount window
                                            ShowGameWindow(AmountWindow) = 1
                                            LastClickedWindow = AmountWindow
                                            AmountWindowValue = vbNullString
                                            AmountWindowItemIndex = i
                                            AmountWindowUsage = AW_InvToMail
                                        'Sell item to shopkeeper
                                        ElseIf ShowGameWindow(ShopWindow) Then
                                            ShowGameWindow(AmountWindow) = 1
                                            LastClickedWindow = AmountWindow
                                            AmountWindowValue = vbNullString
                                            AmountWindowItemIndex = i
                                            AmountWindowUsage = AW_InvToShop
                                        'Put item in the bank
                                        ElseIf ShowGameWindow(BankWindow) Then
                                            ShowGameWindow(AmountWindow) = 1
                                            LastClickedWindow = AmountWindow
                                            AmountWindowValue = vbNullString
                                            AmountWindowItemIndex = i
                                            AmountWindowUsage = AW_InvToBank
                                        'Drop item on ground
                                        Else
                                            ShowGameWindow(AmountWindow) = 1
                                            LastClickedWindow = AmountWindow
                                            AmountWindowValue = vbNullString
                                            AmountWindowItemIndex = i
                                            AmountWindowUsage = AW_Drop
                                        End If
                                    End If
                                End If
                            Else
                                If Game_ClickItem(i) Then
                                    sndBuf.Allocate 2
                                    sndBuf.Put_Byte DataCode.User_Use
                                    sndBuf.Put_Byte i
                                End If
                            End If
                            Exit Function
                        End If
                    Next i
                    'Item was not clicked
                    SelGameWindow = InventoryWindow
                    Exit Function
                End If
            End With
        End If
        
    Case ShopWindow
        If ShowGameWindow(ShopWindow) Then
            With GameWindow.Shop
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_LeftClick_Window = 1
                    LastClickedWindow = ShopWindow
                    'Check if an item was clicked
                    For i = 1 To 49
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                            If Game_ClickItem(i, 2) > 0 Then
                                sndBuf.Allocate 4
                                sndBuf.Put_Byte DataCode.User_Trade_BuyFromNPC
                                sndBuf.Put_Byte i
                                sndBuf.Put_Integer 1
                            End If
                            Exit Function
                        End If
                    Next i
                    'Item was not clicked
                    SelGameWindow = ShopWindow
                    Exit Function
                End If
            End With
        End If
        
    Case BankWindow
        If ShowGameWindow(BankWindow) Then
            With GameWindow.Bank
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_LeftClick_Window = 1
                    LastClickedWindow = BankWindow
                    'Check if an item was clicked
                    For i = 1 To 49
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                            If Game_ClickItem(i, 3) > 0 Then
                                sndBuf.Allocate 4
                                sndBuf.Put_Byte DataCode.User_Bank_TakeItem
                                sndBuf.Put_Byte i
                                sndBuf.Put_Integer 1
                            End If
                            Exit Function
                        End If
                    Next i
                    'Item was not clicked
                    SelGameWindow = BankWindow
                    Exit Function
                End If
            End With
        End If
        
    Case MailboxWindow
        If ShowGameWindow(MailboxWindow) Then
            With GameWindow.Mailbox
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_LeftClick_Window = 1
                    LastClickedWindow = MailboxWindow
                    'Check if Write was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .WriteLbl.X, .Screen.Y + .WriteLbl.Y, .WriteLbl.Width, .WriteLbl.Height) Then
                        For i = 1 To MaxMailObjs
                            WriteMailData.ObjIndex(i) = 0
                            WriteMailData.ObjAmount(i) = 0
                        Next i
                        WriteMailData.Message = vbNullString
                        WriteMailData.Subject = vbNullString
                        WriteMailData.RecieverName = vbNullString
                        ShowGameWindow(MailboxWindow) = 0
                        ShowGameWindow(WriteMessageWindow) = 1
                        LastClickedWindow = WriteMessageWindow
                        Exit Function
                    End If
                    If SelMessage > 0 Then
                        'Check if Delete was clicked
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .DeleteLbl.X, .Screen.Y + .DeleteLbl.Y, .DeleteLbl.Width, .DeleteLbl.Height) Then
                            sndBuf.Allocate 2
                            sndBuf.Put_Byte DataCode.Server_MailDelete
                            sndBuf.Put_Byte SelMessage
                            Exit Function
                        End If
                        'Check if Read was clicked
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .ReadLbl.X, .Screen.Y + .ReadLbl.Y, .ReadLbl.Width, .ReadLbl.Height) Then
                            sndBuf.Allocate 2
                            sndBuf.Put_Byte DataCode.Server_MailMessage
                            sndBuf.Put_Byte SelMessage
                            Exit Function
                        End If
                    End If
                    'Check if List was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .List.X + .List.X, .Screen.Y + .List.Y, .List.Width, .List.Height) Then
                        For i = 1 To (.List.Height \ Font_Default.CharHeight)
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .List.X + .List.X, .Screen.Y + .List.Y + ((i - 1) * Font_Default.CharHeight), .List.Width, Font_Default.CharHeight) Then
                                If SelMessage = i Then
                                    sndBuf.Allocate 2
                                    sndBuf.Put_Byte DataCode.Server_MailMessage
                                    sndBuf.Put_Byte i
                                Else
                                    SelMessage = i
                                End If
                                Exit Function
                            End If
                        Next i
                        Exit Function
                    End If
                    SelGameWindow = MailboxWindow
                    Exit Function
                End If
            End With
        End If
        
    Case ViewMessageWindow
        If ShowGameWindow(ViewMessageWindow) Then
            With GameWindow.ViewMessage
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_LeftClick_Window = 1
                    LastClickedWindow = ViewMessageWindow
                    'Click an item
                    For i = 1 To MaxMailObjs
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Image(i).X, .Screen.Y + .Image(i).Y, .Image(i).Width, .Image(i).Height) Then
                            sndBuf.Allocate 2
                            sndBuf.Put_Byte DataCode.Server_MailItemTake
                            sndBuf.Put_Byte i
                            Exit Function
                        End If
                    Next i
                    SelGameWindow = ViewMessageWindow
                    Exit Function
                End If
            End With
        End If
        
    Case WriteMessageWindow
        If ShowGameWindow(WriteMessageWindow) Then
            With GameWindow.WriteMessage
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_LeftClick_Window = 1
                    LastClickedWindow = WriteMessageWindow
                    'Click From
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .From.X + .Screen.X, .From.Y + .Screen.Y, .From.Width, .From.Height) Then
                        WMSelCon = wmFrom
                        Exit Function
                    End If
                    'Click Subject
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Subject.X + .Screen.X, .Subject.Y + .Screen.Y, .Subject.Width, .Subject.Height) Then
                        WMSelCon = wmSubject
                        Exit Function
                    End If
                    'Click Message
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Message.X + .Screen.X, .Message.Y + .Screen.Y, .Message.Width, .Message.Height) Then
                        WMSelCon = wmMessage
                        Exit Function
                    End If
                    'Click an item
                    For i = 1 To MaxMailObjs
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Image(i).X, .Screen.Y + .Image(i).Y, .Image(i).Width, .Image(i).Height) Then
                            WriteMailData.ObjIndex(i) = 0
                            WriteMailData.ObjAmount(i) = 0
                            Exit Function
                        End If
                    Next i
                    SelGameWindow = WriteMessageWindow
                    Exit Function
                End If
            End With
        End If
        
    Case AmountWindow
        If ShowGameWindow(AmountWindow) Then
            With GameWindow.Amount
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_LeftClick_Window = 1
                    LastClickedWindow = AmountWindow
                End If
                SelGameWindow = AmountWindow
                Exit Function
            End With
        End If
        
    End Select

End Function

Sub Engine_Input_Mouse_Move()

'******************************************
'Move mouse
'******************************************

Dim tX As Integer
Dim tY As Integer

'Make sure engine is running

    If Not EngineRun Then Exit Sub

    'Clear item info display
    ItemDescLines = 0

    'Check if left mouse is pressed
    If MouseLeftDown Then

        'Move QuickBar
        If SelGameWindow = QuickBarWindow Then
            With GameWindow.QuickBar.Screen
                .X = .X + MousePosAdd.X
                .Y = .Y + MousePosAdd.Y
                If WindowsInScreen Then
                    If .X < 0 Then .X = 0
                    If .Y < 0 Then .Y = 0
                    If .X > 800 - .Width Then .X = 800 - .Width
                    If .Y > 600 - .Height Then .Y = 600 - .Height
                End If
            End With
            'Move ChatWindow
        ElseIf SelGameWindow = ChatWindow Then
            With GameWindow.ChatWindow.Screen
                .X = .X + MousePosAdd.X
                .Y = .Y + MousePosAdd.Y
                If WindowsInScreen Then
                    If .X < 0 Then .X = 0
                    If .Y < 0 Then .Y = 0
                    If .X > 800 - .Width Then .X = 800 - .Width
                    If .Y > 600 - .Height Then .Y = 600 - .Height
                End If
                Engine_UpdateChatArray
            End With
            'Move Stat Window
        ElseIf SelGameWindow = StatWindow Then
            With GameWindow.StatWindow.Screen
                .X = .X + MousePosAdd.X
                .Y = .Y + MousePosAdd.Y
                If WindowsInScreen Then
                    If .X < 0 Then .X = 0
                    If .Y < 0 Then .Y = 0
                    If .X > 800 - .Width Then .X = 800 - .Width
                    If .Y > 600 - .Height Then .Y = 600 - .Height
                End If
            End With
            'Move Inventory
        ElseIf SelGameWindow = InventoryWindow Then
            With GameWindow.Inventory.Screen
                .X = .X + MousePosAdd.X
                .Y = .Y + MousePosAdd.Y
                If WindowsInScreen Then
                    If .X < 0 Then .X = 0
                    If .Y < 0 Then .Y = 0
                    If .X > 800 - .Width Then .X = 800 - .Width
                    If .Y > 600 - .Height Then .Y = 600 - .Height
                End If
            End With
            'Move Shop
        ElseIf SelGameWindow = ShopWindow Then
            With GameWindow.Shop.Screen
                .X = .X + MousePosAdd.X
                .Y = .Y + MousePosAdd.Y
                If WindowsInScreen Then
                    If .X < 0 Then .X = 0
                    If .Y < 0 Then .Y = 0
                    If .X > 800 - .Width Then .X = 800 - .Width
                    If .Y > 600 - .Height Then .Y = 600 - .Height
                End If
            End With
            'Move Bank
        ElseIf SelGameWindow = BankWindow Then
            With GameWindow.Bank.Screen
                .X = .X + MousePosAdd.X
                .Y = .Y + MousePosAdd.Y
                If WindowsInScreen Then
                    If .X < 0 Then .X = 0
                    If .Y < 0 Then .Y = 0
                    If .X > 800 - .Width Then .X = 800 - .Width
                    If .Y > 600 - .Height Then .Y = 600 - .Height
                End If
            End With
            'Move Mailbox
        ElseIf SelGameWindow = MailboxWindow Then
            With GameWindow.Mailbox.Screen
                .X = .X + MousePosAdd.X
                .Y = .Y + MousePosAdd.Y
                If WindowsInScreen Then
                    If .X < 0 Then .X = 0
                    If .Y < 0 Then .Y = 0
                    If .X > 800 - .Width Then .X = 800 - .Width
                    If .Y > 600 - .Height Then .Y = 600 - .Height
                End If
            End With
            'Move View Message
        ElseIf SelGameWindow = ViewMessageWindow Then
            With GameWindow.ViewMessage.Screen
                .X = .X + MousePosAdd.X
                .Y = .Y + MousePosAdd.Y
                If WindowsInScreen Then
                    If .X < 0 Then .X = 0
                    If .Y < 0 Then .Y = 0
                    If .X > 800 - .Width Then .X = 800 - .Width
                    If .Y > 600 - .Height Then .Y = 600 - .Height
                End If
            End With
            'Move write message
        ElseIf SelGameWindow = WriteMessageWindow Then
            With GameWindow.WriteMessage.Screen
                .X = .X + MousePosAdd.X
                .Y = .Y + MousePosAdd.Y
                If WindowsInScreen Then
                    If .X < 0 Then .X = 0
                    If .Y < 0 Then .Y = 0
                    If .X > 800 - .Width Then .X = 800 - .Width
                    If .Y > 600 - .Height Then .Y = 600 - .Height
                End If
            End With
            'Move Amount
        ElseIf SelGameWindow = AmountWindow Then
            With GameWindow.Amount.Screen
                .X = .X + MousePosAdd.X
                .Y = .Y + MousePosAdd.Y
                If WindowsInScreen Then
                    If .X < 0 Then .X = 0
                    If .Y < 0 Then .Y = 0
                    If .X > 800 - .Width Then .X = 800 - .Width
                    If .Y > 600 - .Height Then .Y = 600 - .Height
                End If
            End With
        End If

    End If

End Sub

Sub Engine_Input_Mouse_RightClick()

'******************************************
'Right click mouse
'******************************************

Dim tX As Integer
Dim tY As Integer
'Make sure engine is running

    If Not EngineRun Then Exit Sub

    '***Check for a window click***
    'Start with the last clicked window, then move in order of importance
    If Engine_Input_Mouse_RightClick_Window(LastClickedWindow) = 0 Then
        If Engine_Input_Mouse_RightClick_Window(QuickBarWindow) = 0 Then
            If Engine_Input_Mouse_RightClick_Window(InventoryWindow) = 0 Then
                If Engine_Input_Mouse_RightClick_Window(ShopWindow) = 0 Then
                    If Engine_Input_Mouse_RightClick_Window(BankWindow) = 0 Then
                        If Engine_Input_Mouse_RightClick_Window(MailboxWindow) = 0 Then
                            If Engine_Input_Mouse_RightClick_Window(WriteMessageWindow) = 0 Then
                            
                                'No windows clicked, so a tile click will take place
                                'Get the tile positions
                                Engine_ConvertCPtoTP 0, 0, MousePos.X, MousePos.Y, tX, tY
                                'Check if a sign was clicked
                                If MapData(tX, tY).Sign Then Engine_AddToChatTextBuffer Signs(MapData(tX, tY).Sign), FontColor_Info
                                'Send left click
                                sndBuf.Allocate 3
                                sndBuf.Put_Byte DataCode.User_RightClick
                                sndBuf.Put_Byte CByte(tX)
                                sndBuf.Put_Byte CByte(tY)
                                
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

End Sub

Function Engine_Input_Mouse_RightClick_Window(ByVal WindowIndex As Byte) As Byte

'******************************************
'Left click a game window
'******************************************
Dim i As Integer

    Select Case WindowIndex
    Case QuickBarWindow
        If ShowGameWindow(QuickBarWindow) Then
            With GameWindow.QuickBar
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_RightClick_Window = 1
                    LastClickedWindow = QuickBarWindow
                    'Check if an item was clicked
                    For i = 1 To 12
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                            'An item in the quickbar was clicked - get description
                            If QuickBarID(i).Type = QuickBarType_Item Then
                                Engine_SetItemDesc UserInventory(QuickBarID(i).ID).name, UserInventory(QuickBarID(i).ID).Amount
                                'A skill in the quickbar was clicked - get the name
                            ElseIf QuickBarID(i).Type = QuickBarType_Skill Then
                                Engine_SetItemDesc Engine_SkillIDtoSkillName(QuickBarID(i).ID)
                            End If
                            Exit Function
                        End If
                    Next i
                    'Item was not clicked
                    SelGameWindow = QuickBarWindow
                    Exit Function
                End If
            End With
        End If
    Case InventoryWindow
        If ShowGameWindow(InventoryWindow) Then
            With GameWindow.Inventory
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_RightClick_Window = 1
                    LastClickedWindow = InventoryWindow
                    'Check if an item was clicked
                    For i = 1 To 49
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                            If UserInventory(i).GrhIndex > 0 Then
                                Engine_SetItemDesc UserInventory(i).name, UserInventory(i).Amount
                                DragSourceWindow = InventoryWindow
                                DragItemSlot = i
                            End If
                            Exit Function
                        End If
                    Next i
                    'Item was not clicked
                    SelGameWindow = InventoryWindow
                    Exit Function
                End If
            End With
        End If
    Case ShopWindow
        If ShowGameWindow(ShopWindow) Then
            With GameWindow.Shop
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_RightClick_Window = 1
                    LastClickedWindow = ShopWindow
                    'Check if an item was clicked
                    For i = 1 To 49
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                            If i <= NPCTradeItemArraySize Then
                                If NPCTradeItems(i).GrhIndex > 0 Then
                                    Engine_SetItemDesc NPCTradeItems(i).name, 0, NPCTradeItems(i).Price
                                    DragSourceWindow = ShopWindow
                                    DragItemSlot = i
                                End If
                            End If
                            Exit Function
                        End If
                    Next i
                    'Item was not clicked
                    SelGameWindow = ShopWindow
                    Exit Function
                End If
            End With
        End If
    Case BankWindow
        If ShowGameWindow(BankWindow) Then
            With GameWindow.Bank
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_RightClick_Window = 1
                    LastClickedWindow = BankWindow
                    'Check if an item was clicked
                    For i = 1 To 49
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                            If UserBank(i).GrhIndex > 0 Then Engine_SetItemDesc UserBank(i).name, UserBank(i).Amount
                            DragSourceWindow = BankWindow
                            DragItemSlot = i
                            Exit Function
                        End If
                    Next i
                    'Item was not clicked
                    SelGameWindow = ShopWindow
                    Exit Function
                End If
            End With
        End If
    Case ViewMessageWindow
        If ShowGameWindow(ViewMessageWindow) Then
            With GameWindow.ViewMessage
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_RightClick_Window = 1
                    LastClickedWindow = ViewMessageWindow
                    'Click an item
                    For i = 1 To MaxMailObjs
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Image(i).X, .Screen.Y + .Image(i).Y, .Image(i).Width, .Image(i).Height) Then
                            Engine_SetItemDesc ReadMailData.ObjName(i), ReadMailData.ObjAmount(i)
                            Exit Function
                        End If
                    Next i
                    'Item was not clicked
                    SelGameWindow = ViewMessageWindow
                    Exit Function
                End If
            End With
        End If
    Case WriteMessageWindow
        If ShowGameWindow(WriteMessageWindow) Then
            With GameWindow.WriteMessage
                'Check if the screen was clicked
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                    Engine_Input_Mouse_RightClick_Window = 1
                    LastClickedWindow = WriteMessageWindow
                    'Click an item
                    For i = 1 To MaxMailObjs
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Image(i).X, .Screen.Y + .Image(i).Y, .Image(i).Width, .Image(i).Height) Then
                            Engine_SetItemDesc UserInventory(WriteMailData.ObjIndex(i)).name, WriteMailData.ObjAmount(i)
                            Exit Function
                        End If
                    Next i
                    'Item was not clicked
                    SelGameWindow = WriteMessageWindow
                    Exit Function
                End If
            End With
        End If
    End Select

End Function

Sub Engine_Input_Mouse_RightRelease()

'******************************************
'Right mouse button released
'******************************************
Dim i As Byte

    'Check if we released mouse and have an item in being dragged
    If DragItemSlot Then

        'Inventory -> Inventory (change slot)
        If DragSourceWindow = InventoryWindow Then
            If ShowGameWindow(InventoryWindow) Then
                With GameWindow.Inventory
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        For i = 1 To 49
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                                If DragItemSlot <> i Then
                                    'Switch slots
                                    sndBuf.Allocate 3
                                    sndBuf.Put_Byte DataCode.User_ChangeInvSlot
                                    sndBuf.Put_Byte DragItemSlot
                                    sndBuf.Put_Byte i
                                    'Clear and leave
                                    DragSourceWindow = 0
                                    DragItemSlot = 0
                                    Exit Sub
                                End If
                            End If
                        Next i
                    End If
                End With
            End If
        End If

        'Inventory -> Quick Bar
        If DragSourceWindow = InventoryWindow Then
            If ShowGameWindow(QuickBarWindow) Then
                With GameWindow.QuickBar
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        For i = 1 To 12
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                                'Drop into quick use slot
                                QuickBarID(i).Type = QuickBarType_Item
                                QuickBarID(i).ID = DragItemSlot
                                'Clear and leave
                                DragSourceWindow = 0
                                DragItemSlot = 0
                                Exit Sub
                            End If
                        Next i
                    End If
                End With
            End If
        End If
        
        'Inventory -> Depot
        If DragSourceWindow = InventoryWindow Then
            If ShowGameWindow(BankWindow) Then
                With GameWindow.Bank
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        'Single item
                        If UserInventory(DragItemSlot).Amount = 1 Then
                            sndBuf.Allocate 4
                            sndBuf.Put_Byte DataCode.User_Bank_PutItem
                            sndBuf.Put_Byte DragItemSlot
                            sndBuf.Put_Integer 1
                        'Multiple items
                        Else
                            ShowGameWindow(AmountWindow) = 1
                            LastClickedWindow = AmountWindow
                            AmountWindowValue = vbNullString
                            AmountWindowItemIndex = DragItemSlot
                            AmountWindowUsage = AW_InvToBank
                        End If
                        'Clear and leave
                        DragSourceWindow = 0
                        DragItemSlot = 0
                        Exit Sub
                    End If
                End With
            End If
        End If
        
        'Inventory -> Shop
        If DragSourceWindow = InventoryWindow Then
            If ShowGameWindow(ShopWindow) Then
                With GameWindow.Shop
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        'Single item
                        If UserInventory(DragItemSlot).Amount = 1 Then
                            sndBuf.Allocate 4
                            sndBuf.Put_Byte DataCode.User_Trade_SellToNPC
                            sndBuf.Put_Byte DragItemSlot
                            sndBuf.Put_Integer 1
                        'Multiple items
                        Else
                            ShowGameWindow(AmountWindow) = 1
                            LastClickedWindow = AmountWindow
                            AmountWindowValue = vbNullString
                            AmountWindowItemIndex = DragItemSlot
                            AmountWindowUsage = AW_InvToShop
                        End If
                        'Clear and leave
                        DragSourceWindow = 0
                        DragItemSlot = 0
                        Exit Sub
                    End If
                End With
            End If
        End If
        
        'Shop -> Inventory
        If DragSourceWindow = ShopWindow Then
            If ShowGameWindow(InventoryWindow) Then
                With GameWindow.Inventory
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        'Bring up amount window for bulk buying
                        ShowGameWindow(AmountWindow) = 1
                        LastClickedWindow = AmountWindow
                        AmountWindowValue = vbNullString
                        AmountWindowItemIndex = DragItemSlot
                        AmountWindowUsage = AW_ShopToInv
                        'Clear and leave
                        DragSourceWindow = 0
                        DragItemSlot = 0
                        Exit Sub
                    End If
                End With
            End If
        End If
        
        'Bank -> Inventory
        If DragSourceWindow = BankWindow Then
            If ShowGameWindow(InventoryWindow) Then
                With GameWindow.Inventory
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        If UserBank(DragItemSlot).Amount > 1 Then
                            'Bring up amount window for bulk withdrawing
                            ShowGameWindow(AmountWindow) = 1
                            LastClickedWindow = AmountWindow
                            AmountWindowValue = vbNullString
                            AmountWindowItemIndex = DragItemSlot
                            AmountWindowUsage = AW_BankToInv
                        Else
                            sndBuf.Allocate 4
                            sndBuf.Put_Byte DataCode.User_Bank_TakeItem
                            sndBuf.Put_Byte DragItemSlot
                            sndBuf.Put_Integer 1
                        End If
                        'Clear and leave
                        DragSourceWindow = 0
                        DragItemSlot = 0
                        Exit Sub
                    End If
                End With
            End If
        End If
                                
        'Inventory -> Mail
        If DragSourceWindow = InventoryWindow Then
            If ShowGameWindow(WriteMessageWindow) Then
                With GameWindow.WriteMessage
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        'Single item
                        If UserInventory(DragItemSlot).Amount = 1 Then
                            'Check for duplicate entries
                            For i = 1 To MaxMailObjs
                                If WriteMailData.ObjIndex(i) = DragItemSlot Then
                                    DragSourceWindow = 0
                                    DragItemSlot = 0
                                    Exit Sub
                                End If
                            Next i
                            'Place item in next free slot (if any)
                            i = 0
                            Do
                                i = i + 1
                                If i > MaxMailObjs Then
                                    DragSourceWindow = 0
                                    DragItemSlot = 0
                                    Exit Sub
                                End If
                            Loop While WriteMailData.ObjIndex(i) > 0
                            WriteMailData.ObjIndex(i) = DragItemSlot
                            WriteMailData.ObjAmount(i) = 1
                        'Multiple items
                        Else
                            ShowGameWindow(AmountWindow) = 1
                            LastClickedWindow = AmountWindow
                            AmountWindowValue = vbNullString
                            AmountWindowItemIndex = DragItemSlot
                            AmountWindowUsage = AW_InvToMail
                        End If
                        'Clear and leave
                        DragSourceWindow = 0
                        DragItemSlot = 0
                        Exit Sub
                    End If
                End With
            End If
        End If

        'Didn't release over a valid area
        DragSourceWindow = 0
        DragItemSlot = 0

    End If

End Sub

Function Engine_LegalPos(ByVal X As Integer, ByVal Y As Integer, ByVal Heading As Byte) As Boolean

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
    If MapData(X, Y).Blocked = BlockedAll Then Exit Function

    'Check the heading for directional blocking
    If Heading > 0 Then
        If MapData(X, Y).Blocked And BlockedNorth Then
            If Heading = NORTH Then Exit Function
            If Heading = NORTHEAST Then Exit Function
            If Heading = NORTHWEST Then Exit Function
        End If
        If MapData(X, Y).Blocked And BlockedEast Then
            If Heading = EAST Then Exit Function
            If Heading = NORTHEAST Then Exit Function
            If Heading = SOUTHEAST Then Exit Function
        End If
        If MapData(X, Y).Blocked And BlockedSouth Then
            If Heading = SOUTH Then Exit Function
            If Heading = SOUTHEAST Then Exit Function
            If Heading = SOUTHWEST Then Exit Function
        End If
        If MapData(X, Y).Blocked And BlockedWest Then
            If Heading = WEST Then Exit Function
            If Heading = NORTHWEST Then Exit Function
            If Heading = SOUTHWEST Then Exit Function
        End If
    End If

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
Dim Running As Byte
Dim aX As Integer
Dim aY As Integer
    
    'Check for a valid UserCharIndex
    If UserCharIndex <= 0 Or UserCharIndex > LastChar Then
    
        'We have an invalid user char index, so we must have the wrong one - request an update on the right one
        sndBuf.Put_Byte DataCode.User_RequestUserCharIndex
        Exit Sub
        
    End If

    'Dont move if the mail composing window is up
    If ShowGameWindow(WriteMessageWindow) Then Exit Sub

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

    'If the shop, mailbox or read mail window are showing, hide them
    ShowGameWindow(MailboxWindow) = 0
    ShowGameWindow(ShopWindow) = 0
    ShowGameWindow(ViewMessageWindow) = 0
    ShowGameWindow(AmountWindow) = 0
    ShowGameWindow(BankWindow) = 0
    AmountWindowUsage = 0
    AmountWindowItemIndex = 0
    AmountWindowValue = ""

    'Check for legal position
    If Engine_LegalPos(UserPos.X + aX, UserPos.Y + aY, Direction) Then

        'If running
        If GetAsyncKeyState(vbKeyShift) Then Running = 1

        'Send the information to the server
        sndBuf.Allocate 2
        sndBuf.Put_Byte DataCode.User_Move
        
        'Running or not
        If Running Then sndBuf.Put_Byte Direction Or 128 Else sndBuf.Put_Byte Direction

        'If the user changed directions or just started moving, request a position update
        If CharList(UserCharIndex).Moving = 0 Or CharList(UserCharIndex).Heading <> Direction Then
            sndBuf.Allocate 3
            sndBuf.Put_Byte DataCode.Server_SetUserPosition
            sndBuf.Put_Byte UserPos.X
            sndBuf.Put_Byte UserPos.Y
        End If

        'Move the screen and character
        Engine_Char_Move_ByHead UserCharIndex, Direction, Running
        Engine_MoveScreen Direction
        
        'Update the map sounds
        Engine_Sound_UpdateMap

        'Rotate the user to face the direction if needed
    Else

        'Only rotate if the user is not already facing that direction
        If CharList(UserCharIndex).Heading <> Direction Then
            sndBuf.Allocate 2
            sndBuf.Put_Byte DataCode.User_Rotate
            sndBuf.Put_Byte Direction
        End If

    End If

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
    
    'Set a random offset
    OBJList(ObjIndex).Offset.X = -16 + Int(Rnd * 32)
    OBJList(ObjIndex).Offset.Y = -16 + Int(Rnd * 32)

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

Private Function Engine_Collision_Between(ByVal Value As Single, ByVal Bound1 As Single, ByVal Bound2 As Single) As Byte

'*****************************************************************
'Find if a value is between two other values (used for line collision)
'*****************************************************************

    'Checks if a value lies between two bounds
    If Bound1 > Bound2 Then
        If Value >= Bound2 Then
            If Value <= Bound1 Then Engine_Collision_Between = 1
        End If
    Else
        If Value >= Bound1 Then
            If Value <= Bound2 Then Engine_Collision_Between = 1
        End If
    End If
    
End Function

Public Function Engine_Collision_Line(ByVal L1X1 As Long, ByVal L1Y1 As Long, ByVal L1X2 As Long, ByVal L1Y2 As Long, ByVal L2X1 As Long, ByVal L2Y1 As Long, ByVal L2X2 As Long, ByVal L2Y2 As Long) As Byte

'*****************************************************************
'Check if two lines intersect (return 1 if true)
'*****************************************************************

Dim m1 As Single
Dim M2 As Single
Dim B1 As Single
Dim B2 As Single
Dim IX As Single

    'This will fix problems with vertical lines
    If L1X1 = L1X2 Then L1X1 = L1X1 + 1
    If L2X1 = L2X2 Then L2X1 = L2X1 + 1

    'Find the first slope
    m1 = (L1Y2 - L1Y1) / (L1X2 - L1X1)
    B1 = L1Y2 - m1 * L1X2

    'Find the second slope
    M2 = (L2Y2 - L2Y1) / (L2X2 - L2X1)
    B2 = L2Y2 - M2 * L2X2
    
    'Check if the slopes are the same
    If M2 - m1 = 0 Then
    
        If B2 = B1 Then
            'The lines are the same
            Engine_Collision_Line = 1
        Else
            'The lines are parallel (can never intersect)
            Engine_Collision_Line = 0
        End If
        
    Else
        
        'An intersection is a point that lies on both lines. To find this, we set the Y equations equal and solve for X.
        'M1X+B1 = M2X+B2 -> M1X-M2X = -B1+B2 -> X = B1+B2/(M1-M2)
        IX = ((B2 - B1) / (m1 - M2))
        
        'Check for the collision
        If Engine_Collision_Between(IX, L1X1, L1X2) Then
            If Engine_Collision_Between(IX, L2X1, L2X2) Then Engine_Collision_Line = 1
        End If
        
    End If
    
End Function

Public Function Engine_Collision_LineRect(ByVal SX As Long, ByVal SY As Long, ByVal SW As Long, ByVal SH As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Byte

'*****************************************************************
'Check if a line intersects with a rectangle (returns 1 if true)
'*****************************************************************

    'Top line
    If Engine_Collision_Line(SX, SY, SX + SW, SY, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If
    
    'Right line
    If Engine_Collision_Line(SX + SW, SY, SX + SW, SY + SH, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

    'Bottom line
    If Engine_Collision_Line(SX, SY + SH, SX + SW, SY + SH, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

    'Left line
    If Engine_Collision_Line(SX, SY, SX, SY + SW, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

End Function

Function Engine_Collision_Rect(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal Width1 As Integer, ByVal Height1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer, ByVal Width2 As Integer, ByVal Height2 As Integer)

'*****************************************************************
'Check for collision between two rectangles
'*****************************************************************

Dim RetRect As RECT
Dim Rect1 As RECT
Dim Rect2 As RECT

    'Build the rectangles
    Rect1.Left = x1
    Rect1.Right = x1 + Width1
    Rect1.Top = Y1
    Rect1.bottom = Y1 + Height1
    Rect2.Left = x2
    Rect2.Right = x2 + Width2
    Rect2.Top = Y2
    Rect2.bottom = Y2 + Height2

    'Call collision API
    Engine_Collision_Rect = IntersectRect(RetRect, Rect1, Rect2)

End Function

Private Sub Engine_Render_Char(ByVal CharIndex As Long, ByVal PixelOffsetX As Single, ByVal PixelOffsetY As Single)

'*****************************************************************
'Draw a character to the screen by the CharIndex
'First variables are set, then all shadows drawn, then character drawn, then extras (emoticons, icons, etc)
'Any variables not handled in "Set the variables" are set in Shadow calls - do not call a second time in the
'normal character rendering calls
'*****************************************************************

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
    
    'Update blinking
    If CharList(CharIndex).BlinkTimer <= 0 Then
        CharList(CharIndex).StartBlinkTimer = CharList(CharIndex).StartBlinkTimer - ElapsedTime
        If CharList(CharIndex).StartBlinkTimer <= 0 Then
            CharList(CharIndex).BlinkTimer = 300
            CharList(CharIndex).StartBlinkTimer = Engine_GetBlinkTime
        End If
    End If
    
    'Set the map block the char is on to the TempBlock, and the block above the user as TempBlock2
    TempBlock = MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y)
    TempBlock2 = MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y - 1)

    'Check for selected NPC
    If CharIndex = TargetCharIndex Then
    
        'Clear pathway to the targeted character
        If ClearPathToTarget Then
            RenderColor(1) = D3DColorARGB(255, 100, 255, 100)
            RenderColor(2) = RenderColor(1)
            RenderColor(3) = RenderColor(1)
            RenderColor(4) = RenderColor(1)
        Else
            RenderColor(1) = D3DColorARGB(255, 255, 100, 100)
            RenderColor(2) = RenderColor(1)
            RenderColor(3) = RenderColor(1)
            RenderColor(4) = RenderColor(1)
        End If
        
    Else
        RenderColor(1) = TempBlock2.Light(1)
        RenderColor(2) = TempBlock2.Light(2)
        RenderColor(3) = TempBlock.Light(3)
        RenderColor(4) = TempBlock.Light(4)
    End If

    If CharList(CharIndex).Moving Then

        'If needed, move left and right
        If CharList(CharIndex).ScrollDirectionX <> 0 Then
            CharList(CharIndex).MoveOffset.X = CharList(CharIndex).MoveOffset.X + (ScrollPixelsPerFrameX + CharList(CharIndex).Speed + (RunningSpeed * CharList(CharIndex).Running)) * Sgn(CharList(CharIndex).ScrollDirectionX) * TickPerFrame

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
            CharList(CharIndex).MoveOffset.Y = CharList(CharIndex).MoveOffset.Y + (ScrollPixelsPerFrameY + CharList(CharIndex).Speed + (RunningSpeed * CharList(CharIndex).Running)) * Sgn(CharList(CharIndex).ScrollDirectionY) * TickPerFrame

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
                CharList(CharIndex).Moving = 0
                If CharList(CharIndex).ActionIndex = 1 Then CharList(CharIndex).ActionIndex = 0
                
                'If it is the user's character, confirm the position is correct
                If CharIndex = UserCharIndex Then
                    sndBuf.Allocate 3
                    sndBuf.Put_Byte DataCode.Server_SetUserPosition
                    sndBuf.Put_Byte CharList(CharIndex).Pos.X
                    sndBuf.Put_Byte CharList(CharIndex).Pos.Y
                End If

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
        Engine_Render_Grh CharList(CharIndex).Body.Walk(CharList(CharIndex).Heading), PixelOffsetX, PixelOffsetY, 1, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
        Engine_Render_Grh CharList(CharIndex).Weapon.Walk(CharList(CharIndex).Heading), PixelOffsetX, PixelOffsetY, 1, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1

    Else

        'Shadow
        Engine_Render_Grh CharList(CharIndex).Body.Attack(CharList(CharIndex).Heading), PixelOffsetX, PixelOffsetY, 1, 1, False, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
        Engine_Render_Grh CharList(CharIndex).Weapon.Attack(CharList(CharIndex).Heading), PixelOffsetX, PixelOffsetY, 1, 1, False, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1

        'Check if animation has stopped
        If CharList(CharIndex).Body.Attack(CharList(CharIndex).Heading).Started = 0 Then CharList(CharIndex).ActionIndex = 0

    End If

    'Draw Head
    If CharList(CharIndex).Aggressive > 0 Then
        'Aggressive
        If CharList(CharIndex).BlinkTimer > 0 Then
            CharList(CharIndex).BlinkTimer = CharList(CharIndex).BlinkTimer - ElapsedTime
            'Blinking
            Engine_Render_Grh CharList(CharIndex).Head.AgrBlink(CharList(CharIndex).HeadHeading), PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, True, False, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
        Else
            'Normal
            Engine_Render_Grh CharList(CharIndex).Head.AgrHead(CharList(CharIndex).HeadHeading), PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, True, False, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
        End If
    Else
        'Not Aggressive
        If CharList(CharIndex).BlinkTimer > 0 Then
            CharList(CharIndex).BlinkTimer = CharList(CharIndex).BlinkTimer - ElapsedTime
            'Blinking
            Engine_Render_Grh CharList(CharIndex).Head.Blink(CharList(CharIndex).HeadHeading), PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, True, False, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
        Else
            'Normal
            Engine_Render_Grh CharList(CharIndex).Head.Head(CharList(CharIndex).HeadHeading), PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, True, False, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
        End If
    End If

    'Hair
    Engine_Render_Grh CharList(CharIndex).Hair.Hair(CharList(CharIndex).HeadHeading), PixelOffsetX + CharList(CharIndex).Body.HeadOffset.X, PixelOffsetY + CharList(CharIndex).Body.HeadOffset.Y, True, False, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1

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

    'Draw name over head
    Engine_Render_Text CharList(CharIndex).name, PixelOffsetX + 16 - CharList(CharIndex).NameOffset, PixelOffsetY - 40, RenderColor(1)

    'Count the number of icons that will be needed to draw
    With CharList(CharIndex).CharStatus
        IconCount = 0
        IconCount = .Blessed + .Protected + .Strengthened + .Cursed + .WarCursed + .IronSkinned + .Exhausted
    End With
    
    'Health/Mana bars
    Engine_Render_Rectangle PixelOffsetX - 4, PixelOffsetY + 34, (CharList(CharIndex).HealthPercent / 100) * 40, 4, 1, 1, 1, 1, 1, 1, 0, 0, HealthColor, HealthColor, HealthColor, HealthColor, 0
    Engine_Render_Rectangle PixelOffsetX - 4, PixelOffsetY + 38, (CharList(CharIndex).ManaPercent / 100) * 40, 4, 1, 1, 1, 1, 1, 1, 0, 0, ManaColor, ManaColor, ManaColor, ManaColor, 0

    'Draw the icons
    If IconCount > 0 Then

        'Calculate the icon offset
        IconOffset = PixelOffsetX + 16 - (IconCount * 8)

        If CharList(CharIndex).CharStatus.Blessed Then
            Engine_Init_Grh TempGrh, 15
            Engine_Render_Grh TempGrh, IconOffset, PixelOffsetY - 50, 0, 0, False, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
            IconOffset = IconOffset + 16
        End If
        If CharList(CharIndex).CharStatus.Protected Then
            Engine_Init_Grh TempGrh, 20
            Engine_Render_Grh TempGrh, IconOffset, PixelOffsetY - 50, 0, 0, False, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
            IconOffset = IconOffset + 16
        End If
        If CharList(CharIndex).CharStatus.Strengthened Then
            Engine_Init_Grh TempGrh, 17
            Engine_Render_Grh TempGrh, IconOffset, PixelOffsetY - 50, 0, 0, False, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
            IconOffset = IconOffset + 16
        End If
        If CharList(CharIndex).CharStatus.Cursed Then
            Engine_Init_Grh TempGrh, 18
            Engine_Render_Grh TempGrh, IconOffset, PixelOffsetY - 50, 0, 0, False, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
            IconOffset = IconOffset + 16
        End If
        If CharList(CharIndex).CharStatus.WarCursed Then
            Engine_Init_Grh TempGrh, 19
            Engine_Render_Grh TempGrh, IconOffset, PixelOffsetY - 50, 0, 0, False, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
            IconOffset = IconOffset + 16
        End If
        If CharList(CharIndex).CharStatus.IronSkinned Then
            Engine_Init_Grh TempGrh, 16
            Engine_Render_Grh TempGrh, IconOffset, PixelOffsetY - 50, 0, 0, False, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
            IconOffset = IconOffset + 16
        End If
        If CharList(CharIndex).CharStatus.Exhausted Then
            Engine_Init_Grh TempGrh, 22
            Engine_Render_Grh TempGrh, IconOffset, PixelOffsetY - 50, 0, 0, False, RenderColor(1), RenderColor(2), RenderColor(3), RenderColor(4)
            IconOffset = IconOffset + 16
        End If
    End If

    'Emoticons
    If CharList(CharIndex).EmoDir > 0 Then

        'Fade in
        If CharList(CharIndex).EmoDir = 1 Then
            CharList(CharIndex).EmoFade = CharList(CharIndex).EmoFade + (ElapsedTime * 1.5)
            If CharList(CharIndex).EmoFade >= 255 Then
                CharList(CharIndex).EmoFade = 255
                CharList(CharIndex).EmoDir = 2
            End If
        End If

        'Fade out
        If CharList(CharIndex).Emoticon.Started = 0 Then    'Animation has stopped
            If CharList(CharIndex).EmoDir = 2 Then
                CharList(CharIndex).EmoFade = CharList(CharIndex).EmoFade - (ElapsedTime * 1.5)
                If CharList(CharIndex).EmoFade <= 0 Then
                    CharList(CharIndex).EmoFade = 0
                    CharList(CharIndex).EmoDir = 0
                End If
            End If
        End If

        'Render
        Engine_Render_Grh CharList(CharIndex).Emoticon, PixelOffsetX + 8, PixelOffsetY - 40, 0, 1, False, D3DColorARGB(CharList(CharIndex).EmoFade, 255, 255, 255), D3DColorARGB(CharList(CharIndex).EmoFade, 255, 255, 255), D3DColorARGB(CharList(CharIndex).EmoFade, 255, 255, 255), D3DColorARGB(CharList(CharIndex).EmoFade, 255, 255, 255)

    End If

End Sub

Private Sub Engine_Render_ChatTextBuffer()

'************************************************************
'Update and render the chat text buffer
'************************************************************

    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
    
    'Clear the LastTexture, letting the rest of the engine know that the texture needs to be changed for next rect render
    D3DDevice.SetTexture 0, Font_Default.Texture
    LastTexture = 0

    'Set up the vertex buffer
    If ShowGameWindow(ChatWindow) Then
        If ChatArrayUbound > 0 Then
            D3DDevice.SetStreamSource 0, ChatVB, FVF_Size
            D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, ChatArrayUbound \ 3
        End If
    End If

End Sub

Private Function Engine_UpdateGrh(ByRef Grh As Grh, Optional ByVal LoopAnim As Boolean = True) As Boolean

'*****************************************************************
'Updates the grh's animation
'*****************************************************************

    'Check that the grh is started
    If Grh.Started = 1 Then
    
        'Update the frame counter
        Grh.FrameCounter = Grh.FrameCounter + (TimerMultiplier * GrhData(Grh.GrhIndex).Speed)
        
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

Sub Engine_Render_GrhEX(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, Optional ByVal LoopAnim As Boolean = True, Optional ByVal Light1 As Long = -1, Optional ByVal Light2 As Long = -1, Optional ByVal Light3 As Long = -1, Optional ByVal Light4 As Long = -1, Optional ByVal Degrees As Single = 0, Optional ByVal Shadow As Byte = 0)

'*****************************************************************
'Draws a GRH transparently to a X and Y position with more options then the non-EX
'This routine is slower, but hardly slower - it is here just since there is no point
' in passing variables for things we dont use on tiles and such, which are called
' hundreds of times per loop.
'*****************************************************************
Dim CurrGrhIndex As Long    'The grh index we will be working with (acquired after updating animations)
Dim RadAngle As Single      'The angle in Radians
Dim SrcHeight As Integer
Dim SrcWidth As Integer
Dim CenterX As Single
Dim CenterY As Single
Dim Index As Long
Dim SinRad As Single
Dim CosRad As Single
Dim NewX As Single
Dim NewY As Single

    'Check to make sure it is legal
    If Grh.GrhIndex < 1 Then Exit Sub
    If GrhData(Grh.GrhIndex).NumFrames < 1 Then Exit Sub

    'Update the animation frame
    If Animate Then
        If Not Engine_UpdateGrh(Grh, LoopAnim) Then Exit Sub
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrGrhIndex = GrhData(Grh.GrhIndex).Frames(Int(Grh.FrameCounter))

    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrGrhIndex).TileWidth * TilePixelWidth * 0.5) + TilePixelWidth * 0.5
        End If
        If GrhData(CurrGrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If

    'Check for in-bounds
    If X + GrhData(CurrGrhIndex).pixelWidth > 0 Then
        If Y + GrhData(CurrGrhIndex).pixelHeight > 0 Then
            If X < frmMain.ScaleWidth Then
                If Y < frmMain.ScaleHeight Then
                
                    '***** Render the texture *****
                    'Sorry if this confuses anyone - this code was placed in-line (in opposed to calling another sub/function) to
                    ' speed things up. In-line code is always faster, especially with passing as many arguments as there is
                    ' in Engine_Render_Rectangle. This code should be just about the exact same as Engine_Render_Rectangle
                    ' minus possibly a few changes to work more specified to the manner in which it is called.
                    SrcWidth = GrhData(CurrGrhIndex).pixelWidth
                    SrcHeight = GrhData(CurrGrhIndex).pixelHeight
                    
                    'Load the surface into memory if it is not in memory and reset the timer
                    If GrhData(CurrGrhIndex).FileNum > 0 Then
                        If SurfaceTimer(GrhData(CurrGrhIndex).FileNum) = 0 Then Engine_Init_Texture GrhData(CurrGrhIndex).FileNum
                        SurfaceTimer(GrhData(CurrGrhIndex).FileNum) = SurfaceTimerMax
                    End If
                
                    'Set the texture
                    If GrhData(CurrGrhIndex).FileNum <= 0 Then
                        D3DDevice.SetTexture 0, Nothing
                    Else
                        If LastTexture <> GrhData(CurrGrhIndex).FileNum Then
                            D3DDevice.SetTexture 0, SurfaceDB(GrhData(CurrGrhIndex).FileNum)
                            LastTexture = GrhData(CurrGrhIndex).FileNum
                        End If
                    End If
                
                    'Set shadowed settings - shadows only change on the top 2 points
                    If Shadow Then
                
                        SrcWidth = SrcWidth - 1
                
                        'Set the top-left corner
                        VertexArray(0).X = X + (GrhData(CurrGrhIndex).pixelWidth * 0.5)
                        VertexArray(0).Y = Y - (GrhData(CurrGrhIndex).pixelHeight * 0.5)
                
                        'Set the top-right corner
                        VertexArray(1).X = X + GrhData(CurrGrhIndex).pixelWidth + (GrhData(CurrGrhIndex).pixelWidth * 0.5)
                        VertexArray(1).Y = Y - (GrhData(CurrGrhIndex).pixelHeight * 0.5)
                
                    Else
                
                        SrcWidth = SrcWidth + 1
                        SrcHeight = SrcHeight + 1
                
                        'Set the top-left corner
                        VertexArray(0).X = X
                        VertexArray(0).Y = Y
                
                        'Set the top-right corner
                        VertexArray(1).X = X + GrhData(CurrGrhIndex).pixelWidth
                        VertexArray(1).Y = Y
                
                    End If

                    'Set the top-left corner
                    VertexArray(0).Color = Light1
                    VertexArray(0).tu = GrhData(CurrGrhIndex).SX / SurfaceSize(GrhData(CurrGrhIndex).FileNum).X
                    VertexArray(0).tv = GrhData(CurrGrhIndex).SY / SurfaceSize(GrhData(CurrGrhIndex).FileNum).Y
                
                    'Set the top-right corner
                    VertexArray(1).Color = Light2
                    VertexArray(1).tu = (GrhData(CurrGrhIndex).SX + SrcWidth) / SurfaceSize(GrhData(CurrGrhIndex).FileNum).X
                    VertexArray(1).tv = GrhData(CurrGrhIndex).SY / SurfaceSize(GrhData(CurrGrhIndex).FileNum).Y
                
                    'Set the bottom-left corner
                    VertexArray(2).X = X
                    VertexArray(2).Y = Y + GrhData(CurrGrhIndex).pixelHeight
                    VertexArray(2).Color = Light3
                    VertexArray(2).tu = GrhData(CurrGrhIndex).SX / SurfaceSize(GrhData(CurrGrhIndex).FileNum).X
                    VertexArray(2).tv = (GrhData(CurrGrhIndex).SY + SrcHeight) / SurfaceSize(GrhData(CurrGrhIndex).FileNum).Y
                
                    'Set the bottom-right corner
                    VertexArray(3).X = X + GrhData(CurrGrhIndex).pixelWidth
                    VertexArray(3).Y = Y + GrhData(CurrGrhIndex).pixelHeight
                    VertexArray(3).Color = Light4
                    VertexArray(3).tu = (GrhData(CurrGrhIndex).SX + SrcWidth) / SurfaceSize(GrhData(CurrGrhIndex).FileNum).X
                    VertexArray(3).tv = (GrhData(CurrGrhIndex).SY + SrcHeight) / SurfaceSize(GrhData(CurrGrhIndex).FileNum).Y
                
                    'Check if a rotation is required
                    If Degrees <> 0 Then
                
                        'Converts the angle to rotate by into radians
                        RadAngle = Degrees * DegreeToRadian
                
                        'Set the CenterX and CenterY values
                        CenterX = X + (GrhData(CurrGrhIndex).pixelWidth * 0.5)
                        CenterY = Y + (GrhData(CurrGrhIndex).pixelHeight * 0.5)
                
                        'Pre-calculate the cosine and sine of the radiant
                        SinRad = Sin(RadAngle)
                        CosRad = Cos(RadAngle)
                
                        'Loops through the passed vertex buffer
                        For Index = 0 To 3
                
                            'Calculates the new X and Y co-ordinates of the vertices for the given angle around the center co-ordinates
                            NewX = ((VertexArray(Index).X - CenterX) * CosRad) - ((VertexArray(Index).Y - CenterY) * SinRad)
                            NewY = ((VertexArray(Index).X - CenterX) * SinRad) + ((VertexArray(Index).Y - CenterY) * CosRad)
                
                            'Applies the new co-ordinates to the buffer
                            VertexArray(Index).X = NewX + CenterX
                            VertexArray(Index).Y = NewY + CenterY
                
                        Next Index
                
                    End If
                
                    'Render the texture to the device
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
                
                End If
            End If
        End If
    End If

End Sub

Sub Engine_Render_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, Optional ByVal LoopAnim As Boolean = True, Optional ByVal Light1 As Long = -1, Optional ByVal Light2 As Long = -1, Optional ByVal Light3 As Long = -1, Optional ByVal Light4 As Long = -1, Optional ByVal Shadow As Byte = 0)

'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
Dim CurrGrhIndex As Long    'The grh index we will be working with (acquired after updating animations)
Dim SrcHeight As Integer
Dim SrcWidth As Integer
Dim FileNum As Integer

    'Check to make sure it is legal
    If Grh.GrhIndex < 1 Then Exit Sub
    If GrhData(Grh.GrhIndex).NumFrames < 1 Then Exit Sub

    'Update the animation frame
    If Animate Then
        If Not Engine_UpdateGrh(Grh, LoopAnim) Then Exit Sub
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrGrhIndex = GrhData(Grh.GrhIndex).Frames(Int(Grh.FrameCounter))
    
    'Set the file number in a shorter variable
    FileNum = GrhData(CurrGrhIndex).FileNum

    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrGrhIndex).TileWidth * TilePixelWidth * 0.5) + TilePixelWidth * 0.5
        End If
        If GrhData(CurrGrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If

    'Check for in-bounds
    If X + GrhData(CurrGrhIndex).pixelWidth > 0 Then
        If Y + GrhData(CurrGrhIndex).pixelHeight > 0 Then
            If X < frmMain.ScaleWidth Then
                If Y < frmMain.ScaleHeight Then
                
                    '***** Render the texture *****
                    'Sorry if this confuses anyone - this code was placed in-line (in opposed to calling another sub/function) to
                    ' speed things up. In-line code is always faster, especially with passing as many arguments as there is
                    ' in Engine_Render_Rectangle. This code should be just about the exact same as Engine_Render_Rectangle
                    ' minus possibly a few changes to work more specified to the manner in which it is called.
                    SrcWidth = GrhData(CurrGrhIndex).pixelWidth
                    SrcHeight = GrhData(CurrGrhIndex).pixelHeight
                    
                    'Load the surface into memory if it is not in memory and reset the timer
                    If FileNum > 0 Then
                        If SurfaceTimer(FileNum) = 0 Then Engine_Init_Texture FileNum
                        SurfaceTimer(FileNum) = SurfaceTimerMax
                    End If
                
                    'Set the texture
                    If FileNum <= 0 Then
                        D3DDevice.SetTexture 0, Nothing
                    Else
                        If LastTexture <> FileNum Then
                            D3DDevice.SetTexture 0, SurfaceDB(FileNum)
                            LastTexture = FileNum
                        End If
                    End If
                
                    'Set shadowed settings - shadows only change on the top 2 points
                    If Shadow Then
                
                        SrcWidth = SrcWidth - 1
                
                        'Set the top-left corner
                        VertexArray(0).X = X + (GrhData(CurrGrhIndex).pixelWidth * 0.5)
                        VertexArray(0).Y = Y - (GrhData(CurrGrhIndex).pixelHeight * 0.5)
                
                        'Set the top-right corner
                        VertexArray(1).X = X + GrhData(CurrGrhIndex).pixelWidth + (GrhData(CurrGrhIndex).pixelWidth * 0.5)
                        VertexArray(1).Y = Y - (GrhData(CurrGrhIndex).pixelHeight * 0.5)
                
                    Else

                        SrcWidth = SrcWidth + 1
                        SrcHeight = SrcHeight + 1
                
                        'Set the top-left corner
                        VertexArray(0).X = X
                        VertexArray(0).Y = Y
                
                        'Set the top-right corner
                        VertexArray(1).X = X + GrhData(CurrGrhIndex).pixelWidth
                        VertexArray(1).Y = Y
                
                    End If
                
                    'Set the top-left corner
                    VertexArray(0).Color = Light1
                    VertexArray(0).tu = GrhData(CurrGrhIndex).SX / SurfaceSize(FileNum).X
                    VertexArray(0).tv = GrhData(CurrGrhIndex).SY / SurfaceSize(FileNum).Y
                
                    'Set the top-right corner
                    VertexArray(1).Color = Light2
                    VertexArray(1).tu = (GrhData(CurrGrhIndex).SX + SrcWidth) / SurfaceSize(FileNum).X
                    VertexArray(1).tv = GrhData(CurrGrhIndex).SY / SurfaceSize(FileNum).Y
                
                    'Set the bottom-left corner
                    VertexArray(2).X = X
                    VertexArray(2).Y = Y + GrhData(CurrGrhIndex).pixelHeight
                    VertexArray(2).Color = Light3
                    VertexArray(2).tu = GrhData(CurrGrhIndex).SX / SurfaceSize(FileNum).X
                    VertexArray(2).tv = (GrhData(CurrGrhIndex).SY + SrcHeight) / SurfaceSize(FileNum).Y
                
                    'Set the bottom-right corner
                    VertexArray(3).X = X + GrhData(CurrGrhIndex).pixelWidth
                    VertexArray(3).Y = Y + GrhData(CurrGrhIndex).pixelHeight
                    VertexArray(3).Color = Light4
                    VertexArray(3).tu = (GrhData(CurrGrhIndex).SX + SrcWidth) / SurfaceSize(FileNum).X
                    VertexArray(3).tv = (GrhData(CurrGrhIndex).SY + SrcHeight) / SurfaceSize(FileNum).Y

                    'Render the texture to the device
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
                
                End If
            End If
        End If
    End If

End Sub

Private Sub Engine_Render_ChatBubble(ByVal Text As String, ByVal X As Integer, ByVal Y As Integer)

'*****************************************************************
'Renders a chat bubble and the text for the given text and co-ordinates
'*****************************************************************
Const RenderColor As Long = -1761607681
Dim TempGrh As Grh
Dim BubbleWidth As Long
Dim BubbleHeight As Long
Dim TempSplit() As String
Dim i As Long
Dim j As Long

    'Set up the temp grh
    TempGrh.FrameCounter = 1
    TempGrh.Started = 1

    'Split up the string
    TempSplit = Split(Text, vbNewLine)
    
    '*** Calculate the bubble width and height ***
    If UBound(TempSplit) > 0 Then
    
        'If there are multiple lines, it is assumed it is the max width
        BubbleWidth = BubbleMaxWidth
        
        'Because there are multiple lines, we have to calculate the height, too
        BubbleHeight = Font_Default.CharHeight * (UBound(TempSplit) + 1)
        
    Else
    
        'Theres only one line, so that line is the width
        BubbleWidth = Engine_GetTextWidth(Text)
        BubbleHeight = Font_Default.CharHeight
        
    End If
    
    'Round the width and height to the nearest 10 (the size of each chat bubble side section)
    BubbleWidth = BubbleWidth + 10
    If BubbleWidth Mod 10 Then BubbleWidth = BubbleWidth + (10 - (BubbleWidth Mod 10))
    If BubbleHeight Mod 10 Then BubbleHeight = BubbleHeight + (10 - (BubbleHeight Mod 10))
    
    'Modify the X and Y values the center the bubble
    X = X - (BubbleWidth * 0.5) + 16
    Y = Y - BubbleHeight - 20
    
    '*** Draw the bubble ***
    
    'Top-left corner
    TempGrh.GrhIndex = 109
    Engine_Render_Grh TempGrh, X, Y, 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    
    'Top-right corner
    TempGrh.GrhIndex = 111
    Engine_Render_Grh TempGrh, X + BubbleWidth + 5, Y, 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    
    'Bottom-left corner
    TempGrh.GrhIndex = 115
    Engine_Render_Grh TempGrh, X, Y + BubbleHeight + 5, 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    
    'Bottom-right corner
    TempGrh.GrhIndex = 117
    Engine_Render_Grh TempGrh, X + BubbleWidth + 5, Y + BubbleHeight + 5, 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    
    'Top side
    TempGrh.GrhIndex = 110
    For i = 0 To (BubbleWidth \ 10) - 1
        Engine_Render_Grh TempGrh, X + 5 + (i * 10), Y, 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    Next i
    
    'Left side
    TempGrh.GrhIndex = 112
    For i = 0 To (BubbleHeight \ 10) - 1
        Engine_Render_Grh TempGrh, X, Y + 5 + (i * 10), 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    Next i
    
    'Right side
    TempGrh.GrhIndex = 114
    For i = 0 To (BubbleHeight \ 10) - 1
        Engine_Render_Grh TempGrh, X + BubbleWidth + 5, Y + 5 + (i * 10), 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    Next i
    
    'Bottom side
    TempGrh.GrhIndex = 116
    For i = 0 To (BubbleWidth \ 10) - 1
        Engine_Render_Grh TempGrh, X + 5 + (i * 10), Y + BubbleHeight + 5, 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    Next i
    
    'Middle
    TempGrh.GrhIndex = 113
    For i = 1 To (BubbleWidth \ 10)
        For j = 1 To (BubbleHeight \ 10)
            Engine_Render_Grh TempGrh, X + (i * 10) - 5, Y + (j * 10) - 5, 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
        Next j
    Next i
    
    'Render the text (finally!)
    Engine_Render_Text Text, X + 5, Y + 5, D3DColorARGB(255, 0, 0, 0)

End Sub

Private Sub Engine_Render_GUI()

'*****************************************************************
'Render the GUI
'*****************************************************************

Dim TempGrh As Grh
Dim i As Byte

'Render the rest of the windows

    For i = 1 To NumGameWindows
        If i <> LastClickedWindow Then
            If ShowGameWindow(i) Then Engine_Render_GUI_Window i
        End If
    Next i

    'Render the last clicked window
    If LastClickedWindow > 0 Then
        If ShowGameWindow(LastClickedWindow) Then Engine_Render_GUI_Window LastClickedWindow
    End If

    'Render the spells list
    If DrawSkillList Then Engine_Render_Skills

    'Render an item where the cursor should be (item being dragged)
    If DragItemSlot Then
        
        Select Case DragSourceWindow
            Case InventoryWindow
                TempGrh.GrhIndex = UserInventory(DragItemSlot).GrhIndex
            Case ShopWindow
                TempGrh.GrhIndex = NPCTradeItems(DragItemSlot).GrhIndex
            Case BankWindow
                TempGrh.GrhIndex = UserBank(DragItemSlot).GrhIndex
        End Select

        'Draw
        TempGrh.FrameCounter = 1
        Engine_Render_Grh TempGrh, MousePos.X, MousePos.Y, 0, 0, False
        
    End If

    'Render the cursor
    TempGrh.FrameCounter = 1
    TempGrh.GrhIndex = 69
    Engine_Render_Grh TempGrh, MousePos.X, MousePos.Y, 0, 0, False

    'Draw item description
    Engine_Render_ItemDesc

End Sub

Private Sub Engine_Render_GUI_Window(WindowIndex As Byte)

'*****************************************************************
'Render a GUI window
'*****************************************************************
Dim TempGrh As Grh
Dim t As String
Dim s() As String
Dim i As Byte
Dim j As Long

    TempGrh.FrameCounter = 1

    Select Case WindowIndex
    
     Case StatWindow
        With GameWindow.StatWindow
            Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue, GUIColorValue
            Engine_Render_Text "Str: " & BaseStats(SID.Str) & " + " & ModStats(SID.Str) - BaseStats(SID.Str) & " (" & ModStats(SID.Str) & ")", .Screen.X + .Str.X, .Screen.Y + .Str.Y, -1
            Engine_Render_Text "Agi: " & BaseStats(SID.Agi) & " + " & ModStats(SID.Agi) - BaseStats(SID.Agi) & " (" & ModStats(SID.Agi) & ")", .Screen.X + .Agi.X, .Screen.Y + .Agi.Y, -1
            Engine_Render_Text "Mag: " & BaseStats(SID.Mag) & " + " & ModStats(SID.Mag) - BaseStats(SID.Mag) & " (" & ModStats(SID.Mag) & ")", .Screen.X + .Mag.X, .Screen.Y + .Mag.Y, -1
            If BaseStats(SID.Points) > 0 Then
                Engine_Render_Grh .AddGrh, .Screen.X + .AddStr.X, .Screen.Y + .AddStr.Y, 0, 1
                Engine_Render_Grh .AddGrh, .Screen.X + .AddAgi.X, .Screen.Y + .AddAgi.Y, 0, 1
                Engine_Render_Grh .AddGrh, .Screen.X + .AddMag.X, .Screen.Y + .AddMag.Y, 0, 1
            End If
            Engine_Render_Text "Points: " & BaseStats(SID.Points), .Screen.X + .Points.X, .Screen.Y + .Points.Y, -1
            Engine_Render_Text "Gold: " & BaseStats(SID.Gold), .Screen.X + .Gold.X, .Screen.Y + .Gold.Y, -1
            Engine_Render_Text "Def: " & BaseStats(SID.DEF) & " + " & ModStats(SID.DEF) - BaseStats(SID.DEF) & " (" & ModStats(SID.DEF) & ")", .Screen.X + .DEF.X, .Screen.Y + .DEF.Y, -1
            Engine_Render_Text "Dmg: " & BaseStats(SID.MinHIT) & "+" & ModStats(SID.MinHIT) - BaseStats(SID.MinHIT) & " ~ " & BaseStats(SID.MaxHIT) & "+" & ModStats(SID.MaxHIT) - BaseStats(SID.MaxHIT) & " (" & ModStats(SID.MinHIT) & " ~ " & ModStats(SID.MaxHIT) & ")", .Screen.X + .Dmg.X, .Screen.Y + .Dmg.Y, -1
        End With
    
     Case ChatWindow
        With GameWindow.ChatWindow
            Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue
        End With
        
        'Render the chat text
        Engine_Render_ChatTextBuffer
        
    Case MenuWindow
        With GameWindow.Menu
            Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue, GUIColorValue
        End With
            
    Case QuickBarWindow
        With GameWindow.QuickBar
            Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue, GUIColorValue
            For i = 1 To 12
                Select Case QuickBarID(i).Type
                Case QuickBarType_Skill
                    TempGrh.GrhIndex = Engine_SkillIDtoGRHID(QuickBarID(i).ID)
                    If TempGrh.GrhIndex Then Engine_Render_Grh TempGrh, .Screen.X + .Image(i).X, .Screen.Y + .Image(i).Y, 0, 0, False
                Case QuickBarType_Item
                    TempGrh.GrhIndex = UserInventory(QuickBarID(i).ID).GrhIndex
                    If TempGrh.GrhIndex Then Engine_Render_Grh TempGrh, .Screen.X + .Image(i).X, .Screen.Y + .Image(i).Y, 0, 0, False
                End Select
            Next i
        End With

    Case InventoryWindow
        With GameWindow.Inventory
            Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue, GUIColorValue
            Engine_Render_Inventory
        End With

    Case ShopWindow
        With GameWindow.Shop
            Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue, GUIColorValue
            Engine_Render_Inventory 2
        End With
    
    Case BankWindow
        With GameWindow.Bank
            Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue, GUIColorValue
            Engine_Render_Inventory 3
        End With

    Case MailboxWindow
        With GameWindow.Mailbox
            Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue, GUIColorValue
            Engine_Render_Text MailboxListBuffer, .Screen.X + .List.X, .Screen.Y + .List.Y, -1
            Engine_Render_Text "Read", .Screen.X + .ReadLbl.X, .Screen.Y + .ReadLbl.Y, -1
            Engine_Render_Text "Write", .Screen.X + .WriteLbl.X, .Screen.Y + .WriteLbl.Y, -1
            Engine_Render_Text "Delete", .Screen.X + .DeleteLbl.X, .Screen.Y + .DeleteLbl.Y, -1
            If SelMessage > 0 Then Engine_Render_Rectangle .Screen.X + .List.X, .Screen.Y + .List.Y + ((SelMessage - 1) * Font_Default.CharHeight), .List.Width, Font_Default.CharHeight, 1, 1, 1, 1, 1, 1, 0, 0, 2097217280, 2097217280, 2097217280, 2097217280    'ARGB: 125/0/255/0
        End With

    Case ViewMessageWindow
        With GameWindow.ViewMessage
            Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue, GUIColorValue
            Engine_Render_Text ReadMailData.WriterName, .Screen.X + .From.X, .Screen.Y + .From.Y, -1
            Engine_Render_Text ReadMailData.Subject, .Screen.X + .Subject.X, .Screen.Y + .Subject.Y, -1
            Engine_Render_Text ReadMailData.Message, .Screen.X + .Message.X, .Screen.Y + .Message.Y, -1
            For i = 1 To MaxMailObjs
                If ReadMailData.Obj(i) > 0 Then
                    TempGrh.GrhIndex = ReadMailData.Obj(i)
                    Engine_Render_Grh TempGrh, .Screen.X + .Image(i).X, .Screen.Y + .Image(i).Y, 0, 0, False
                End If
            Next i
        End With

    Case WriteMessageWindow
        With GameWindow.WriteMessage
            Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue, GUIColorValue
            
            '"To" text box
            If LenB(WriteMailData.RecieverName) Then Engine_Render_Text WriteMailData.RecieverName, .Screen.X + .From.X, .Screen.Y + .From.Y, -1
            If WMSelCon = wmFrom Then
                If timeGetTime Mod CursorFlashRate * 2 < CursorFlashRate Then
                    TempGrh.GrhIndex = 39
                    Engine_Render_Grh TempGrh, .Screen.X + .From.X + Engine_GetTextWidth(WriteMailData.RecieverName), .Screen.Y + .From.Y, 0, 0, False
                End If
            End If
            'Subject text box
            If LenB(WriteMailData.Subject) Then Engine_Render_Text WriteMailData.Subject, .Screen.X + .Subject.X, .Screen.Y + .Subject.Y, -1
            If WMSelCon = wmSubject Then
                If timeGetTime Mod CursorFlashRate * 2 < CursorFlashRate Then
                    TempGrh.GrhIndex = 39
                    Engine_Render_Grh TempGrh, .Screen.X + .Subject.X + Engine_GetTextWidth(WriteMailData.Subject), .Screen.Y + .Subject.Y, 0, 0, False
                End If
            End If
            'Message body text box
            t = Engine_WordWrap(WriteMailData.Message, GameWindow.WriteMessage.Message.Width)
            If LenB(WriteMailData.Message) Then Engine_Render_Text t, .Screen.X + .Message.X, .Screen.Y + .Message.Y, -1
            If WMSelCon = wmMessage Then
                If timeGetTime Mod CursorFlashRate * 2 < CursorFlashRate Then
                    If InStr(1, t, vbNewLine) Then
                        s = Split(t, vbNewLine)
                        i = UBound(s)
                        j = Engine_GetTextWidth(s(i))
                    Else
                        i = 0   'Ubound
                        j = Engine_GetTextWidth(t)  'Size
                    End If
                    TempGrh.GrhIndex = 39
                    Engine_Render_Grh TempGrh, .Screen.X + .Message.X + j, .Screen.Y + .Message.Y + (i * Font_Default.CharHeight), 0, 0, False
                End If
            End If
            'Objects
            For i = 1 To MaxMailObjs
                If WriteMailData.ObjIndex(i) > 0 Then
                    TempGrh.GrhIndex = UserInventory(WriteMailData.ObjIndex(i)).GrhIndex
                    Engine_Render_Grh TempGrh, .Screen.X + .Image(i).X, .Screen.Y + .Image(i).Y, 0, 0, False
                End If
            Next i
            
        End With

    Case AmountWindow
        With GameWindow.Amount
            Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue, GUIColorValue
            If AmountWindowValue <> "" Then Engine_Render_Text AmountWindowValue, .Screen.X + .Value.X, .Screen.Y + .Value.Y, -1
        End With

    End Select

End Sub

Public Sub Engine_Render_Inventory(Optional ByVal InventoryType As Long = 1)

'*****************************************************************
'Renders the inventory
'*****************************************************************

Dim TempGrh As Grh
Dim DestX As Single
Dim DestY As Single
Dim LoopC As Long

    Select Case InventoryType
        'User inventory
    Case 1
        For LoopC = 1 To 49
            If UserInventory(LoopC).GrhIndex Then
                DestX = GameWindow.Inventory.Screen.X + GameWindow.Inventory.Image(LoopC).X
                DestY = GameWindow.Inventory.Screen.Y + GameWindow.Inventory.Image(LoopC).Y
                TempGrh.FrameCounter = 1
                TempGrh.GrhIndex = UserInventory(LoopC).GrhIndex
                If DragItemSlot = LoopC And DragSourceWindow = InventoryWindow Then
                    Engine_Render_Grh TempGrh, DestX, DestY, 0, 0, False, -1761607681, -1761607681, -1761607681, -1761607681    'ARGB 150/255/255/255
                Else
                    Engine_Render_Grh TempGrh, DestX, DestY, 0, 0, False
                End If
                If UserInventory(LoopC).Amount <> -1 Then Engine_Render_Text UserInventory(LoopC).Amount, DestX, DestY, -1
                If UserInventory(LoopC).Equipped Then Engine_Render_Text "E", DestX + (30 - Engine_GetTextWidth("E")), DestY, -16711936
            End If
        Next LoopC
        'Shop inventory
    Case 2
        For LoopC = 1 To NPCTradeItemArraySize
            If NPCTradeItems(LoopC).GrhIndex Then
                DestX = GameWindow.Shop.Screen.X + GameWindow.Shop.Image(LoopC).X
                DestY = GameWindow.Shop.Screen.Y + GameWindow.Shop.Image(LoopC).Y
                TempGrh.FrameCounter = 1
                TempGrh.GrhIndex = NPCTradeItems(LoopC).GrhIndex
                If DragItemSlot = LoopC And DragSourceWindow = ShopWindow Then
                    Engine_Render_Grh TempGrh, DestX, DestY, 0, 0, False, -1761607681, -1761607681, -1761607681, -1761607681    'ARGB 150/255/255/255
                Else
                    Engine_Render_Grh TempGrh, DestX, DestY, 0, 0, False
                End If
            End If
        Next LoopC
        'Bank inventory
    Case 3
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            If UserBank(LoopC).GrhIndex Then
                DestX = GameWindow.Bank.Screen.X + GameWindow.Bank.Image(LoopC).X
                DestY = GameWindow.Bank.Screen.Y + GameWindow.Bank.Image(LoopC).Y
                TempGrh.FrameCounter = 1
                TempGrh.GrhIndex = UserBank(LoopC).GrhIndex
                If DragItemSlot = LoopC And DragSourceWindow = BankWindow Then
                    Engine_Render_Grh TempGrh, DestX, DestY, 0, 0, False, -1761607681, -1761607681, -1761607681, -1761607681    'ARGB 150/255/255/255
                Else
                    Engine_Render_Grh TempGrh, DestX, DestY, 0, 0, False
                End If
                If UserBank(LoopC).Amount <> -1 Then Engine_Render_Text UserBank(LoopC).Amount, DestX, DestY, -1
            End If
        Next LoopC
    End Select

End Sub

Private Sub Engine_Render_ItemDesc()

'************************************************************
'Draw description text
'************************************************************

Dim X As Integer
Dim Y As Integer
Dim i As Byte

'Check if the description text is there

    If ItemDescLines = 0 Then Exit Sub

    'Check the description position
    X = MousePos.X
    Y = MousePos.Y
    If X < 0 Then X = 0
    If X + ItemDescWidth > 800 Then X = 800 - ItemDescWidth
    If Y < 0 Then Y = 0
    If Y + (ItemDescLines * Font_Default.CharHeight) > 600 Then Y = 600 - (ItemDescLines * Font_Default.CharHeight)

    'Draw backdrop
    Engine_Render_Rectangle X - 5, Y - 5, ItemDescWidth + 10, (Font_Default.CharHeight * ItemDescLines) + 10, 1, 1, 1, 1, 1, 1, 0, 0, -1761607681, -1761607681, -1761607681, -1761607681

    'Draw text
    For i = 1 To ItemDescLines
        Engine_Render_Text ItemDescLine(i), X, Y + ((i - 1) * Font_Default.CharHeight), -16777216
    Next i

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
    If TextureNum <= 0 Then
        D3DDevice.SetTexture 0, Nothing
    Else
        If LastTexture <> TextureNum Then
            D3DDevice.SetTexture 0, SurfaceDB(TextureNum)
            LastTexture = TextureNum
        End If
    End If

    'Set the bitmap dimensions if needed
    If SrcBitmapWidth = -1 Then SrcBitmapWidth = SurfaceSize(TextureNum).X
    If SrcBitmapHeight = -1 Then SrcBitmapHeight = SurfaceSize(TextureNum).Y

    'Set shadowed settings - shadows only change on the top 2 points
    If Shadow Then

        'Set the top-left corner
        VertexArray(0).X = X + (Width * 0.5)
        VertexArray(0).Y = Y - (Height * 0.5)

        'Set the top-right corner
        VertexArray(1).X = X + Width + (Width * 0.5)
        VertexArray(1).Y = Y - (Height * 0.5)

    Else
    
        'Set the top-left corner
        VertexArray(0).X = X
        VertexArray(0).Y = Y

        'Set the top-right corner
        VertexArray(1).X = X + Width
        VertexArray(1).Y = Y

    End If
    
    'Subtract one from the width/height to get it to display correctly
    SrcWidth = SrcWidth - 1
    SrcHeight = SrcHeight - 1

    'Set the top-left corner
    VertexArray(0).Color = Color0
    VertexArray(0).tu = SrcX / SrcBitmapWidth
    VertexArray(0).tv = SrcY / SrcBitmapHeight

    'Set the top-right corner
    VertexArray(1).Color = Color1
    VertexArray(1).tu = (SrcX + SrcWidth) / SrcBitmapWidth
    VertexArray(1).tv = SrcY / SrcBitmapHeight

    'Set the bottom-left corner
    VertexArray(2).X = X
    VertexArray(2).Y = Y + Height
    VertexArray(2).Color = Color2
    VertexArray(2).tu = SrcX / SrcBitmapWidth
    VertexArray(2).tv = (SrcY + SrcHeight) / SrcBitmapHeight

    'Set the bottom-right corner
    VertexArray(3).X = X + Width
    VertexArray(3).Y = Y + Height
    VertexArray(3).Color = Color3
    VertexArray(3).tu = (SrcX + SrcWidth) / SrcBitmapWidth
    VertexArray(3).tv = (SrcY + SrcHeight) / SrcBitmapHeight

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
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size

End Sub

Sub Engine_Render_Screen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

'***********************************************
'Draw current visible to scratch area based on TileX and TileY
'***********************************************
Dim TempGrh As Grh
Dim TempRect As RECT    'Used to calculate/display our zoom level
Dim ScreenX As Integer  'Keeps track of where to place tile on screen
Dim ScreenY As Integer
Dim ChrID() As Integer
Dim ChrY() As Integer
Dim Grh As Grh          'Temp Grh for show tile and blocked
Dim x2 As Long
Dim Y2 As Long
Dim Y As Long           'Keeps track of where on map we are
Dim X As Long
Dim j As Long
Dim Angle As Single

    'Check for valid positions
    If UserPos.X = 0 Then Exit Sub
    If UserPos.Y = 0 Then Exit Sub
    
    'Clear the offset variables
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
    ParticleOffsetX = (Engine_PixelPosX(ScreenMinX) - PixelOffsetX)
    ParticleOffsetY = (Engine_PixelPosY(ScreenMinY) - PixelOffsetY)

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

    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then

        'Do a loop while device is lost
        If D3DDevice.TestCooperativeLevel = D3DERR_DEVICELOST Then Exit Sub

        'Reset the device
        D3DDevice.Reset D3DWindow

        'Reset the device and states
        If DX Is Nothing Then Set DX = New DirectX8
        If D3D Is Nothing Then Set D3D = DX.Direct3DCreate()
        If D3DX Is Nothing Then Set D3DX = New D3DX8
        
        Engine_Init_RenderStates

    Else
    
        'We have to bypass the present the first time through here or else we get an error
        If NotFirstRender = 1 Then
        
            'End the device rendering
            D3DDevice.EndScene
            
            'Get the zooming information
            If ZoomLevel > 0 Then
                TempRect.Right = 800 - (800 * ZoomLevel)
                TempRect.Left = 800 * ZoomLevel
                TempRect.bottom = 600 - (600 * ZoomLevel)
                TempRect.Top = 600 * ZoomLevel
                
                'Display the textures drawn to the device with a zoom
                D3DDevice.Present TempRect, ByVal 0, 0, ByVal 0
                
            Else
            
                TempRect.Right = 800
                TempRect.Left = 0
                TempRect.bottom = 600
                TempRect.Top = 0
        
                'Display the textures drawn to the device normally
                D3DDevice.Present TempRect, TempRect, 0, ByVal 0
            
            End If
            
        Else
        
            'Set NotFirstRender to 1 so we can start displaying
            NotFirstRender = 1
            
        End If
    
    End If

    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    D3DDevice.BeginScene

    '************** Layer 1 **************
    For Y = ScreenMinY To ScreenMaxY
        For X = ScreenMinX To ScreenMaxX
            If MapData(X, Y).Shadow(1) = 1 Then
                Engine_Render_Grh MapData(X, Y).Graphic(1), Engine_PixelPosX(ScreenX) + PixelOffsetX + ((TileBufferSize - 1) * TilePixelWidth), Engine_PixelPosY(ScreenY) + PixelOffsetY + ((TileBufferSize - 1) * TilePixelHeight), 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
                Engine_Render_Grh MapData(X, Y).Graphic(1), Engine_PixelPosX(ScreenX) + PixelOffsetX + ((TileBufferSize - 1) * TilePixelWidth), Engine_PixelPosY(ScreenY) + PixelOffsetY + ((TileBufferSize - 1) * TilePixelHeight), 0, 0, True, MapData(X, Y).Light(1), MapData(X, Y).Light(2), MapData(X, Y).Light(3), MapData(X, Y).Light(4)
            Else
                Engine_Render_Grh MapData(X, Y).Graphic(1), Engine_PixelPosX(ScreenX) + PixelOffsetX + ((TileBufferSize - 1) * TilePixelWidth), Engine_PixelPosY(ScreenY) + PixelOffsetY + ((TileBufferSize - 1) * TilePixelHeight), 0, 1, True, MapData(X, Y).Light(1), MapData(X, Y).Light(2), MapData(X, Y).Light(3), MapData(X, Y).Light(4)
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenX = ScreenX - X + ScreenMinX
        ScreenY = ScreenY + 1
    Next Y

    '************** Layer 2 **************
    ScreenY = minYOffset
    For Y = minY To maxY
        ScreenX = minXOffset
        For X = minX To maxX
            If MapData(X, Y).Graphic(2).GrhIndex Then
                If MapData(X, Y).Shadow(2) = 1 Then
                    Engine_Render_Grh MapData(X, Y).Graphic(2), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
                    Engine_Render_Grh MapData(X, Y).Graphic(2), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0, True, MapData(X, Y).Light(5), MapData(X, Y).Light(6), MapData(X, Y).Light(7), MapData(X, Y).Light(8)
                Else
                    Engine_Render_Grh MapData(X, Y).Graphic(2), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, MapData(X, Y).Light(5), MapData(X, Y).Light(6), MapData(X, Y).Light(7), MapData(X, Y).Light(8)
                End If
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    '************** Layer 3 **************
    ScreenY = minYOffset
    For Y = minY To maxY
        ScreenX = minXOffset
        For X = minX To maxX
            If MapData(X, Y).Graphic(3).GrhIndex Then
                If MapData(X, Y).Shadow(3) = 1 Then
                    Engine_Render_Grh MapData(X, Y).Graphic(3), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
                    Engine_Render_Grh MapData(X, Y).Graphic(3), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0, True, MapData(X, Y).Light(9), MapData(X, Y).Light(10), MapData(X, Y).Light(11), MapData(X, Y).Light(12)
                Else
                    Engine_Render_Grh MapData(X, Y).Graphic(3), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, MapData(X, Y).Light(9), MapData(X, Y).Light(10), MapData(X, Y).Light(11), MapData(X, Y).Light(12)
                End If
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y

    '************** Objects **************
    For j = 1 To LastObj
        If OBJList(j).Grh.GrhIndex Then
            X = Engine_PixelPosX(minXOffset + (OBJList(j).Pos.X - minX)) + PixelOffsetX + OBJList(j).Offset.X
            Y = Engine_PixelPosY(minYOffset + (OBJList(j).Pos.Y - minY)) + PixelOffsetY + OBJList(j).Offset.Y
            If Y >= -32 Then
                If Y <= 632 Then
                    If X >= -32 Then
                        If X <= 832 Then
                            x2 = minXOffset + (OBJList(j).Pos.X - minX)
                            Y2 = minYOffset + (OBJList(j).Pos.Y - minY)
                            Engine_Render_Grh OBJList(j).Grh, X, Y, 1, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
                            Engine_Render_Grh OBJList(j).Grh, X, Y, 1, 1, True, MapData(OBJList(j).Pos.X, OBJList(j).Pos.Y).Light(1), _
                                MapData(OBJList(j).Pos.X, OBJList(j).Pos.Y).Light(2), MapData(OBJList(j).Pos.X, OBJList(j).Pos.Y).Light(3), _
                                MapData(OBJList(j).Pos.X, OBJList(j).Pos.Y).Light(4)
                        End If
                    End If
                End If
            End If
        End If
    Next j

    '************** Characters **************
    'Size the to the smallest safe size (LastChar)
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
            If Y >= -32 And Y <= 632 And X >= -32 And X <= 832 Then
                        
                'Update the NPC chat
                Engine_NPCChat_Update ChrID(j)
            
                'Draw the character
                Engine_Render_Char ChrID(j), X, Y
                
            Else
                
                'Update just the real position
                CharList(ChrID(j)).RealPos.X = X + CharList(ChrID(j)).MoveOffset.X
                CharList(ChrID(j)).RealPos.Y = Y + CharList(ChrID(j)).MoveOffset.Y
            
            End If
        End If
    Next j

    '************** Layer 4 **************
    ScreenY = minYOffset
    For Y = minY To maxY
        ScreenX = minXOffset
        For X = minX To maxX
            If MapData(X, Y).Graphic(4).GrhIndex Then
                If MapData(X, Y).Shadow(4) = 1 Then
                    Engine_Render_Grh MapData(X, Y).Graphic(4), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
                    Engine_Render_Grh MapData(X, Y).Graphic(4), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0, True, MapData(X, Y).Light(13), MapData(X, Y).Light(14), MapData(X, Y).Light(15), MapData(X, Y).Light(16)
                Else
                    Engine_Render_Grh MapData(X, Y).Graphic(4), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, MapData(X, Y).Light(13), MapData(X, Y).Light(14), MapData(X, Y).Light(15), MapData(X, Y).Light(16)
                End If
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y

    '************** Layer 5 **************
    ScreenY = minYOffset
    For Y = minY To maxY
        ScreenX = minXOffset
        For X = minX To maxX
            If MapData(X, Y).Graphic(5).GrhIndex Then
                If MapData(X, Y).Shadow(5) = 1 Then
                    Engine_Render_Grh MapData(X, Y).Graphic(5), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
                    Engine_Render_Grh MapData(X, Y).Graphic(5), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0, True, MapData(X, Y).Light(17), MapData(X, Y).Light(18), MapData(X, Y).Light(19), MapData(X, Y).Light(20)
                Else
                    Engine_Render_Grh MapData(X, Y).Graphic(5), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, MapData(X, Y).Light(17), MapData(X, Y).Light(18), MapData(X, Y).Light(19), MapData(X, Y).Light(20)
                End If
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    '************** Layer 6 **************
    ScreenY = minYOffset
    For Y = minY To maxY
        ScreenX = minXOffset
        For X = minX To maxX
            If MapData(X, Y).Graphic(6).GrhIndex Then
                If MapData(X, Y).Shadow(6) = 1 Then
                    Engine_Render_Grh MapData(X, Y).Graphic(6), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
                    Engine_Render_Grh MapData(X, Y).Graphic(6), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 0, True, MapData(X, Y).Light(21), MapData(X, Y).Light(22), MapData(X, Y).Light(23), MapData(X, Y).Light(24)
                Else
                    Engine_Render_Grh MapData(X, Y).Graphic(6), Engine_PixelPosX(ScreenX) + PixelOffsetX, Engine_PixelPosY(ScreenY) + PixelOffsetY, 0, 1, True, MapData(X, Y).Light(21), MapData(X, Y).Light(22), MapData(X, Y).Light(23), MapData(X, Y).Light(24)
                End If
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y

    '************** Effects **************
    'Loop to do drawing
    If LastEffect > 0 Then
        For j = 1 To LastEffect
            If EffectList(j).Grh.GrhIndex Then
                X = Engine_PixelPosX(minXOffset + (EffectList(j).Pos.X - minX)) + PixelOffsetX
                Y = Engine_PixelPosY(minYOffset + (EffectList(j).Pos.Y - minY)) + PixelOffsetY
                If EffectList(j).Time <> 0 And EffectList(j).Time < timeGetTime Then
                
                    'Timer ran out
                    Engine_Effect_Erase j
                    
                ElseIf Y >= -32 And Y <= 632 And X >= -32 And X <= 832 Then
                
                    'Timer or animation going, render
                    Engine_Render_GrhEX EffectList(j).Grh, X, Y, 1, 1, 0, , , , , EffectList(j).Angle
                    
                    'Check if the animation is still running
                    If EffectList(j).Animated = 1 Then
                        If EffectList(j).Grh.Started = 0 Then Engine_Effect_Erase j
                    End If
                    
                Else
                
                    'Animation is going but not in screen, update the animation frame
                    Engine_UpdateGrh EffectList(j).Grh, False
                    
                    'Check if the animation is still running
                    If EffectList(j).Animated = 1 Then
                        If EffectList(j).Grh.Started = 0 Then Engine_Effect_Erase j
                    End If
                    
                End If
            End If
        Next j

    End If
    
    '************** Projectiles **************
    'Loop to do drawing
    If LastProjectile > 0 Then
        For j = 1 To LastProjectile
            If ProjectileList(j).Grh.GrhIndex Then
                
                'Update the position
                Angle = DegreeToRadian * Engine_GetAngle(ProjectileList(j).X, ProjectileList(j).Y, ProjectileList(j).tX, ProjectileList(j).tY)
                ProjectileList(j).X = ProjectileList(j).X + Sin(Angle) * 10
                ProjectileList(j).Y = ProjectileList(j).Y - Cos(Angle) * 10
                
                'Update the rotation
                If ProjectileList(j).RotateSpeed > 0 Then
                    ProjectileList(j).Rotate = ProjectileList(j).Rotate + (ProjectileList(j).RotateSpeed * ElapsedTime * 0.01)
                    Do While ProjectileList(j).Rotate > 360
                        ProjectileList(j).Rotate = ProjectileList(j).Rotate - 360
                    Loop
                End If

                'Draw if within range
                X = ((minXOffset - minX - 1) * 32) + ProjectileList(j).X + PixelOffsetX
                Y = ((minYOffset - minY - 1) * 32) + ProjectileList(j).Y + PixelOffsetY
                If Y >= -32 Then
                    If Y <= 632 Then
                        If X >= -32 Then
                            If X <= 832 Then
                                If ProjectileList(j).Rotate = 0 Then
                                    Engine_Render_Grh ProjectileList(j).Grh, X, Y, 0, 1, 1, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
                                    Engine_Render_Grh ProjectileList(j).Grh, X, Y, 0, 0, 1
                                Else
                                    Engine_Render_GrhEX ProjectileList(j).Grh, X, Y, 0, 0, 1, ShadowColor, ShadowColor, ShadowColor, ShadowColor, ProjectileList(j).Rotate, 1
                                    Engine_Render_GrhEX ProjectileList(j).Grh, X, Y, 0, 1, 1, , , , , ProjectileList(j).Rotate
                                End If
                            End If
                        End If
                    End If
                End If
                
            End If
        Next j
        
        'Check if it is close enough to the target to remove
        For j = 1 To LastProjectile
            If ProjectileList(j).Grh.GrhIndex Then
                If Abs(ProjectileList(j).X - ProjectileList(j).tX) < 20 Then
                    If Abs(ProjectileList(j).Y - ProjectileList(j).tY) < 20 Then
                        Engine_Projectile_Erase j
                    End If
                End If
            End If
        Next j
        
    End If
    
    '************** Blood Splatters **************
    'Loop to do drawing
    For j = 1 To LastBlood
        If BloodList(j).Grh.GrhIndex Then
            X = Engine_PixelPosX(minXOffset + (BloodList(j).Pos.X - minX)) + PixelOffsetX
            Y = Engine_PixelPosY(minYOffset + (BloodList(j).Pos.Y - minY)) + PixelOffsetY
            If Y >= -32 Then
                If Y <= 632 Then
                    If X >= -32 Then
                        If X <= 832 Then
                            Engine_Render_Grh BloodList(j).Grh, X, Y, 1, 1, False
                        End If
                    End If
                End If
            End If
        End If
    Next j

    'Seperate loop to remove the unused - I dont like removing while drawing
    For j = 1 To LastBlood
        If BloodList(j).Grh.GrhIndex Then
            If BloodList(j).Grh.Started = 0 Then Engine_Blood_Erase j
        End If
    Next j

    '************** Update weather **************
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

    '************** Chat bubbles **************
    'Loop through the chars
    For j = 1 To LastChar
        If CharList(j).Active Then
            If LenB(CharList(j).BubbleStr) Then
                If CharList(j).RealPos.X > -25 Then
                    If CharList(j).RealPos.X < 825 Then
                        If CharList(j).RealPos.Y > -25 Then
                            If CharList(j).RealPos.Y < 625 Then
                                Engine_Render_ChatBubble CharList(j).BubbleStr, CharList(j).RealPos.X, CharList(j).RealPos.Y
                                CharList(j).BubbleTime = CharList(j).BubbleTime - ElapsedTime
                                If CharList(j).BubbleTime <= 0 Then
                                    CharList(j).BubbleTime = 0
                                    CharList(j).BubbleStr = ""
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next j

    '************** Damage text **************
    'Loop to do drawing
    For j = 1 To LastDamage
        If DamageList(j).Counter > 0 Then
            DamageList(j).Counter = DamageList(j).Counter - ElapsedTime
            X = ((minXOffset + (DamageList(j).Pos.X - minX) - 1) * TilePixelWidth) + PixelOffsetX
            Y = ((minYOffset + (DamageList(j).Pos.Y - minY) - 1) * TilePixelHeight) + PixelOffsetY
            If Y >= -32 Then
                If Y <= 632 Then
                    If X >= -32 Then
                        If X <= 832 Then
                            Engine_Render_Text DamageList(j).Value, X, Y, D3DColorARGB(255, 255, 0, 0)
                        End If
                    End If
                End If
            End If
            DamageList(j).Pos.Y = DamageList(j).Pos.Y - (ElapsedTime * 0.001)
        End If
    Next j

    'Seperate loop to remove the unused - I dont like removing while drawing
    For j = 1 To LastDamage
        If DamageList(j).Width Then
            If DamageList(j).Counter <= 0 Then Engine_Damage_Erase j
        End If
    Next j

    '************** Misc Rendering **************

    'Update and render particle effects
    Effect_UpdateAll

    'Clear the shift-related variables
    LastOffsetX = ParticleOffsetX
    LastOffsetY = ParticleOffsetY

    'Render the GUI
    Engine_Render_GUI
    
    'Draw entered text
    If EnterText = True Then
        If EnterTextBufferWidth = 0 Then EnterTextBufferWidth = 1   'Dividing by 0 is never good
        If LenB(ShownText) Then Engine_Render_Text ShownText, GameWindow.ChatWindow.Screen.X + GameWindow.ChatWindow.Text.X, GameWindow.ChatWindow.Screen.Y + GameWindow.ChatWindow.Text.Y, -1
        If timeGetTime Mod CursorFlashRate * 2 < CursorFlashRate Then
            TempGrh.GrhIndex = 39
            TempGrh.FrameCounter = 1
            TempGrh.Started = 1
            TempGrh.SpeedCounter = 0
            Engine_Render_Grh TempGrh, GameWindow.ChatWindow.Screen.X + GameWindow.ChatWindow.Text.X + Engine_GetTextWidth(ShownText), GameWindow.ChatWindow.Screen.Y + GameWindow.ChatWindow.Text.Y, 0, 0, False
        End If
    End If
    
    '************** Mini-map **************
    Const tS As Single = 5  'Size of the mini-map dots
    
    'Check if the mini-map is being shown
    If ShowMiniMap Then
    
        'Make sure the mini-map vertex buffer is valid
        If MiniMapVBSize > 0 Then
            
            'Clear the texture
            LastTexture = 0
            D3DDevice.SetTexture 0, Nothing
            
            'Draw the map outline
            D3DDevice.SetStreamSource 0, MiniMapVB, FVF_Size
            D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, MiniMapVBSize \ 3
  
            'Draw the characters
            For X = 1 To LastChar
                If CharList(X).Pos.X > ScreenMinX Then
                    If CharList(X).Pos.X < ScreenMaxX Then
                        If CharList(X).Pos.Y > ScreenMinY Then
                            If CharList(X).Pos.Y < ScreenMaxY Then
                                If X = UserCharIndex Then j = D3DColorARGB(200, 0, 255, 0) Else j = D3DColorARGB(200, 0, 255, 255)
                                Engine_Render_Rectangle CharList(X).Pos.X * tS, CharList(X).Pos.Y * tS, tS, tS, 1, 1, 1, 1, 1, 1, 0, 0, j, j, j, j
                            End If
                        End If
                    End If
                End If
            Next X
            
        End If
        
    End If

    'Show FPS & Lag
    Engine_Render_Text "FPS: " & FPS, 720, 2, -1
    Engine_Render_Text "PTD: " & PTD & " ms", 720, 15, -1
    
End Sub

Public Sub Engine_BuildMiniMap()

'***************************************************
'Builds the array for the minimap. Theres multiple styles available, but only one
'is used in the demo, so experiment with them and check which one you like!
'***************************************************
Dim NumMiniMapTiles As Integer      'UBound of the MiniMapTile array
Dim MiniMapTile() As MiniMapTile    'Color of each tile and their position
Dim MMC_Blocked As Long
Dim MMC_Exit As Long
Dim MMC_Sign As Long
Dim Offset As Long
Dim X As Long
Dim Y As Long
Dim j As Long

    'Change to the type of map you want
    Const UseOption As Byte = 2
    
    'The size of the tiles
    Const MiniMapSize As Single = 5

    'Create the colors (character colors are defined in Engine_RenderScreen when it is rendered)
    MMC_Blocked = D3DColorARGB(75, 255, 255, 255)   'Blocked tiles
    MMC_Exit = D3DColorARGB(150, 255, 0, 0)         'Exit tiles (warps)
    MMC_Sign = D3DColorARGB(125, 255, 255, 0)       'Tiles with a sign
    
    'Clear the old array by resizing to the largest array we can possibly use
    ReDim MiniMapTile(1 To CInt(XMaxMapSize) * CInt(YMaxMapSize)) As MiniMapTile
    NumMiniMapTiles = 0
    
    Select Case UseOption
        
        '***** Option 1 *****
        Case 1

            For Y = YMinMapSize To YMaxMapSize
                For X = XMinMapSize To XMaxMapSize
                    
                    'Check for signs
                    If MapData(X, Y).Sign > 1 Then
                        NumMiniMapTiles = NumMiniMapTiles + 1
                        MiniMapTile(NumMiniMapTiles).X = X
                        MiniMapTile(NumMiniMapTiles).Y = Y
                        MiniMapTile(NumMiniMapTiles).Color = MMC_Sign
                    Else
                    
                        'Check for exits
                        If MapData(X, Y).Warp = 1 Then
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

            For Y = YMinMapSize To YMaxMapSize
                j = 0   'Clear the row settings
                For X = XMinMapSize To XMaxMapSize
                    
                    'Check if there is a sign
                    If MapData(X, Y).Sign > 1 Then
                        NumMiniMapTiles = NumMiniMapTiles + 1
                        MiniMapTile(NumMiniMapTiles).X = X
                        MiniMapTile(NumMiniMapTiles).Y = Y
                        MiniMapTile(NumMiniMapTiles).Color = MMC_Sign
                    Else
                    
                        'Check if there is an exit
                        If MapData(X, Y).Warp = 1 Then
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
                                        If X + 1 <= XMaxMapSize Then
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
                                        If Y > YMinMapSize Then
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
                                        If Y < YMaxMapSize Then
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
                                        If Y > YMinMapSize Then
                                            If Y < YMaxMapSize Then
                                                If X > XMinMapSize Then
                                                    If X < XMaxMapSize Then
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
                                    If X < XMaxMapSize Then
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
        Exit Sub
    Else
        ReDim Preserve MiniMapTile(1 To NumMiniMapTiles)
    End If
    
    '***** Build the vertex buffer according to the information we gathered in the MiniMapTile array *****
    
    'Create the temp vertex array large enough to fit every tile (2 triangles per tile, 3 points per triangle)
    ReDim tVA(0 To (NumMiniMapTiles * 6) - 1) As TLVERTEX
    
    'Start our offset at -6 so the first offset is 0
    Offset = -6
    
    'Fill the temp vertex array
    For j = 1 To NumMiniMapTiles
    
        'Raise the offset count
        Offset = Offset + 6
    
        '*** Triangle 1 ***
        
        'Top-left corner
        With tVA(0 + Offset)
            .X = MiniMapTile(j).X * MiniMapSize
            .Y = MiniMapTile(j).Y * MiniMapSize
            .Color = MiniMapTile(j).Color
            .Rhw = 1
        End With
        
        'Top-right corner
        With tVA(1 + Offset)
            .X = MiniMapTile(j).X * MiniMapSize + MiniMapSize
            .Y = MiniMapTile(j).Y * MiniMapSize
            .Color = MiniMapTile(j).Color
            .Rhw = 1
        End With
        
        'Bottom-left corner
        With tVA(2 + Offset)
            .X = MiniMapTile(j).X * MiniMapSize
            .Y = MiniMapTile(j).Y * MiniMapSize + MiniMapSize
            .Color = MiniMapTile(j).Color
            .Rhw = 1
        End With
        
        '*** Triangle 2 ***
        
        'Top-right corner
        tVA(3 + Offset) = tVA(1 + Offset)
        
        'Bottom-right corner
        With tVA(4 + Offset)
            .X = MiniMapTile(j).X * MiniMapSize + MiniMapSize
            .Y = MiniMapTile(j).Y * MiniMapSize + MiniMapSize
            .Color = MiniMapTile(j).Color
            .Rhw = 1
        End With
        
        'Bottom-left corner
        tVA(5 + Offset) = tVA(2 + Offset)
        
    Next j
    
    'Build the vertex buffer
    MiniMapVBSize = Offset + 6
    Set MiniMapVB = D3DDevice.CreateVertexBuffer(FVF_Size * MiniMapVBSize, 0, FVF, D3DPOOL_MANAGED)
    D3DVertexBuffer8SetData MiniMapVB, 0, FVF_Size * MiniMapVBSize, 0, tVA(0)
    
    'Clear the temp arrays
    Erase tVA
    Erase MiniMapTile

End Sub

Private Function Engine_NPCChat_MeetsConditions(ByVal NPCIndex As Integer, ByVal LineIndex As Byte, Optional ByVal SayLine As String = "") As Byte

'***************************************************
'Checks if the conditions have been satisfied for a chat line
'***************************************************
Dim s() As String
Dim j As Byte
Dim i As Byte

    'Make sure we have a valid line and index
    If LineIndex = 0 Then Exit Function
    If CharList(NPCIndex).NPCChatIndex = 0 Then Exit Function
    If CharList(NPCIndex).NPCChatIndex > UBound(NPCChat()) Then Exit Function
    If LineIndex > UBound(NPCChat(CharList(NPCIndex).NPCChatIndex).ChatLine()) Then Exit Function

    'Woo baby, we're not going to want to type THIS line more then once!
    With NPCChat(CharList(NPCIndex).NPCChatIndex).ChatLine(LineIndex)
        
        'If the SayLine is used, it must be the user just talked - so we ONLY want a trigger line!
        If LenB(SayLine) Then   'If the string is not empty
            SayLine = UCase$(SayLine)   'We compair it in UCase$(), since case doesn't matter
            If .NumConditions = 0 Then Exit Function        'If there are no conditions, then theres definintely no SAY condition
            For i = 1 To .NumConditions
                If .Conditions(i).Condition = NPCCHAT_COND_SAY Then Exit For    'Good, we have a SAY condition! We can continue...
                If i = .NumConditions Then Exit Function    'Last condition checked, and it wasn't a SAY, so no SAYs found - goodbye :(
            Next i
        End If
        
        'Loop through all the conditions
        For i = 1 To .NumConditions
        
            'Check what condition it is - keep in mind we exit on a "False" situation, so are checks
            ' are written to check if the condition is false, not true (a little more confusing, but effecient)
            Select Case .Conditions(i).Condition
                
                'If there is a SAY requirement, things get tricky...
                Case NPCCHAT_COND_SAY
                    If SayLine = "" Then Exit Function  'No chance it can be right if theres no text!
                    s() = Split(.Conditions(i).ValueStr, ",")   'Split up our commas (which allow us to have multiple valid words)
                    For j = 0 To UBound(s)  'Loop through each word so we can check if it is in the SayLine
                        If InStr(1, SayLine, s(j)) Then 'Check if the trigger word is in the SayLine
                            Exit For    'Match made! We're good to go - get the hell outta here!
                        End If
                        If j = UBound(s) Then Exit Function 'Oh bummer, the last trigger word was checked and was a no-go, we loose!
                    Next j
                    
                'User doesn't know skill X
                Case NPCCHAT_COND_DONTKNOWSKILL
                    If Not (UserKnowSkill(.Conditions(i).Value) = 0) Then Exit Function
                    
                'User knows skill X
                Case NPCCHAT_COND_KNOWSKILL
                    If Not (UserKnowSkill(.Conditions(i).Value) = 1) Then Exit Function
                
                'NPC's HP is less then or equal to X percent
                Case NPCCHAT_COND_HPLESSTHAN
                    If Not (CharList(UserCharIndex).HealthPercent <= .Conditions(i).Value) Then Exit Function
                    
                'NPC's HP is greater then or equal to X percent
                Case NPCCHAT_COND_HPMORETHAN
                    If Not (CharList(UserCharIndex).HealthPercent >= .Conditions(i).Value) Then Exit Function

                'User's level is less than or equal to X
                Case NPCCHAT_COND_LEVELLESSTHAN
                    If Not (BaseStats(SID.ELV) <= .Conditions(i).Value) Then Exit Function
                    
                'User level is greater than or equal to X
                Case NPCCHAT_COND_LEVELMORETHAN
                    If Not (BaseStats(SID.ELV) >= .Conditions(i).Value) Then Exit Function
            
            End Select
            
        Next i
        
    End With
    
    'We made it, horray!
    Engine_NPCChat_MeetsConditions = 1
    
End Function

Public Sub Engine_NPCChat_CheckForChatTriggers(ByVal ChatTxt As String)

'***************************************************
'Checks for a NPC chat triggers
'***************************************************
Dim i As Integer
Dim j As Byte

    For i = 1 To LastChar
        
        'We're going to be using this object a hell of a lot...
        With CharList(i)
            
            'We only want an active char
            If .Active Then
            
                'Make sure the NPC has automated chat
                If .NPCChatIndex > 0 Then
    
                    'Check for a valid distance
                    If Engine_RectDistance(.RealPos.X, .RealPos.Y, .RealPos.X - 350, .RealPos.Y - 250, 351, 251) Then
                    
                        'Get the next line to use
                        j = Engine_NPCChat_NextLine(i, ChatTxt)
                        
                        'If j = 0, then no valid lines were found
                        If j > 0 Then
                        
                            'Assign the new line
                            .NPCChatLine = j
                            
                            'Say the chat (delay assigned through the routine)
                            Engine_NPCChat_AddText i
                            
                        End If
                    
                    End If
                    
                End If
                    
            End If
            
        End With
    
    Next i
                    

End Sub

Private Sub Engine_NPCChat_Update(ByVal CharIndex As Integer)

'***************************************************
'Updates the automated NPC chatting
'***************************************************
Dim i As Byte

    'We're going to be using this object a hell of a lot...
    With CharList(CharIndex)
        
        'Make sure the NPC has automated chat
        If .NPCChatIndex > 0 Then
            
            'Check for a valid distance
            If Engine_RectDistance(.RealPos.X, .RealPos.Y, .RealPos.X - 350, .RealPos.Y - 250, 351, 251) Then
            
                'Update the delay time
                If .NPCChatDelay > 0 Then
                    .NPCChatDelay = .NPCChatDelay - ElapsedTime
                    
                'Time to get a new line!
                Else
                    
                    'Get the new NPCChat line
                    i = Engine_NPCChat_NextLine(CharIndex)
                    If i = 0 Then Exit Sub
                    .NPCChatLine = i
                    
                    'Add the chat
                    Engine_NPCChat_AddText CharIndex

                End If
            End If
        End If
        
    End With

End Sub

Private Sub Engine_NPCChat_AddText(ByVal CharIndex As Integer)

'***************************************************
'Adds the NPCChat text according to the style
'***************************************************
    
    With CharList(CharIndex)

        'Check for text before adding it
        If LenB(NPCChat(.NPCChatIndex).ChatLine(.NPCChatLine).Text) Then
    
            'Find out the style used, and add the chat according to the style
            Select Case NPCChat(.NPCChatIndex).ChatLine(.NPCChatLine).Style
                Case NPCCHAT_STYLE_BUBBLE
                    Engine_MakeChatBubble CharIndex, Engine_WordWrap(.name & ": " & NPCChat(.NPCChatIndex).ChatLine(.NPCChatLine).Text, BubbleMaxWidth)
                Case NPCCHAT_STYLE_BOX
                    Engine_AddToChatTextBuffer .name & ": " & NPCChat(.NPCChatIndex).ChatLine(.NPCChatLine).Text, FontColor_Talk
                Case NPCCHAT_STYLE_BOTH
                    Engine_MakeChatBubble CharIndex, Engine_WordWrap(.name & ": " & NPCChat(.NPCChatIndex).ChatLine(.NPCChatLine).Text, BubbleMaxWidth)
                    Engine_AddToChatTextBuffer .name & ": " & NPCChat(.NPCChatIndex).ChatLine(.NPCChatLine).Text, FontColor_Talk
            End Select
            
        End If
            
        'Add the chat delay (we do the delay even if theres no text)
        .NPCChatDelay = NPCChat(.NPCChatIndex).ChatLine(.NPCChatLine).Delay
        
    End With

End Sub

Private Function Engine_NPCChat_NextLine(ByVal CharIndex As Integer, Optional ByVal ChatTxt As String = "")

'***************************************************
'Gets the next free line to use for the NPC chat (0 if none found)
'***************************************************
Dim b() As Byte
Dim k As Byte
Dim j As Byte
Dim i As Byte

    With CharList(CharIndex)
    
        'Select the new line to start from according to the format
        Select Case NPCChat(.NPCChatIndex).Format
        
            'Linear selection
            Case NPCCHAT_FORMAT_LINEAR
            
                'Start with the next line
                i = .NPCChatLine + 1
                If i > NPCChat(.NPCChatIndex).NumLines Then i = 1
                
                'Loop through all the lines, checking for the next line with a valid condition
                For j = 1 To NPCChat(.NPCChatIndex).NumLines
                    
                    'Get the new line number to check - roll over to the start if needed
                    k = i + j
                    If k > NPCChat(.NPCChatIndex).NumLines Then k = k - NPCChat(.NPCChatIndex).NumLines
                    
                    'Check if the conditions were met
                    If Engine_NPCChat_MeetsConditions(CharIndex, k, ChatTxt) = 1 Then Exit For
                    
                    'If j is on the last index, then no conditions were met - put on a delay and leave
                    If j = NPCChat(.NPCChatIndex).NumLines Then
                        .NPCChatDelay = 1500    'This delay lets a load off the client
                        Exit Function
                    End If
                    
                Next j
                
            'Random selection
            Case NPCCHAT_FORMAT_RANDOM
            
                'Scramble the numbers so we can pick randomly
                ReDim b(1 To NPCChat(.NPCChatIndex).NumLines)       'Room for all the lines
                For i = 1 To NPCChat(.NPCChatIndex).NumLines        'Loop through every line
                    Do  'Keep looping until we get what we want
                        j = Int(Rnd * NPCChat(.NPCChatIndex).NumLines) + 1  'We have to hold the value in a temp variable
                        If b(j) = 0 Then    'If = 0, the index is free
                            b(j) = i        'Store the index in the random array slot
                            Exit Do         'Leave the DO loop since we have what we want
                        End If
                    Loop
                Next i

                'Now b() holds all the line numbers scrambled up, so we can go through one by one just like with linear
                For j = 1 To NPCChat(.NPCChatIndex).NumLines - 1    '-1 because we are took out the index we already used
                    
                    'Make sure the number is valid (just in case)
                    If b(j) <> 0 Then
                        
                        'Don't check the line we just used (yet)
                        If .NPCChatLine <> b(j) Then
                            
                            'Check the conditions
                            If Engine_NPCChat_MeetsConditions(CharIndex, b(j), ChatTxt) = 1 Then
                                k = b(j)    'Store the successful value in the k variable for below
                                Exit For
                            End If
                        
                        End If
                        
                    End If
                        
                    'If j is on the last index, and no conditions were met, we try the line we last used
                    If j = NPCChat(.NPCChatIndex).NumLines - 1 Then 'If the For loop is just about to end
                        If b(j) > 0 Then    'If this is the NPC's first line, it'd be 0, so check to make sure its not 0 just in case
                            If Engine_NPCChat_MeetsConditions(CharIndex, .NPCChatLine, ChatTxt) = 1 Then
                                k = b(j)    'Store the successful value in the k variable for below
                                Exit For    'We got the text!
                            Else
                                Exit Function   'None of the lines worked :(
                            End If
                        End If
                    End If
                
                Next j
  
        End Select

        'Return the value
        Engine_NPCChat_NextLine = k
        
    End With

End Function

Public Function Engine_ClearPath(ByVal UserX As Long, ByVal UserY As Long, ByVal TargetX As Long, ByVal TargetY As Long) As Byte

'***************************************************
'Check if the path is clear from the user to the target of blocked tiles
'For the line-rect collision, we pretend that each tile is 2 units wide so we can give them a width of 1 to center things
'***************************************************
Dim X As Long
Dim Y As Long

    '****************************************
    '***** Target is on top of the user *****
    '****************************************
    
    'If the target position = user position, we must be targeting ourself, so nothing can be blocking us from us (I hope o.O )
    If UserX = TargetX Then
        If UserY = TargetY Then
            Engine_ClearPath = 1
            Exit Function
        End If
    End If

    '********************************************
    '***** Target is right next to the user *****
    '********************************************
    
    'Target is at one of the 4 diagonals of the user
    If Abs(UserX - TargetX) = 1 Then
        If Abs(UserY - TargetY) = 1 Then
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    'Target is above or below the user
    If UserX = TargetX Then
        If Abs(UserY - TargetY) = 1 Then
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    'Target is to the left or right of the user
    If UserY = TargetY Then
        If Abs(UserX - TargetX) = 1 Then
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    '********************************************
    '***** Target is diagonal from the user *****
    '********************************************
    
    'Check if the target is diagonal from the user - only do the following checks if diagonal from the target
    If Abs(UserX - TargetX) = Abs(UserY - TargetY) Then

        If UserX > TargetX Then
                        
            'Diagonal to the top-left
            If UserY > TargetY Then
                For X = TargetX To UserX - 1
                    For Y = TargetY To UserY - 1
                        If MapData(X, Y).BlockedAttack Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    Next Y
                Next X
            
            'Diagonal to the bottom-left
            Else
                For X = TargetX To UserX - 1
                    For Y = UserY + 1 To TargetY
                        If MapData(X, Y).BlockedAttack Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    Next Y
                Next X
            End If

        End If
        
        If UserX < TargetX Then
        
            'Diagonal to the top-right
            If UserY > TargetY Then
                For X = UserX + 1 To TargetX
                    For Y = TargetY To UserY - 1
                        If MapData(X, Y).BlockedAttack Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    Next Y
                Next X
                
            'Diagonal to the bottom-right
            Else
                For X = UserX + 1 To TargetX
                    For Y = UserY + 1 To TargetY
                        If MapData(X, Y).BlockedAttack Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    Next Y
                Next X
            End If
        
        End If
    
        Engine_ClearPath = 1
        Exit Function
    
    End If

    '*******************************************************************
    '***** Target is directly vertical or horizontal from the user *****
    '*******************************************************************
    
    'Check if target is directly above the user
    If UserX = TargetX Then 'Check if x values are the same (straight line between the two)
        If UserY > TargetY Then
            For Y = TargetY + 1 To UserY - 1
                If MapData(UserX, Y).BlockedAttack Then
                    Engine_ClearPath = 0
                    Exit Function
                End If
            Next Y
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    'Check if the target is directly below the user
    If UserX = TargetX Then
        If UserY < TargetY Then
            For Y = UserY + 1 To TargetY - 1
                If MapData(UserX, Y).BlockedAttack Then
                    Engine_ClearPath = 0
                    Exit Function
                End If
            Next Y
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    'Check if the target is directly to the left of the user
    If UserY = TargetY Then
        If UserX > TargetX Then
            For X = TargetX + 1 To UserX - 1
                If MapData(X, UserY).BlockedAttack Then
                    Engine_ClearPath = 0
                    Exit Function
                End If
            Next X
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    'Check if the target is directly to the right of the user
    If UserY = TargetY Then
        If UserX < TargetX Then
            For X = UserX + 1 To TargetX - 1
                If MapData(X, UserY).BlockedAttack Then
                    Engine_ClearPath = 0
                    Exit Function
                End If
            Next X
            Engine_ClearPath = 1
            Exit Function
        End If
    End If

    '***************************************************
    '***** Target is directly not in a direct path *****
    '***************************************************
    
    
    If UserY > TargetY Then
    
        'Check if the target is to the top-left of the user
        If UserX > TargetX Then
            For X = TargetX To UserX
                For Y = TargetY To UserY
                    'We must do * 2 on the tiles so we can use +1 to get the center (its like * 32 and + 16 - this does the same affect)
                    If Engine_Collision_LineRect(X * 2, Y * 2, 2, 2, UserX * 2 + 1, UserY * 2 + 1, TargetX * 2 + 1, TargetY * 2 + 1) Then
                        If MapData(X, Y).BlockedAttack Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    End If
                Next Y
            Next X
            Engine_ClearPath = 1
            Exit Function
    
        'Check if the target is to the top-right of the user
        Else
            For X = UserX To TargetX
                For Y = TargetY To UserY
                    If Engine_Collision_LineRect(X * 2, Y * 2, 2, 2, UserX * 2 + 1, UserY * 2 + 1, TargetX * 2 + 1, TargetY * 2 + 1) Then
                        If MapData(X, Y).BlockedAttack Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    End If
                Next Y
            Next X
        End If
        
    Else
    
        'Check if the target is to the bottom-left of the user
        If UserX > TargetX Then
            For X = TargetX To UserX
                For Y = UserY To TargetY
                    If Engine_Collision_LineRect(X * 2, Y * 2, 2, 2, UserX * 2 + 1, UserY * 2 + 1, TargetX * 2 + 1, TargetY * 2 + 1) Then
                        If MapData(X, Y).BlockedAttack Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    End If
                Next Y
            Next X
        
        'Check if the target is to the bottom-right of the user
        Else
            For X = UserX To TargetX
                For Y = UserY To TargetY
                    If Engine_Collision_LineRect(X * 2, Y * 2, 2, 2, UserX * 2 + 1, UserY * 2 + 1, TargetX * 2 + 1, TargetY * 2 + 1) Then
                        If MapData(X, Y).BlockedAttack Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    End If
                Next Y
            Next X
        End If
    
    End If
    
    Engine_ClearPath = 1

End Function

Public Sub Engine_Render_Skills()

'***************************************************
'Render the spells list
'***************************************************

Const ListWidth As Byte = 10
Dim TempGrh As Grh
Dim i As Byte

    TempGrh.FrameCounter = 1

    'Loop through the skills
    For i = 1 To SkillListSize
        If SkillList(i).SkillID = 0 Then Exit For

        'Render the icon
        TempGrh.GrhIndex = 106
        Engine_Render_Grh TempGrh, SkillList(i).X, SkillList(i).Y, 0, 0, False, GUIColorValue, GUIColorValue, GUIColorValue, GUIColorValue
        TempGrh.GrhIndex = Engine_SkillIDtoGRHID(SkillList(i).SkillID)
        Engine_Render_Grh TempGrh, SkillList(i).X, SkillList(i).Y, 0, 0, False

    Next i

End Sub

Public Sub Engine_Render_Text(ByVal Text As String, ByVal X As Integer, ByVal Y As Integer, ByVal Color As Long)

'************************************************************
'Draw text on D3DDevice
'************************************************************

Dim TempStr() As String
Dim Count As Integer
Dim Ascii As Byte
Dim Row As Integer
Dim u As Single
Dim V As Single
Dim i As Long
Dim j As Long
Dim KeyPhrase As Byte
Dim TempColor As Long
Dim ResetColor As Byte

    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub

    'Get the text into arrays (split by vbCrLf)
    TempStr = Split(Text, vbCrLf)

    'Clear the LastTexture, letting the rest of the engine know that the texture needs to be changed for next rect render
    LastTexture = 0
    
    'Set the texture
    D3DDevice.SetTexture 0, Font_Default.Texture
    
    'Set the temp color (or else the first character has no color)
    TempColor = Color

    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(TempStr)
        Count = 0
        If Len(TempStr(i)) > 0 Then
        
            'Loop through the characters
            For j = 1 To Len(TempStr(i))
            
                'Convert the character to the ascii value
                Ascii = Asc(Mid$(TempStr(i), j, 1))
                
                'Check for a key phrase
                If Ascii = 124 Then 'If Ascii = "|"
                    KeyPhrase = (Not KeyPhrase)
                    If KeyPhrase Then TempColor = D3DColorARGB(255, 255, 0, 0) Else ResetColor = 1
                Else
                    
                    'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
                    Row = (Ascii - Font_Default.HeaderInfo.BaseCharOffset) \ Font_Default.RowPitch
                    u = ((Ascii - Font_Default.HeaderInfo.BaseCharOffset) - (Row * Font_Default.RowPitch)) * Font_Default.ColFactor
                    V = Row * Font_Default.RowFactor
                
                    'Set up the verticies
                    VertexArray(0).Color = TempColor
                    VertexArray(0).X = X + Count
                    VertexArray(0).Y = Y + (Font_Default.CharHeight * i)
                    VertexArray(0).tu = u
                    VertexArray(0).tv = V
                    
                    VertexArray(1).Color = TempColor
                    VertexArray(1).X = X + Count + Font_Default.HeaderInfo.CellWidth
                    VertexArray(1).Y = Y + (Font_Default.CharHeight * i)
                    VertexArray(1).tu = u + Font_Default.ColFactor
                    VertexArray(1).tv = V
                    
                    VertexArray(2).Color = TempColor
                    VertexArray(2).X = X + Count
                    VertexArray(2).Y = Y + Font_Default.HeaderInfo.CellHeight + (Font_Default.CharHeight * i)
                    VertexArray(2).tu = u
                    VertexArray(2).tv = V + Font_Default.RowFactor
                
                    VertexArray(3).Color = TempColor
                    VertexArray(3).X = X + Count + Font_Default.HeaderInfo.CellWidth
                    VertexArray(3).Y = Y + Font_Default.HeaderInfo.CellHeight + (Font_Default.CharHeight * i)
                    VertexArray(3).tu = u + Font_Default.ColFactor
                    VertexArray(3).tv = V + Font_Default.RowFactor
                
                    'Render
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
                    
                    'Shift over the the position to render the next character
                    Count = Count + Font_Default.HeaderInfo.CharWidth(Ascii)
                
                End If
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = Color
                End If
                
            Next j
            
        End If
    Next i

End Sub

Public Sub Engine_SetItemDesc(ByVal name As String, Optional ByVal Amount As Integer = 0, Optional ByVal Price As Long = 0)

'************************************************************
'Set item description values
'************************************************************

Dim i As Byte
Dim X As Long

'Set the item values

    ItemDescLine(1) = name
    ItemDescLines = 1
    If Amount <> 0 Then
        ItemDescLines = ItemDescLines + 1
        ItemDescLine(ItemDescLines) = "Amount: " & Amount
    End If
    If Price <> 0 Then
        ItemDescLines = ItemDescLines + 1
        ItemDescLine(ItemDescLines) = "Price: " & Price
    End If

    'Get the largest size
    ItemDescWidth = Engine_GetTextWidth(ItemDescLine(1))
    If ItemDescLines > 1 Then
        For i = 2 To ItemDescLines
            X = Engine_GetTextWidth(ItemDescLine(i))
            If X > ItemDescWidth Then ItemDescWidth = X
        Next i
    End If

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
                OffsetCounterX = OffsetCounterX - (ScrollPixelsPerFrameX + CharList(UserCharIndex).Speed + (RunningSpeed * CharList(UserCharIndex).Running)) * AddtoUserPos.X * TickPerFrame
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - (ScrollPixelsPerFrameY + CharList(UserCharIndex).Speed + (RunningSpeed * CharList(UserCharIndex).Running)) * AddtoUserPos.Y * TickPerFrame
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
        If ElapsedTime > 200 Then ElapsedTime = 200
        TickPerFrame = (ElapsedTime * EngineBaseSpeed)
        TimerMultiplier = TickPerFrame * 0.075
        If FPSLastCheck + 1000 < timeGetTime Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            FPSLastCheck = timeGetTime
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
        
        'Auto-save config every 30 seconds
        If SaveLastCheck + 30000 < timeGetTime Then
            SaveLastCheck = timeGetTime
            Game_Config_Save
        End If
        
    End If

End Sub

Public Function Engine_SkillIDtoGRHID(ByVal SkillID As Byte) As Integer

'*****************************************************************
'Takes in a SkillID and returns the GrhIndex used for that SkillID
'*****************************************************************

    Select Case SkillID
        Case SkID.Bless: Engine_SkillIDtoGRHID = 46
        Case SkID.IronSkin: Engine_SkillIDtoGRHID = 47
        Case SkID.Strengthen: Engine_SkillIDtoGRHID = 48
        Case SkID.Warcry: Engine_SkillIDtoGRHID = 49
        Case SkID.Protection: Engine_SkillIDtoGRHID = 50
        Case SkID.SpikeField: Engine_SkillIDtoGRHID = 62
        Case SkID.Heal: Engine_SkillIDtoGRHID = 63
    End Select

End Function

Public Function Engine_SkillIDtoSkillName(ByVal SkillID As Byte) As String

'*****************************************************************
'Takes in a SkillID and returns the name of that skill
'*****************************************************************

    Select Case SkillID
        Case SkID.Bless: Engine_SkillIDtoSkillName = "Bless"
        Case SkID.IronSkin: Engine_SkillIDtoSkillName = "Iron Skin"
        Case SkID.Strengthen: Engine_SkillIDtoSkillName = "Strengthen"
        Case SkID.Warcry: Engine_SkillIDtoSkillName = "War Cry"
        Case SkID.Protection: Engine_SkillIDtoSkillName = "Protection"
        Case SkID.SpikeField: Engine_SkillIDtoSkillName = "Spike Field"
        Case SkID.Heal: Engine_SkillIDtoSkillName = "Heal"
        Case Else: Engine_SkillIDtoSkillName = "Unknown Skill"
    End Select

End Function

Public Sub Engine_SortIntArray(TheArray() As Integer, TheIndex() As Integer, ByVal LowerBound As Integer, ByVal UpperBound As Integer)

'*****************************************************************
'Sort an array of integers
'*****************************************************************

Dim s(1 To 64) As Integer   'Stack space for pending Subarrays
Dim indxt As Long   'Stored index
Dim swp As Integer  'Swap variable
Dim F As Integer    'Subarray Minimum
Dim G As Integer    'Subarray Maximum
Dim h As Integer    'Subarray Middle
Dim i As Integer    'Subarray Low  Scan Index
Dim j As Integer    'Subarray High Scan Index
Dim t As Integer    'Stack pointer

'Set the array boundries to f and g

    F = LowerBound
    G = UpperBound

    'Start the loop
    Do

        For j = F + 1 To G
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
        G = s(t)
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
        Unload frm
    Next

End Sub

Function Engine_Distance(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer) As Single

'*****************************************************************
'Finds the distance between two points
'*****************************************************************

    Engine_Distance = Sqr(((Y1 - Y2) ^ 2 + (x1 - x2) ^ 2))
    
End Function

Sub Engine_UseQuickBar(ByVal Slot As Byte)

'******************************************
'Use the object in the quickbar slot
'******************************************

    Select Case QuickBarID(Slot).Type

        'Use an item
    Case QuickBarType_Item
        If QuickBarID(Slot).ID > 0 Then
            sndBuf.Allocate 2
            sndBuf.Put_Byte DataCode.User_Use
            sndBuf.Put_Byte QuickBarID(Slot).ID
        End If

        'Use a skill
    Case QuickBarType_Skill
        If QuickBarID(Slot).ID > 0 Then
            If LastAttackTime + AttackDelay < timeGetTime Then
                If CharList(UserCharIndex).CharStatus.Exhausted = 0 Then
                    LastAttackTime = timeGetTime
                    sndBuf.Allocate 4
                    sndBuf.Put_Byte DataCode.User_CastSkill
                    sndBuf.Put_Byte QuickBarID(Slot).ID
                    sndBuf.Put_Integer TargetCharIndex
                End If
            End If
        End If

    End Select

End Sub

Function Engine_Var_Get(File As String, Main As String, Var As String) As String

'*****************************************************************
'Gets a Var from a text file
'*****************************************************************

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

    szReturn = vbNullString

    sSpaces = Space$(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish
    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File
    Engine_Var_Get = RTrim$(sSpaces)
    If Len(Engine_Var_Get) > 0 Then
        Engine_Var_Get = Left$(Engine_Var_Get, Len(Engine_Var_Get) - 1)
    Else
        Engine_Var_Get = ""
    End If
    
End Function

Sub Engine_Var_Write(File As String, Main As String, Var As String, Value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    writeprivateprofilestring Main, Var, Value, File

End Sub

Public Function Engine_WordWrap(ByVal Text As String, ByVal MaxLineLen As Integer, Optional ByVal ReplaceChar As String = vbCrLf) As String

'************************************************************
'Wrap a long string to multiple lines by vbNewLine
'************************************************************
Dim TempSplit() As String
Dim TSLoop As Long
Dim LastSpace As Long
Dim Size As Long
Dim i As Long
Dim b As Long
Dim j As Long

    'Too small of text
    If Len(Text) < 2 Then
        Engine_WordWrap = Text
        Exit Function
    End If

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(Text, vbNewLine)
    
    For TSLoop = 0 To UBound(TempSplit)
    
        'Clear the values for the new line
        Size = 0
        b = 1
        LastSpace = 1
        
        'Loop through all the characters
        For i = 1 To Len(TempSplit(TSLoop))
        
            'If it is a space, store it so we can easily break at it
            Select Case Mid$(TempSplit(TSLoop), i, 1)
                Case " ": LastSpace = i
                Case "_": LastSpace = i
                Case "-": LastSpace = i
            End Select

            'Add up the size - Do not count the "|" character (high-lighter)!
            If Not Mid$(TempSplit(TSLoop), i, 1) = "|" Then
                Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
            End If
            
            'Check for too large of a size
            If Size > MaxLineLen Then
                
                'Check if the last space was too far back
                If i - LastSpace > 4 Then
                    
                    'Too far away to the last space, so break at the last character
                    Engine_WordWrap = Engine_WordWrap & Trim$(Mid$(TempSplit(TSLoop), b, (i - 1) - b)) & vbNewLine
                    b = i - 1
                    Size = 0
                    
                Else
                
                    'Break at the last space to preserve the word
                    Engine_WordWrap = Engine_WordWrap & Trim$(Mid$(TempSplit(TSLoop), b, LastSpace - b)) & vbNewLine
                    b = LastSpace + 1
                    
                    'Count all the words we ignored (the ones that weren't printed, but are before "i")
                    Size = Engine_GetTextWidth(Mid$(TempSplit(TSLoop), LastSpace, i - LastSpace))
                    
                End If
                
            End If
            
            'This handles the remainder
            If i = Len(TempSplit(TSLoop)) Then
                If b <> i Then
                    Engine_WordWrap = Engine_WordWrap & Mid$(TempSplit(TSLoop), b, i)
                End If
            End If
            
        Next i
        
    Next TSLoop

End Function

Public Function Engine_Music_Load(ByVal FilePath As String, ByVal BufferNumber As Long) As Boolean

'************************************************************
'Loads a mp3 by the specified path
'************************************************************

    On Error GoTo Error_Handler
                
        If Right(FilePath, 4) = ".mp3" Then
        
            Set DirectShow_Control(BufferNumber) = New FilgraphManager
            DirectShow_Control(BufferNumber).RenderFile FilePath
        
            Set DirectShow_Audio(BufferNumber) = DirectShow_Control(BufferNumber)
            
            DirectShow_Audio(BufferNumber).Volume = 0
            DirectShow_Audio(BufferNumber).Balance = 0
        
            Set DirectShow_Event(BufferNumber) = DirectShow_Control(BufferNumber)
            Set DirectShow_Position(BufferNumber) = DirectShow_Control(BufferNumber)
            
            DirectShow_Position(BufferNumber).Rate = 1
            
            DirectShow_Position(BufferNumber).CurrentPosition = 0
                            
        Else
        
            GoTo Error_Handler
        
        End If

    Engine_Music_Load = True
    
    Exit Function
    
Error_Handler:

    Engine_Music_Load = False

End Function

Public Sub Engine_Music_Play(ByVal BufferNumber As Long)

'************************************************************
'Plays the mp3 in the specified buffer
'************************************************************
    On Error GoTo Error_Handler

    DirectShow_Control(BufferNumber).Run

Error_Handler:

End Sub

Public Sub Engine_Music_Stop(ByVal BufferNumber As Long)

'************************************************************
'Stops the mp3 in the specified buffer
'************************************************************

    On Error GoTo Error_Handler
    
    DirectShow_Control(BufferNumber).Stop
    
    DirectShow_Position(BufferNumber).CurrentPosition = 0

    Exit Sub

Error_Handler:

End Sub

Public Sub Engine_Music_Pause(ByVal BufferNumber As Long)

'************************************************************
'Pause the music in the specified buffer
'************************************************************

    On Error GoTo Error_Handler
    
    DirectShow_Control(BufferNumber).Stop
    
Error_Handler:

End Sub

Public Sub Engine_Music_Volume(ByVal Volume As Long, ByVal BufferNumber As Long)

'************************************************************
'Set the volume of the music in the specified buffer
'************************************************************

    On Error GoTo Error_Handler
    
    If Volume >= Music_MaxVolume Then Volume = Music_MaxVolume
    
    If Volume <= 0 Then Volume = 0
    
    DirectShow_Audio(BufferNumber).Volume = (Volume * Music_MaxVolume) - 10000
    
Error_Handler:

End Sub

Public Sub Engine_Music_Balance(ByVal Balance As Long, ByVal BufferNumber As Long)

'************************************************************
'Set the balance of the music in the specified buffer
'************************************************************

    On Error GoTo Error_Handler
    
    If Balance >= Music_MaxBalance Then Balance = Music_MaxBalance
    
    If Balance <= -Music_MaxBalance Then Balance = -Music_MaxBalance
    
    DirectShow_Audio(BufferNumber).Balance = Balance * Music_MaxBalance

Error_Handler:

End Sub

Public Sub Engine_Music_Speed(ByVal Speed As Single, ByVal BufferNumber As Long)

'************************************************************
'Set the speed of the music in the specified buffer
'************************************************************

    On Error GoTo Error_Handler

    If Speed >= Music_MaxSpeed Then Speed = Music_MaxSpeed
    
    If Speed <= 0 Then Speed = 0

    DirectShow_Position(BufferNumber).Rate = Speed / 100

Error_Handler:

End Sub

Public Sub Engine_Music_SetPosition(ByVal Hours As Long, ByVal Minutes As Long, ByVal Seconds As Long, Milliseconds As Single, ByVal BufferNumber As Long)
    
'************************************************************
'Set the speed of the music in the specified buffer
'************************************************************
    
    On Error GoTo Error_Handler
    
    Dim Max_Position As Single
    
    Dim Position As Double
    
    Dim Decimal_Milliseconds As Single
    
    'Keep minutes within range
    
    Minutes = Minutes Mod 60
        
    'Keep seconds within range
    
    Seconds = Seconds Mod 60
        
    'Keep milliseconds within range and keep decimal
    Decimal_Milliseconds = Milliseconds - Int(Milliseconds)
    Milliseconds = Milliseconds Mod 1000
    Milliseconds = Milliseconds + Decimal_Milliseconds
    
    'Convert Minutes & Seconds to Position time
    Position = (Hours * 3600) + (Minutes * 60) + Seconds + (Milliseconds * 0.001)
    
    Max_Position = DirectShow_Position(BufferNumber).StopTime

    If Position >= Max_Position Then
        Position = 0
        GoTo Error_Handler
    End If
    
    If Position <= 0 Then
        Position = 0
        GoTo Error_Handler
    End If
    
    DirectShow_Position(BufferNumber).CurrentPosition = Position

Error_Handler:

End Sub

Public Sub Engine_Music_End(ByVal BufferNumber As Long)

'************************************************************
'End the music in the specified buffer
'************************************************************

    On Error GoTo Error_Handler
    
    'Check if the buffer is looping
    If Not Engine_Music_Loop(BufferNumber) Then
    
        'Check if the current position is past the stop time
        If DirectShow_Position(BufferNumber).CurrentPosition >= DirectShow_Position(BufferNumber).StopTime Then Engine_Music_Stop BufferNumber
    
    End If

Error_Handler:

End Sub

Public Function Engine_Music_Loop(ByVal Media_Number As Long) As Boolean

'************************************************************
'Loop the music in the specified buffer
'************************************************************

    On Error GoTo Error_Handler
    
    'Check if the current position is past the stop time - if so, reset it
    If DirectShow_Position(Media_Number).CurrentPosition >= DirectShow_Position(Media_Number).StopTime Then
        DirectShow_Position(Media_Number).CurrentPosition = 0
    End If
    
    Engine_Music_Loop = True

    Exit Function

Error_Handler:

    Engine_Music_Loop = False

End Function

Public Function Engine_GetBlinkTime() As Long

'************************************************************
'Return a value on how long until the next blink happens
'************************************************************

    Engine_GetBlinkTime = 4000 + Int(Rnd * 5000)
    
End Function

Public Function Engine_RectDistance(ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long, ByVal MaxXDist As Long, ByVal MaxYDist As Long) As Byte

'*****************************************************************
'Check if two tile points are in the same area
'*****************************************************************

    If Abs(x1 - x2) < MaxXDist + 1 Then
        If Abs(Y1 - Y2) < MaxYDist + 1 Then
            Engine_RectDistance = True
        End If
    End If

End Function
