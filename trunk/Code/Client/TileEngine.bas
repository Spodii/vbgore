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

Public AlternateRender As Byte
Public AlternateRenderDefault As Byte
Public AlternateRenderMap As Byte
Public AlternateRenderText As Byte

'Describes a transformable lit vertex
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    Rhw As Single
    Color As Long
    tU As Single
    tV As Single
End Type

'The size of a FVF vertex
Public Const FVF_Size As Long = 28

'********** CONSTANTS ***********
'Keep window in the game screen - dont let them move outside of the window bounds
Public Const WindowsInScreen As Boolean = True

'Screen resolution and information (resolution must be identical to the values on the server!)
Public ReverseSound As Integer      'Reverse the left and right speakers
Public TextureCompress As Long      'Compress textures, saving lots of RAM at an insignifcant CPU usage cost (may reduce graphic quality!)
Public DisableChatBubbles As Byte   'If chat bubbles are drawn or not - chat bubbles can be a huge FPS drainer
Public Bit32 As Byte        'If 32-bit format is used (0 = 16-bit)
Public UseVSync As Byte     'If vertical synchronization copy is used
Public Windowed As Boolean  'If the screen is windowed or fullscreen
Public Const ScreenWidth As Long = 800  'Keep this identical to the value on the server!
Public Const ScreenHeight As Long = 600 'Keep this identical to the value on the server!
Private Const BufferWidth As Long = 1024    'If ScreenWidth is <= 1024, this will = 1024, else set it as 2048
Private Const BufferHeight As Long = 1024   'Same as the BufferWidth, but with the ScreenHeight

'Heading constants
Public Const NORTH As Byte = 1
Public Const EAST As Byte = 2
Public Const SOUTH As Byte = 3
Public Const WEST As Byte = 4
Public Const NORTHEAST As Byte = 5
Public Const SOUTHEAST As Byte = 6
Public Const SOUTHWEST As Byte = 7
Public Const NORTHWEST As Byte = 8

'Font colors
Public Const FontColor_Talk As Long = -1
Public Const FontColor_Info As Long = -16711936
Public Const FontColor_Fight As Long = -65536
Public Const FontColor_Quest As Long = -256
Public Const FontColor_Group As Long = -16711681
Private Const ChatTextBufferSize As Integer = 200
Public Const DamageDisplayTime As Integer = 2000
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

'Point API
Public Type POINTAPI
    X As Long
    Y As Long
End Type

'vbGORE Font Header
Private Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type
Private Type VFH
    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
    CharVA(0 To 255) As CharVA
End Type

Private Type CustomFont
    HeaderInfo As VFH           'Holds the header information
    Texture As Direct3DTexture8 'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As POINTAPI     'Size of the texture
End Type

Public Const Font_Default_TextureNum As Long = -1   'The texture number used to represent this font - only used for AlternateRendering - keep negative to prevent interfering with game textures
Public Font_Default As CustomFont   'Describes our custom font "default"

'********** TYPES ***********

'Text buffer
Type ChatTextBuffer
    Text As String
    Color As Long
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
    CharType As Byte
    OwnerChar As Integer        'If CharType = Slave then this is the index of the owner (used for summoned NPCs to display on the mini-map)
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
    AggressiveCounter As Long
    MoveOffset As FloatPos
    BlinkTimer As Single        'The length of the actual blinking
    StartBlinkTimer As Single   'How long until a blink starts
    ScrollDirectionX As Integer
    ScrollDirectionY As Integer
    BubbleStr As String
    BubbleTime As Long
    Name As String
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
    Warp As Byte
    Sfx As DirectSoundSecondaryBuffer8
End Type

'Hold info about each map
Public Type MapInfo
    Name As String
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

'How many tiles the engine "looks ahead" when drawing the screen
Public TileBufferSize As Integer
Public TileBufferOffset As Long 'Used to calculate offset value in certain cases

'Main view size size in tiles
Public Const WindowTileWidth As Integer = ScreenWidth \ 32
Public Const WindowTileHeight As Integer = ScreenHeight \ 32

'Tile size in pixels
Public Const TilePixelHeight As Integer = 32
Public Const TilePixelWidth As Integer = 32

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public Const ScrollPixelsPerFrameX As Integer = 4
Public Const ScrollPixelsPerFrameY As Integer = 4

'Totals
Private NumBodies As Integer    'Number of bodies
Private NumHeads As Integer     'Number of heads
Private NumHairs As Integer     'Number of hairs
Private NumWeapons As Integer   'Number of weapons
Private NumGrhs As Long         'Number of grhs
Private NumWings As Integer     'Number of wings
Public NumSfx As Integer        'Number of sound effects
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
Public ScreenMinY As Integer    'Start Y pos on current screen
Public ScreenMaxY As Integer    'End Y pos on current screen
Public ScreenMinX As Integer    'Start X pos on current screen
Public ScreenMaxX As Integer    'End X pos on current screen
Public LastTileX As Integer
Public LastTileY As Integer

'********** GAME WINDOWS ***********
Public Const SkillListX As Integer = 750            'Position where the skill list where appear
Public Const SkillListY As Integer = 525            ' (indicates the bottom-right corner)
Public Const SkillListWidth As Integer = 5          'How many skills wide the skill popup list is
Public Const GUIColorValue As Long = -1090519041    'ARGB 190/255/255/255

'Important: Windows are ordered by priority, where 1 = highest!
Public Const AmountWindow As Byte = 1
Public Const MenuWindow As Byte = 2
Public Const NPCChatWindow As Byte = 3
Public Const TradeWindow As Byte = 4
Public Const WriteMessageWindow As Byte = 5
Public Const ViewMessageWindow As Byte = 6
Public Const MailboxWindow As Byte = 7
Public Const InventoryWindow As Byte = 8
Public Const ShopWindow As Byte = 9
Public Const BankWindow As Byte = 10
Public Const StatWindow As Byte = 11
Public Const ChatWindow As Byte = 12
Public Const QuickBarWindow As Byte = 13
Public Const NumGameWindows As Byte = 13

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
Public Const AW_InvToShop As Byte = 2
Public Const AW_InvToBank As Byte = 3
Public Const AW_InvToMail As Byte = 4
Public Const AW_ShopToInv As Byte = 5
Public Const AW_BankToInv As Byte = 6
Public Const AW_InvToTrade As Byte = 7

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
    Message As String
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

Private Type WindowNPCChat
    Screen As Rectangle
    NumAnswers As Byte
    Answer() As Rectangle
    SkinGrh As Grh
End Type

'Info about the trade window
Public Type TradeWindow
    Screen As Rectangle
    User1Name As Rectangle
    User2Name As Rectangle
    Trade1(1 To 9) As Rectangle
    Trade2(1 To 9) As Rectangle
    Gold1 As Rectangle
    Gold2 As Rectangle
    Accept As Rectangle
    Trade As Rectangle
    Cancel As Rectangle
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
    NPCChat As WindowNPCChat
    Trade As TradeWindow
End Type

Public GameWindow As GameWindow

'********** Direct X ***********
Public Const SurfaceTimerMax As Long = 300000       'How long a texture stays in memory unused (miliseconds)
Public SurfaceDB() As Direct3DTexture8          'The list of all the textures
Public SurfaceTimer() As Long                   'How long until the surface unloads
Public LastTexture As Long                      'The last texture used
Public D3DWindow As D3DPRESENT_PARAMETERS       'Describes the viewport and used to restore when in fullscreen
Public UsedCreateFlags As CONST_D3DCREATEFLAGS  'The flags we used to create the device when it first succeeded
Public DispMode As D3DDISPLAYMODE               'Describes the display mode

'Texture for particle effects - this is handled differently then the rest of the graphics
Public ParticleTexture(1 To 12) As Direct3DTexture8

'DirectX 8 Objects
Public DX As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8
Public D3DDevice As Direct3DDevice8

'Used for alternate rendering only
Private Sprite As D3DXSprite
Private SpriteBegun As Byte
Private SpriteScaleVector As D3DVECTOR2

'Motion-bluring information
Public UseMotionBlur As Byte    'If motion blur is enabled or not
Public BlurIntensity As Single
Public BlurTexture As Direct3DTexture8
Public BlurSurf As Direct3DSurface8
Public BlurStencil As Direct3DSurface8
Public DeviceStencil As Direct3DSurface8
Public DeviceBuffer As Direct3DSurface8
Public BlurTA(0 To 3) As TLVERTEX

'Chat vertex buffer (only kept in memory if using alternate rendering)
Private ChatVA() As TLVERTEX

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
End Type

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
Public WeatherSfx1 As DirectSoundSecondaryBuffer8   'Weather buffers - dont add more unless you need more for
Public WeatherSfx2 As DirectSoundSecondaryBuffer8   ' one weather effect (ie rain, wind, lightning)

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
Public Const EngineBaseSpeed As Single = 0.011
Public OffsetCounterX As Single
Public OffsetCounterY As Single

Private NotFirstRender As Boolean

Public ShownText As String

'Weather information
Public LastWeather As Byte
Public UseWeather As Byte

'Mini-map tiles
Public Type MiniMapTile
    X As Single         'X and Y index of the tile (using the tile position, not pixel position)
    Y As Single
    Color As Long       'The color of the tile
End Type
Public MiniMapVBSize As Long    'Size of the vertex buffer (number of verticies, or Tiles x 8)
Public MiniMapVB As Direct3DVertexBuffer8   'Holds the information needed to render the mini-map (not including characters)
Public ShowMiniMap As Byte

'********** OUTSIDE FUNCTIONS ***********
Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Sub Engine_MakeChatBubble(ByVal CharIndex As Integer, ByVal Text As String)
'************************************************************
'Adds text to a chat bubble
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_MakeChatBubble
'************************************************************
    
    If DisableChatBubbles Then Exit Sub
    If LenB(Text) = 0 Then Exit Sub 'No text passed
    CharList(CharIndex).BubbleStr = Text
    CharList(CharIndex).BubbleTime = 5000
    
End Sub

Public Function Engine_TPtoSPX(ByVal X As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_TPtoSPX
'************************************************************

    Engine_TPtoSPX = Engine_PixelPosX(X - minX) + OffsetCounterX - 288 + TileBufferOffset

End Function

Public Function Engine_TPtoSPY(ByVal Y As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_TPtoSPY
'************************************************************

    Engine_TPtoSPY = Engine_PixelPosY(Y - minY) + OffsetCounterY - 288 + TileBufferOffset

End Function

Public Sub Engine_AddToChatTextBuffer(ByVal Text As String, ByVal Color As Long)
'************************************************************
'Adds text to the chat text buffer
'Buffer is order from bottom to top
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_AddToChatTextBuffer
'************************************************************
Dim TempSplit() As String
Dim TSLoop As Long
Dim LastSpace As Long
Dim Size As Long
Dim i As Long
Dim b As Long

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
                    Size = Engine_GetTextWidth(Font_Default, Mid$(TempSplit(TSLoop), LastSpace, i - LastSpace))
 
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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_AddToChatTextBuffer2
'************************************************************
Dim LoopC As Long

    'Move all other text up
    For LoopC = (ChatTextBufferSize - 1) To 1 Step -1
        ChatTextBuffer(LoopC + 1) = ChatTextBuffer(LoopC)
    Next LoopC
    
    'Set the values
    ChatTextBuffer(1).Text = Text
    ChatTextBuffer(1).Color = Color

End Sub

Public Sub Engine_UpdateChatArray()
'************************************************************
'Update the array representing the text in the chat buffer
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_UpdateChatArray
'************************************************************
Dim Chunk As Integer
Dim Count As Integer
Dim LoopC As Byte
Dim Ascii As Byte
Dim Row As Long
Dim Pos As Long
Dim u As Single
Dim v As Single
Dim X As Single
Dim Y As Single
Dim Y2 As Single
Dim i As Long
Dim j As Long
Dim Size As Integer
Dim KeyPhrase As Byte
Dim ResetColor As Byte
Dim TempColor As Long

    On Error Resume Next

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
    If ChatArrayUbound < 0 Then Exit Sub
    ReDim ChatVA(0 To ChatArrayUbound) 'Size our array to fix the 6 verticies of each character

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
        If LenB(ChatTextBuffer(LoopC).Text) <> 0 Then  'Dont bother with empty strings
            
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
                    v = Row * Font_Default.RowFactor

                    'Set up the verticies
                    '    4____5
                    ' 1|\\    |  1 = 4
                    '  | \\   |  3 = 6
                    '  |  \\  |
                    '  |   \\ |
                    ' 2|____\\|
                    '       3 6
                    
                    'Triangle 1
                    With ChatVA(0 + (6 * Pos))   'Top-left corner
                        .Color = TempColor
                        .X = X + Count
                        .Y = Y2
                        .tU = u
                        .tV = v
                        .Rhw = 1
                    End With
                    With ChatVA(1 + (6 * Pos))   'Bottom-left corner
                        .Color = TempColor
                        .X = X + Count
                        .Y = Y2 + Font_Default.HeaderInfo.CellHeight
                        .tU = u
                        .tV = v + Font_Default.RowFactor
                        .Rhw = 1
                    End With
                    With ChatVA(2 + (6 * Pos))   'Bottom-right corner
                        .Color = TempColor
                        .X = X + Count + Font_Default.HeaderInfo.CellWidth
                        .Y = Y2 + Font_Default.HeaderInfo.CellHeight
                        .tU = u + Font_Default.ColFactor
                        .tV = v + Font_Default.RowFactor
                        .Rhw = 1
                    End With
                    
                    'Triangle 2 (only one new verticy is needed)
                    ChatVA(3 + (6 * Pos)) = ChatVA(0 + (6 * Pos)) 'Top-left corner
                    With ChatVA(4 + (6 * Pos))   'Top-right corner
                        .Color = TempColor
                        .X = X + Count + Font_Default.HeaderInfo.CellWidth
                        .Y = Y2
                        .tU = u + Font_Default.ColFactor
                        .tV = v
                        .Rhw = 1
                    End With
                    ChatVA(5 + (6 * Pos)) = ChatVA(2 + (6 * Pos))

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
    
    On Error GoTo 0

    'Check what rendering method we're using
    If AlternateRenderText = 0 Then
    
        'Set the vertex array to the vertex buffer
        If Pos <= 0 Then Pos = 1
        If Not D3DDevice Is Nothing Then   'Make sure the D3DDevice exists - this will only return false if we received messages before it had time to load
            Set ChatVB = D3DDevice.CreateVertexBuffer(FVF_Size * Pos * 6, 0, FVF, D3DPOOL_MANAGED)
            D3DVertexBuffer8SetData ChatVB, 0, FVF_Size * Pos * 6, 0, ChatVA(0)
        End If
        Erase ChatVA()
        
    End If
    
End Sub

Public Sub Engine_Blood_Create(ByVal X As Integer, ByVal Y As Integer)
'*****************************************************************
'Create a blood splatter
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Blood_Create
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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Blood_Erase
'*****************************************************************

    'Clear the selected index
    BloodList(BloodIndex).Grh.FrameCounter = 0
    BloodList(BloodIndex).Grh.GrhIndex = 0
    BloodList(BloodIndex).Pos.X = 0
    BloodList(BloodIndex).Pos.Y = 0

    'Update LastBlood
    If BloodIndex = LastBlood Then
        Do Until BloodList(LastBlood).Grh.GrhIndex > 1

            'Move down one splatter
            LastBlood = LastBlood - 1

            If LastBlood = 0 Then
                Erase BloodList
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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_ChangeHeading
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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Char_Erase
'*****************************************************************

    'Check for targeted character
    If TargetCharIndex = CharIndex Then TargetCharIndex = 0
    If CharIndex = 0 Then Exit Sub
    If CharIndex > LastChar Then Exit Sub
    
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

Function Engine_UserIsFacingChar() As Boolean
'*****************************************************************
'Checks if the user is facing a character - used to check if a character
' is at a tile before making a melee attack
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_UserIsFacingChar
'*****************************************************************
Dim i As Long
Dim X As Long
Dim Y As Long
Dim AddX As Long
Dim AddY As Long

    'Get the co-ordinates of the tile the user is facing
    Select Case CharList(UserCharIndex).Heading
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
    X = CharList(UserCharIndex).Pos.X + AddX
    Y = CharList(UserCharIndex).Pos.Y + AddY
    
    'Make sure the tile is valid
    If X <= 0 Then Exit Function
    If Y <= 0 Then Exit Function
    If X > MapInfo.Width Then Exit Function
    If Y > MapInfo.Height Then Exit Function
    
    'Loop through all the characters
    For i = 1 To LastChar
        If i <> UserCharIndex Then
            
            'Check if the character is located at the tile
            If CharList(i).Pos.X = X Then
                If CharList(i).Pos.Y = Y Then
                    
                    'We have an character here!
                    Engine_UserIsFacingChar = True
                    Exit Function
                    
                End If
            End If
            
        End If
    Next i

End Function

Sub Engine_Char_Make(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Speed As Byte, ByVal Name As String, ByVal Weapon As Integer, ByVal Hair As Integer, ByVal Wings As Integer, ByVal ChatID As Byte, ByVal CharType As Byte, Optional ByVal HP As Byte = 100, Optional ByVal MP As Byte = 100)
'*****************************************************************
'Makes a new character and puts it on the map
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Char_Make
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
    CharList(CharIndex).CharType = CharType
    
    'Update position
    CharList(CharIndex).Pos.X = X
    CharList(CharIndex).Pos.Y = Y

    'Make active
    CharList(CharIndex).Active = 1
    
    'Calculate the name length so we can center the name above the head
    CharList(CharIndex).Name = Name
    CharList(CharIndex).NameOffset = Engine_GetTextWidth(Font_Default, Name) * 0.5

    'Set action index
    CharList(CharIndex).ActionIndex = 0

End Sub

Sub Engine_Char_Move_ByHead(ByVal CharIndex As Integer, ByVal nHeading As Byte, ByVal Running As Byte)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Char_Move_ByHead
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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Char_Move_ByPos
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

Sub Engine_ConvertCPtoTP(ByVal cx As Integer, ByVal cy As Integer, ByRef tX As Integer, ByRef tY As Integer)
'******************************************
'Converts where the user clicks in the main window
'to a tile position
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_ConvertCPtoTP
'******************************************

    tX = UserPos.X + cx \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.Y + cy \ TilePixelHeight - WindowTileHeight \ 2

End Sub

Public Sub Engine_Damage_Create(ByVal X As Integer, ByVal Y As Integer, ByVal Value As Integer)
'*****************************************************************
'Create damage text
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Damage_Create
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
    If Value < 1 Then DamageList(DamageIndex).Value = "Miss" Else DamageList(DamageIndex).Value = Value
    DamageList(DamageIndex).Counter = DamageDisplayTime
    DamageList(DamageIndex).Width = Engine_GetTextWidth(Font_Default, DamageList(DamageIndex).Value)
    DamageList(DamageIndex).Pos.X = X
    DamageList(DamageIndex).Pos.Y = Y

End Sub

Public Sub Engine_Damage_Erase(ByVal DamageIndex As Integer)
'*****************************************************************
'Erase damage text
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Damage_Erase
'*****************************************************************

    'Clear the selected index
    DamageList(DamageIndex).Counter = 0
    DamageList(DamageIndex).Value = vbNullString
    DamageList(DamageIndex).Width = 0

    'Update LastDamage
    If DamageIndex = LastDamage Then
        Do Until DamageList(LastDamage).Counter > 0

            'Move down one splatter
            LastDamage = LastDamage - 1

            If LastDamage = 0 Then
                Erase DamageList
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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Projectile_Create
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

Public Sub Engine_Effect_Create(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Long, Optional ByVal Angle As Single = 0, Optional ByVal Time As Long = 0, Optional ByVal Animated As Byte = 1, Optional ByVal DelayFrames As Single = 0)
'*****************************************************************
'Creates an effect layer for spells and such
'Life is only used if the effect is looped
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Effect_Create
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
    EffectList(EffectIndex).Grh.FrameCounter = 1 - DelayFrames
    
End Sub

Public Sub Engine_Projectile_Erase(ByVal ProjectileIndex As Integer)
'*****************************************************************
'Erase a projectile by the projectile index
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Projectile_Erase
'*****************************************************************
 
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
    If ProjectileIndex = LastProjectile Then
        Do Until ProjectileList(ProjectileIndex).Grh.GrhIndex > 1
            'Move down one projectile
            LastProjectile = LastProjectile - 1
            If LastProjectile = 0 Then Exit Do
        Loop
        If ProjectileIndex <> LastProjectile Then
            'We still have projectiles, resize the array to end at the last used slot
            If LastProjectile > 0 Then
                ReDim Preserve ProjectileList(1 To LastProjectile)
            Else
                Erase ProjectileList
            End If
        End If
    End If
 
End Sub

Public Sub Engine_Effect_Erase(ByVal EffectIndex As Integer)
'*****************************************************************
'Erase an effect by the effect index
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Effect_Erase
'*****************************************************************

    'Clear the selected index
    ZeroMemory EffectList(EffectIndex), Len(EffectList(EffectIndex))

    'Update LastEffect
    If EffectIndex = LastEffect Then
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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_ElapsedTime
'**************************************************************
Dim Start_Time As Long

    'Get current time
    Start_Time = timeGetTime

    'Calculate elapsed time
    Engine_ElapsedTime = Start_Time - EndTime

    'Get next end time
    EndTime = Start_Time

End Function

Function Engine_FileExist(File As String, FileType As VbFileAttribute) As Boolean
'*****************************************************************
'Checks to see if a file exists
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_FileExist
'*****************************************************************
On Error GoTo ErrOut

    If LenB(Dir$(File, FileType)) <> 0 Then Engine_FileExist = True

Exit Function

'An error will most likely be caused by invalid filenames (those that do not follow the file name rules)
ErrOut:

    Engine_FileExist = False
    
End Function

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetAngle
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

Public Function Engine_GetTextWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer
'***************************************************
'Returns the width of text
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetTextWidth
'***************************************************
Dim i As Integer

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
    
    'Loop through the text
    For i = 1 To Len(Text)
        
        'Add up the stored character widths
        Engine_GetTextWidth = Engine_GetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(Mid$(Text, i, 1)))
        
    Next i

End Function

Sub Engine_Init_Signs(ByVal Language As String)
'*****************************************************************
'Loads the sign messages
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_Signs
'*****************************************************************
Dim NumSigns As Integer
Dim LoopC As Integer
Dim s As String

    'Get the number of signs
    NumSigns = Val(Var_Get(SignsPath & "_numsigns.ini", "MAIN", "NumSigns"))
    If NumSigns = 0 Then Exit Sub
    ReDim Signs(1 To NumSigns)
    
    'Grab the English text first
    For LoopC = 1 To NumSigns
        Signs(LoopC) = Trim$(Var_Get(SignsPath & "english.ini", "SIGNS", LoopC))
    Next LoopC
    
    'If we're not using English, grab the foreign language, this way any missing is still presented as English
    If LCase$(Language) <> "english" Then
        For LoopC = 1 To NumSigns
            s = Trim$(Var_Get(SignsPath & LCase$(Language) & ".ini", "SIGNS", LoopC))
            If s <> vbNullString Then Signs(LoopC) = s
        Next LoopC
    End If
    
End Sub

Function Engine_Init_Messages(ByVal Language As String) As String
'*****************************************************************
'Loads the game messages
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_Messages
'*****************************************************************
Dim LoopC As Byte
Dim s As String

    'Make sure we are working in lowercase (since all our files are in lowercase)
    Language = LCase$(Language)
    
    'Check for a redirection flag (will return nothing if the flag doesn't exist)
    Do  'This "Do" will allow us to do redirections to redirections, even though we shouldn't even do that
        s = Var_Get(MessagePath & Language & ".ini", "REDIRECT", "TO")
        If LenB(s) <> 0 Then
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
    NumMessages = CByte(Var_Get(MessagePath & "_nummessages.ini", "MAIN", "NumMessages"))
    
    'Check for a valid number of messages
    If NumMessages = 0 Then
        MsgBox "Error loading message count!", vbOKOnly
        Exit Function
    End If
    
    'Resize our message array to hold all the messages
    ReDim Message(1 To NumMessages)
    
    'Loop through every message and find the message string
    For LoopC = 1 To NumMessages
        Message(LoopC) = Var_Get(MessagePath & Language & ".ini", "MAIN", CStr(LoopC))
        
        'If the message wasn't found, resort to the primary language, English, since that should hold all messages
        If LCase$(Language) <> "english" Then   'Make sure we're not already using English
            If LenB(Trim$(Message(LoopC))) = 0 Then
                Message(LoopC) = Var_Get(MessagePath & "english.ini", "MAIN", CStr(LoopC))
            End If
        End If
        
    Next LoopC
    
    'Load the NPC chat messages
    Engine_Init_NPCChat Language
    
End Function

Sub Engine_Init_BodyData()
'*****************************************************************
'Loads Body.dat
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_BodyData
'*****************************************************************
Dim LoopC As Long
Dim j As Long

'Get number of bodies

    NumBodies = CLng(Var_Get(DataPath & "Body.dat", "INIT", "NumBodies"))
    
    'Resize array
    ReDim BodyData(0 To NumBodies) As BodyData
    
    'Fill list
    For LoopC = 1 To NumBodies
        For j = 1 To 8
            Engine_Init_Grh BodyData(LoopC).Walk(j), CLng(Var_Get(DataPath & "Body.dat", LoopC, j)), 0
            Engine_Init_Grh BodyData(LoopC).Attack(j), CLng(Var_Get(DataPath & "Body.dat", LoopC, "a" & j)), 1
        Next j
        BodyData(LoopC).HeadOffset.X = CLng(Var_Get(DataPath & "Body.dat", LoopC, "HeadOffsetX"))
        BodyData(LoopC).HeadOffset.Y = CLng(Var_Get(DataPath & "Body.dat", LoopC, "HeadOffsetY"))
    Next LoopC

End Sub

Sub Engine_Init_WingData()
'*****************************************************************
'Loads Wing.dat
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_WingData
'*****************************************************************
Dim LoopC As Long
Dim j As Long

    'Get number of wings
    NumWings = CLng(Var_Get(DataPath & "Wing.dat", "INIT", "NumWings"))
    
    'Resize array
    ReDim WingData(0 To NumWings) As WingData
    
    'Fill list
    For LoopC = 1 To NumWings
        For j = 1 To 8
            Engine_Init_Grh WingData(LoopC).Walk(j), CLng(Var_Get(DataPath & "Wing.dat", LoopC, j)), 0
            Engine_Init_Grh WingData(LoopC).Attack(j), CLng(Var_Get(DataPath & "Wing.dat", LoopC, "a" & j)), 1
        Next j
    Next LoopC

End Sub

Private Function Engine_Init_D3DDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS) As Boolean
'************************************************************
'Initialize the Direct3D Device - start off trying with the
'best settings and move to the worst until one works
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_D3DDevice
'************************************************************

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
        If Bit32 = 1 Then DispMode.Format = D3DFMT_X8R8G8B8 Else DispMode.Format = D3DFMT_R5G6B5
        If UseVSync = 1 Then D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC Else D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
        DispMode.Width = ScreenWidth
        DispMode.Height = ScreenHeight
        D3DWindow.BackBufferCount = 1
        D3DWindow.BackBufferFormat = DispMode.Format
        D3DWindow.BackBufferWidth = ScreenWidth
        D3DWindow.BackBufferHeight = ScreenHeight
        D3DWindow.hDeviceWindow = frmMain.hwnd
    End If

    If UseMotionBlur Then
        D3DWindow.EnableAutoDepthStencil = 1
        D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
    End If
    
    'Make sure the form is the correct side
    frmMain.Width = ScreenWidth * Screen.TwipsPerPixelX
    frmMain.Height = ScreenHeight * Screen.TwipsPerPixelY
    
    'Set the D3DDevices
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATEFLAGS, D3DWindow)
    
    'Store the create flags
    UsedCreateFlags = D3DCREATEFLAGS

    'Everything was successful
    Engine_Init_D3DDevice = True
    
    'Force the main form to refresh - vital for widescreen! Remove and find out why if you dare... >:D
    frmMain.Show
    frmMain.Refresh
    DoEvents

Exit Function

ErrOut:

    'Destroy the D3DDevice so it can be remade
    Set D3DDevice = Nothing

    'Return a failure
    Engine_Init_D3DDevice = False

End Function

Sub Engine_Init_Grh(ByRef Grh As Grh, ByVal GrhIndex As Long, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_Grh
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

End Sub

Sub Engine_Init_GrhData()
'*****************************************************************
'Loads Grh.dat
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_GrhData
'*****************************************************************
Dim FileNum As Byte
Dim Grh As Long
Dim Frame As Long

    'Get Number of Graphics
    NumGrhs = CLng(Var_Get(DataPath & "Grh.ini", "INIT", "NumGrhs"))
    
    'Resize arrays
    ReDim GrhData(1 To NumGrhs) As GrhData
    
    'Open files
    FileNum = FreeFile
    Open DataPath & "Grh.dat" For Binary As #FileNum
    Seek #FileNum, 1
    
    'Fill Grh List
    Get #FileNum, , Grh
    
    Do Until Grh <= 0
    
        'Get number of frames
        Get #FileNum, , GrhData(Grh).NumFrames
        If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
        
        If GrhData(Grh).NumFrames > 1 Then
        
            'Read a animation GRH set
            ReDim GrhData(Grh).Frames(1 To GrhData(Grh).NumFrames)
            For Frame = 1 To GrhData(Grh).NumFrames
                Get #FileNum, , GrhData(Grh).Frames(Frame)
                If GrhData(Grh).Frames(Frame) <= 0 Then
                    GoTo ErrorHandler
                End If
            Next Frame
            
            Get #FileNum, , GrhData(Grh).Speed
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
            
            'Compute width and height
            GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
            GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
            GrhData(Grh).Frames(1) = Grh

        End If

        'Get Next Grh Number
        Get #FileNum, , Grh
        
    Loop
    
    Close #FileNum

Exit Sub

ErrorHandler:
    Close #FileNum
    MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh
    IsUnloading = 1

End Sub

Public Sub Engine_Init_GUI(Optional ByVal LoadCustomPos As Byte = 1)
'************************************************************
'Load skin GUI data
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_GUI
'************************************************************
Dim ImageOffsetX As Long
Dim ImageOffsetY As Long
Dim ImageSpaceX As Long
Dim ImageSpaceY As Long
Dim LoopC As Long
Dim s As String 'Stores the path to our master skins file (.ini)
Dim t As String 'Stores the path to our custom window positions file (.dat)
Dim X As Long
Dim Y As Long

    s = DataPath & "Skins\" & CurrentSkin & ".ini"
    t = DataPath & "Skins\" & CurrentSkin & ".dat"
    
    'Load Quickbar
    With GameWindow.QuickBar
        If LoadCustomPos Then
            .Screen.X = Val(Var_Get(t, "QUICKBAR", "ScreenX"))
            .Screen.Y = Val(Var_Get(t, "QUICKBAR", "ScreenY"))
        Else
            .Screen.X = Val(Var_Get(s, "QUICKBAR", "ScreenX"))
            .Screen.Y = Val(Var_Get(s, "QUICKBAR", "ScreenY"))
        End If
        .Screen.Width = Val(Var_Get(s, "QUICKBAR", "ScreenWidth"))
        .Screen.Height = Val(Var_Get(s, "QUICKBAR", "ScreenHeight"))
        Engine_Init_Grh .SkinGrh, Val(Var_Get(s, "QUICKBAR", "Grh"))
    End With
    For LoopC = 1 To 12
        With GameWindow.QuickBar.Image(LoopC)
            .X = Val(Var_Get(s, "QUICKBAR", "Image" & LoopC & "X"))
            .Y = Val(Var_Get(s, "QUICKBAR", "Image" & LoopC & "Y"))
            .Width = 32
            .Height = 32
        End With
    Next LoopC
    
    'Load stats window
    With GameWindow.StatWindow
        If LoadCustomPos Then
            .Screen.X = Val(Var_Get(t, "STATWINDOW", "ScreenX"))
            .Screen.Y = Val(Var_Get(t, "STATWINDOW", "ScreenY"))
        Else
            .Screen.X = Val(Var_Get(s, "STATWINDOW", "ScreenX"))
            .Screen.Y = Val(Var_Get(s, "STATWINDOW", "ScreenY"))
        End If
        .Screen.Width = Val(Var_Get(s, "STATWINDOW", "ScreenWidth"))
        .Screen.Height = Val(Var_Get(s, "STATWINDOW", "ScreenHeight"))
        .AddStr.X = Val(Var_Get(s, "STATWINDOW", "AddStrX"))
        .AddStr.Y = Val(Var_Get(s, "STATWINDOW", "AddStrY"))
        .AddStr.Width = Val(Var_Get(s, "STATWINDOW", "AddStrWidth"))
        .AddStr.Height = Val(Var_Get(s, "STATWINDOW", "AddStrHeight"))
        .AddAgi.X = Val(Var_Get(s, "STATWINDOW", "AddAgiX"))
        .AddAgi.Y = Val(Var_Get(s, "STATWINDOW", "AddAgiY"))
        .AddAgi.Width = Val(Var_Get(s, "STATWINDOW", "AddAgiWidth"))
        .AddAgi.Height = Val(Var_Get(s, "STATWINDOW", "AddAgiHeight"))
        .AddMag.X = Val(Var_Get(s, "STATWINDOW", "AddMagX"))
        .AddMag.Y = Val(Var_Get(s, "STATWINDOW", "AddMagY"))
        .AddMag.Width = Val(Var_Get(s, "STATWINDOW", "AddMagWidth"))
        .AddMag.Height = Val(Var_Get(s, "STATWINDOW", "AddMagHeight"))
        .Str.X = Val(Var_Get(s, "STATWINDOW", "StrX"))
        .Str.Y = Val(Var_Get(s, "STATWINDOW", "StrY"))
        .Agi.X = Val(Var_Get(s, "STATWINDOW", "AgiX"))
        .Agi.Y = Val(Var_Get(s, "STATWINDOW", "AgiY"))
        .Mag.X = Val(Var_Get(s, "STATWINDOW", "MagX"))
        .Mag.Y = Val(Var_Get(s, "STATWINDOW", "MagY"))
        .Gold.X = Val(Var_Get(s, "STATWINDOW", "GoldX"))
        .Gold.Y = Val(Var_Get(s, "STATWINDOW", "GoldY"))
        .DEF.X = Val(Var_Get(s, "STATWINDOW", "DefX"))
        .DEF.Y = Val(Var_Get(s, "STATWINDOW", "DefY"))
        .Dmg.X = Val(Var_Get(s, "STATWINDOW", "DmgX"))
        .Dmg.Y = Val(Var_Get(s, "STATWINDOW", "DmgY"))
        .Points.X = Val(Var_Get(s, "STATWINDOW", "PointsX"))
        .Points.Y = Val(Var_Get(s, "STATWINDOW", "PointsY"))
        Engine_Init_Grh .AddGrh, Val(Var_Get(s, "STATWINDOW", "AddGrh"))
        Engine_Init_Grh .SkinGrh, Val(Var_Get(s, "STATWINDOW", "Grh"))
    End With
    
    'Load chat window
    With GameWindow.ChatWindow
        If LoadCustomPos Then
            .Screen.X = Val(Var_Get(t, "CHATWINDOW", "ScreenX"))
            .Screen.Y = Val(Var_Get(t, "CHATWINDOW", "ScreenY"))
        Else
            .Screen.X = Val(Var_Get(s, "CHATWINDOW", "ScreenX"))
            .Screen.Y = Val(Var_Get(s, "CHATWINDOW", "ScreenY"))
        End If
        .Screen.Width = Val(Var_Get(s, "CHATWINDOW", "ScreenWidth"))
        .Screen.Height = Val(Var_Get(s, "CHATWINDOW", "ScreenHeight"))
        .Text.X = Val(Var_Get(s, "CHATWINDOW", "ChatX"))
        .Text.Y = Val(Var_Get(s, "CHATWINDOW", "ChatY"))
        .Text.Width = Val(Var_Get(s, "CHATWINDOW", "ChatWidth"))
        .Text.Height = Val(Var_Get(s, "CHATWINDOW", "ChatHeight"))
        Engine_Init_Grh .SkinGrh, Val(Var_Get(s, "CHATWINDOW", "Grh"))
    End With

    'Load Inventory
    With GameWindow.Inventory
        If LoadCustomPos Then
            .Screen.X = Val(Var_Get(t, "INVENTORY", "ScreenX"))
            .Screen.Y = Val(Var_Get(t, "INVENTORY", "ScreenY"))
        Else
            .Screen.X = Val(Var_Get(s, "INVENTORY", "ScreenX"))
            .Screen.Y = Val(Var_Get(s, "INVENTORY", "ScreenY"))
        End If
        .Screen.Width = Val(Var_Get(s, "INVENTORY", "ScreenWidth"))
        .Screen.Height = Val(Var_Get(s, "INVENTORY", "ScreenHeight"))
        Engine_Init_Grh .SkinGrh, Val(Var_Get(s, "INVENTORY", "Grh"))
    End With
    ImageOffsetX = Val(Var_Get(s, "INVENTORY", "ImageOffsetX"))
    ImageOffsetY = Val(Var_Get(s, "INVENTORY", "ImageOffsetY"))
    ImageSpaceX = Val(Var_Get(s, "INVENTORY", "ImageSpaceX"))
    ImageSpaceY = Val(Var_Get(s, "INVENTORY", "ImageSpaceY"))
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
            .Screen.X = Val(Var_Get(t, "SHOP", "ScreenX"))
            .Screen.Y = Val(Var_Get(t, "SHOP", "ScreenY"))
        Else
            .Screen.X = Val(Var_Get(s, "SHOP", "ScreenX"))
            .Screen.Y = Val(Var_Get(s, "SHOP", "ScreenY"))
        End If
        Engine_Init_Grh .SkinGrh, Val(Var_Get(s, "SHOP", "Grh"))
    End With
    
    'Load bank window
    GameWindow.Bank = GameWindow.Inventory
    With GameWindow.Bank
        If LoadCustomPos Then
            .Screen.X = Val(Var_Get(t, "BANK", "ScreenX"))
            .Screen.Y = Val(Var_Get(t, "BANK", "ScreenY"))
        Else
            .Screen.X = Val(Var_Get(s, "BANK", "ScreenX"))
            .Screen.Y = Val(Var_Get(s, "BANK", "ScreenY"))
        End If
        Engine_Init_Grh .SkinGrh, Val(Var_Get(s, "BANK", "Grh"))
    End With

    'Load Mailbox window
    With GameWindow.Mailbox.Screen
        If LoadCustomPos Then
            .X = Val(Var_Get(t, "MAILBOX", "ScreenX"))
            .Y = Val(Var_Get(t, "MAILBOX", "ScreenY"))
        Else
            .X = Val(Var_Get(s, "MAILBOX", "ScreenX"))
            .Y = Val(Var_Get(s, "MAILBOX", "ScreenY"))
        End If
        .Width = Val(Var_Get(s, "MAILBOX", "ScreenWidth"))
        .Height = Val(Var_Get(s, "MAILBOX", "ScreenHeight"))
    End With
    Engine_Init_Grh GameWindow.Mailbox.SkinGrh, Val(Var_Get(s, "MAILBOX", "Grh"))
    With GameWindow.Mailbox.WriteLbl
        .X = Val(Var_Get(s, "MAILBOX", "WriteMessageX"))
        .Y = Val(Var_Get(s, "MAILBOX", "WriteMessageY"))
        .Width = Val(Var_Get(s, "MAILBOX", "WriteMessageWidth"))
        .Height = Val(Var_Get(s, "MAILBOX", "WriteMessageHeight"))
    End With
    With GameWindow.Mailbox.DeleteLbl
        .X = Val(Var_Get(s, "MAILBOX", "DeleteMessageX"))
        .Y = Val(Var_Get(s, "MAILBOX", "DeleteMessageY"))
        .Width = Val(Var_Get(s, "MAILBOX", "DeleteMessageWidth"))
        .Height = Val(Var_Get(s, "MAILBOX", "DeleteMessageHeight"))
    End With
    With GameWindow.Mailbox.ReadLbl
        .X = Val(Var_Get(s, "MAILBOX", "ReadMessageX"))
        .Y = Val(Var_Get(s, "MAILBOX", "ReadMessageY"))
        .Width = Val(Var_Get(s, "MAILBOX", "ReadMessageWidth"))
        .Height = Val(Var_Get(s, "MAILBOX", "ReadMessageHeight"))
    End With
    With GameWindow.Mailbox.List
        .X = Val(Var_Get(s, "MAILBOX", "ListX"))
        .Y = Val(Var_Get(s, "MAILBOX", "ListY"))
        .Width = Val(Var_Get(s, "MAILBOX", "ListWidth"))
        .Height = Val(Var_Get(s, "MAILBOX", "ListHeight"))
    End With

    'Load View Message window
    With GameWindow.ViewMessage.Screen
        If LoadCustomPos Then
            .X = Val(Var_Get(t, "VIEWMESSAGE", "ScreenX"))
            .Y = Val(Var_Get(t, "VIEWMESSAGE", "ScreenY"))
        Else
            .X = Val(Var_Get(s, "VIEWMESSAGE", "ScreenX"))
            .Y = Val(Var_Get(s, "VIEWMESSAGE", "ScreenY"))
        End If
        .Width = Val(Var_Get(s, "VIEWMESSAGE", "ScreenWidth"))
        .Height = Val(Var_Get(s, "VIEWMESSAGE", "ScreenHeight"))
    End With
    Engine_Init_Grh GameWindow.ViewMessage.SkinGrh, Val(Var_Get(s, "VIEWMESSAGE", "Grh"))
    With GameWindow.ViewMessage.From
        .X = Val(Var_Get(s, "VIEWMESSAGE", "FromX"))
        .Y = Val(Var_Get(s, "VIEWMESSAGE", "FromY"))
        .Width = Val(Var_Get(s, "VIEWMESSAGE", "FromWidth"))
        .Height = Val(Var_Get(s, "VIEWMESSAGE", "FromHeight"))
    End With
    With GameWindow.ViewMessage.Subject
        .X = Val(Var_Get(s, "VIEWMESSAGE", "SubjectX"))
        .Y = Val(Var_Get(s, "VIEWMESSAGE", "SubjectY"))
        .Width = Val(Var_Get(s, "VIEWMESSAGE", "SubjectWidth"))
        .Height = Val(Var_Get(s, "VIEWMESSAGE", "SubjectHeight"))
    End With
    With GameWindow.ViewMessage.Message
        .X = Val(Var_Get(s, "VIEWMESSAGE", "MessageX"))
        .Y = Val(Var_Get(s, "VIEWMESSAGE", "MessageY"))
        .Width = Val(Var_Get(s, "VIEWMESSAGE", "MessageWidth"))
        .Height = Val(Var_Get(s, "VIEWMESSAGE", "MessageHeight"))
    End With
    ImageOffsetX = Val(Var_Get(s, "VIEWMESSAGE", "ImageOffsetX"))
    ImageOffsetY = Val(Var_Get(s, "VIEWMESSAGE", "ImageOffsetY"))
    ImageSpaceX = Val(Var_Get(s, "VIEWMESSAGE", "ImageSpaceX"))
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
    With GameWindow.WriteMessage.Screen
        If LoadCustomPos Then
            .X = Val(Var_Get(t, "WRITEMESSAGE", "ScreenX"))
            .Y = Val(Var_Get(t, "WRITEMESSAGE", "ScreenY"))
        Else
            .X = Val(Var_Get(s, "WRITEMESSAGE", "ScreenX"))
            .Y = Val(Var_Get(s, "WRITEMESSAGE", "ScreenY"))
        End If
    End With
    Engine_Init_Grh GameWindow.WriteMessage.SkinGrh, Val(Var_Get(s, "WRITEMESSAGE", "Grh"))

    'Load Amount window
    With GameWindow.Amount.Screen
        If LoadCustomPos Then
            .X = Val(Var_Get(t, "AMOUNT", "ScreenX"))
            .Y = Val(Var_Get(t, "AMOUNT", "ScreenY"))
        Else
            .X = Val(Var_Get(s, "AMOUNT", "ScreenX"))
            .Y = Val(Var_Get(s, "AMOUNT", "ScreenY"))
        End If
        .Width = Val(Var_Get(s, "AMOUNT", "ScreenWidth"))
        .Height = Val(Var_Get(s, "AMOUNT", "ScreenHeight"))
    End With
    Engine_Init_Grh GameWindow.Amount.SkinGrh, Val(Var_Get(s, "AMOUNT", "Grh"))
    With GameWindow.Amount.Value
        .X = Val(Var_Get(s, "AMOUNT", "ValueX"))
        .Y = Val(Var_Get(s, "AMOUNT", "ValueY"))
        .Width = Val(Var_Get(s, "AMOUNT", "ValueWidth"))
        .Height = Val(Var_Get(s, "AMOUNT", "ValueHeight"))
    End With

    'Load Menu Window
    With GameWindow.Menu.Screen
        If LoadCustomPos Then
            .X = Val(Var_Get(t, "MENU", "ScreenX"))
            .Y = Val(Var_Get(t, "MENU", "ScreenY"))
        Else
            .X = Val(Var_Get(s, "MENU", "ScreenX"))
            .Y = Val(Var_Get(s, "MENU", "ScreenY"))
        End If
        .Width = Val(Var_Get(s, "MENU", "ScreenWidth"))
        .Height = Val(Var_Get(s, "MENU", "ScreenHeight"))
    End With
    Engine_Init_Grh GameWindow.Menu.SkinGrh, Val(Var_Get(s, "MENU", "Grh"))
    With GameWindow.Menu.QuitLbl
        .X = Val(Var_Get(s, "MENU", "QuitX"))
        .Y = Val(Var_Get(s, "MENU", "QuitY"))
        .Width = Val(Var_Get(s, "MENU", "QuitWidth"))
        .Height = Val(Var_Get(s, "MENU", "QuitHeight"))
    End With
    
    'Load the NPC Chat window
    With GameWindow.NPCChat.Screen
        .X = Val(Var_Get(s, "NPCCHAT", "ScreenX"))
        .Y = Val(Var_Get(s, "NPCCHAT", "ScreenY"))
        .Width = Val(Var_Get(s, "NPCCHAT", "ScreenWidth"))
        .Height = Val(Var_Get(s, "NPCCHAT", "ScreenHeight"))
    End With
    Engine_Init_Grh GameWindow.NPCChat.SkinGrh, Val(Var_Get(s, "NPCCHAT", "Grh"))
    
    'Load the trade window
    With GameWindow.Trade
        .Screen.X = Val(Var_Get(s, "TRADE", "ScreenX"))
        .Screen.Y = Val(Var_Get(s, "TRADE", "ScreenY"))
        .Screen.Width = Val(Var_Get(s, "TRADE", "ScreenWidth"))
        .Screen.Height = Val(Var_Get(s, "TRADE", "ScreenHeight"))
        
        .User1Name.X = Val(Var_Get(s, "TRADE", "User1NameX"))
        .User1Name.Y = Val(Var_Get(s, "TRADE", "User1NameY"))
        
        .User2Name.X = Val(Var_Get(s, "TRADE", "User2NameX"))
        .User2Name.Y = Val(Var_Get(s, "TRADE", "User2NameY"))
        
        .Accept.X = Val(Var_Get(s, "TRADE", "AcceptX"))
        .Accept.Y = Val(Var_Get(s, "TRADE", "AcceptY"))
        .Accept.Width = Val(Var_Get(s, "TRADE", "AcceptWidth"))
        .Accept.Height = Val(Var_Get(s, "TRADE", "AcceptHeight"))
        
        .Trade.X = Val(Var_Get(s, "TRADE", "TradeX"))
        .Trade.Y = Val(Var_Get(s, "TRADE", "TradeY"))
        .Trade.Width = Val(Var_Get(s, "TRADE", "TradeWidth"))
        .Trade.Height = Val(Var_Get(s, "TRADE", "TradeHeight"))
        
        .Cancel.X = Val(Var_Get(s, "TRADE", "CancelX"))
        .Cancel.Y = Val(Var_Get(s, "TRADE", "CancelY"))
        .Cancel.Width = Val(Var_Get(s, "TRADE", "CancelWidth"))
        .Cancel.Height = Val(Var_Get(s, "TRADE", "CancelHeight"))
        
        .Gold1.X = Val(Var_Get(s, "TRADE", "Gold1X"))
        .Gold1.Y = Val(Var_Get(s, "TRADE", "gold1Y"))
        
        .Gold2.X = Val(Var_Get(s, "TRADE", "Gold2X"))
        .Gold2.Y = Val(Var_Get(s, "TRADE", "gold2Y"))
        
        ImageOffsetX = Val(Var_Get(s, "TRADE", "Sec1X"))
        ImageOffsetY = Val(Var_Get(s, "TRADE", "Sec1Y"))
        ImageSpaceX = Val(Var_Get(s, "TRADE", "DividerSize"))
        X = 0
        Y = 0
        
        For LoopC = 1 To 9
            .Trade1(LoopC).X = ImageOffsetX + (X * (ImageSpaceX + 32))
            .Trade1(LoopC).Y = ImageOffsetY + (Y * (ImageSpaceX + 32))
            .Trade1(LoopC).Width = 32
            .Trade1(LoopC).Height = 32
            X = X + 1
            If X = 3 Then
                X = 0
                Y = Y + 1
            End If
        Next LoopC
        ImageOffsetX = Val(Var_Get(s, "TRADE", "Sec2X"))
        ImageOffsetY = Val(Var_Get(s, "TRADE", "Sec2Y"))
        X = 0
        Y = 0
        For LoopC = 1 To 9
            .Trade2(LoopC).X = ImageOffsetX + (X * (ImageSpaceX + 32))
            .Trade2(LoopC).Y = ImageOffsetY + (Y * (ImageSpaceX + 32))
            .Trade2(LoopC).Width = 32
            .Trade2(LoopC).Height = 32
            X = X + 1
            If X = 3 Then
                X = 0
                Y = Y + 1
            End If
        Next LoopC
    
    
    End With
    Engine_Init_Grh GameWindow.Trade.SkinGrh, Val(Var_Get(s, "TRADE", "Grh"))
    
    'Reset text position
    If CurMap > 0 Then Engine_UpdateChatArray

End Sub

Sub Engine_Init_HairData()
'*****************************************************************
'Loads Hair.dat
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_HairData
'*****************************************************************
Dim LoopC As Long
Dim i As Integer

    'Get Number of hairs
    NumHairs = CLng(Var_Get(DataPath & "Hair.dat", "INIT", "NumHairs"))
    
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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_HeadData
'*****************************************************************
Dim LoopC As Long
Dim i As Integer

    'Get Number of heads
    NumHeads = CLng(Var_Get(DataPath & "Head.dat", "INIT", "NumHeads"))
    
    'Resize array
    ReDim HeadData(0 To NumHeads) As HeadData
    
    'Fill List
    For LoopC = 1 To NumHeads
        For i = 1 To 8
            Engine_Init_Grh HeadData(LoopC).Head(i), CLng(Var_Get(DataPath & "Head.dat", LoopC, i)), 0
            Engine_Init_Grh HeadData(LoopC).Blink(i), CLng(Var_Get(DataPath & "Head.dat", LoopC, "b" & i)), 0
            Engine_Init_Grh HeadData(LoopC).AgrHead(i), CLng(Var_Get(DataPath & "Head.dat", LoopC, "a" & i)), 0
            Engine_Init_Grh HeadData(LoopC).AgrBlink(i), CLng(Var_Get(DataPath & "Head.dat", LoopC, "ab" & i)), 0
        Next i
    Next LoopC

End Sub

Public Sub Engine_Init_NPCChat(ByVal Language As String)
'*****************************************************************
'Loads the NPC messages according to the language
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_NPCChat
'*****************************************************************
Dim Conditions() As NPCChatLineCondition
Dim NumConditions As Byte   'The number of conditions
Dim ConditionFlags As Long  'States what conditions are currently used (so we don't have to loop through the Conditions() array)
Dim ChatLine As Byte    'The chat line for the current index
Dim ErrTxt As String    'If there is an error, this extra text is added
Dim HighIndex As Long   'Highest index retrieved
Dim Index As Long       'Current index
Dim FileNum As Byte
Dim ln As String        'Used to grab our lines
Dim Style As Byte       'Style used for the current index
Dim TempSplit() As String
Dim i As Long
Dim F As Long

Dim AskIndex As Byte
Dim HighAskIndex As Long
Dim AnswerIndex As Byte
Dim ln2 As String

    On Error GoTo ErrOut

    'Make sure the file exists
    If Not Engine_FileExist(DataPath & "NPC Chat\" & LCase$(Language) & ".ini", vbNormal) Then
        
        'Error! Change to English before we die!
        Language = "english"
    
    Else
    
        'Load English first, in case any messages are missing from the other language
        If Left$(LCase$(Language), 3) <> "eng" Then Engine_Init_NPCChat "english"

        'Set the initial high-index (to preserve messages from English in case any are missing in the foreign language)
        On Error Resume Next
        HighIndex = UBound(NPCChat)
        On Error GoTo 0

    End If
    
    'Open the file
    FileNum = FreeFile
    Open DataPath & "NPC Chat\" & LCase$(Language) & ".ini" For Input Access Read As FileNum
        
        'Loop until we reach the BEGINFILE line, stating the data is going to start coming in
        Do
            Line Input #FileNum, ln
        Loop While UCase$(Left$(ln, 9)) <> "BEGINFILE"
        
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
                HighAskIndex = 0

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
                Select Case UCase$(TempSplit(0))
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
                        Else
                            For F = 1 To NumConditions
                                If Conditions(F).Condition = NPCCHAT_COND_LEVELLESSTHAN Then
                                    Conditions(F).Value = Val(TempSplit(1))
                                    Exit For
                                End If
                            Next F
                        End If
                    Case "LEVELMORETHAN"
                        If Not ConditionFlags And NPCCHAT_COND_LEVELMORETHAN Then
                            NumConditions = NumConditions + 1
                            ReDim Preserve Conditions(1 To NumConditions)
                            Conditions(NumConditions).Condition = NPCCHAT_COND_LEVELMORETHAN
                            Conditions(NumConditions).Value = Val(TempSplit(1))
                            ConditionFlags = ConditionFlags Or NPCCHAT_COND_LEVELMORETHAN
                        Else
                            For F = 1 To NumConditions
                                If Conditions(F).Condition = NPCCHAT_COND_LEVELMORETHAN Then
                                    Conditions(F).Value = Val(TempSplit(1))
                                    Exit For
                                End If
                            Next F
                        End If
                    Case "HPLESSTHAN"
                        If Not ConditionFlags And NPCCHAT_COND_HPLESSTHAN Then
                            NumConditions = NumConditions + 1
                            ReDim Preserve Conditions(1 To NumConditions)
                            Conditions(NumConditions).Condition = NPCCHAT_COND_HPLESSTHAN
                            Conditions(NumConditions).Value = Val(TempSplit(1))
                            ConditionFlags = ConditionFlags Or NPCCHAT_COND_HPLESSTHAN
                        Else
                            For F = 1 To NumConditions
                                If Conditions(F).Condition = NPCCHAT_COND_HPLESSTHAN Then
                                    Conditions(F).Value = Val(TempSplit(1))
                                    Exit For
                                End If
                            Next F
                        End If
                    Case "HPMORETHAN"
                        If Not ConditionFlags And NPCCHAT_COND_HPMORETHAN Then
                            NumConditions = NumConditions + 1
                            ReDim Preserve Conditions(1 To NumConditions)
                            Conditions(NumConditions).Condition = NPCCHAT_COND_HPMORETHAN
                            Conditions(NumConditions).Value = Val(TempSplit(1))
                            ConditionFlags = ConditionFlags Or NPCCHAT_COND_HPMORETHAN
                        Else
                            For F = 1 To NumConditions
                                If Conditions(F).Condition = NPCCHAT_COND_HPMORETHAN Then
                                    Conditions(F).Value = Val(TempSplit(1))
                                    Exit For
                                End If
                            Next F
                        End If
                    Case "KNOWSKILL"
                        If Not ConditionFlags And NPCCHAT_COND_KNOWSKILL Then
                            NumConditions = NumConditions + 1
                            ReDim Preserve Conditions(1 To NumConditions)
                            Conditions(NumConditions).Condition = NPCCHAT_COND_KNOWSKILL
                            Conditions(NumConditions).Value = Val(TempSplit(1))
                            ConditionFlags = ConditionFlags Or NPCCHAT_COND_KNOWSKILL
                        Else
                            For F = 1 To NumConditions
                                If Conditions(F).Condition = NPCCHAT_COND_KNOWSKILL Then
                                    Conditions(F).Value = Val(TempSplit(1))
                                    Exit For
                                End If
                            Next F
                        End If
                    Case "DONTKNOWSKILL"
                        If Not ConditionFlags And NPCCHAT_COND_DONTKNOWSKILL Then
                            NumConditions = NumConditions + 1
                            ReDim Preserve Conditions(1 To NumConditions)
                            Conditions(NumConditions).Condition = NPCCHAT_COND_DONTKNOWSKILL
                            Conditions(NumConditions).Value = Val(TempSplit(1))
                            ConditionFlags = ConditionFlags Or NPCCHAT_COND_DONTKNOWSKILL
                        Else
                            For F = 1 To NumConditions
                                If Conditions(F).Condition = NPCCHAT_COND_DONTKNOWSKILL Then
                                    Conditions(F).Value = Val(TempSplit(1))
                                    Exit For
                                End If
                            Next F
                        End If
                    Case "SAY"
                        If Not ConditionFlags And NPCCHAT_COND_SAY Then
                            NumConditions = NumConditions + 1
                            ReDim Preserve Conditions(1 To NumConditions)
                            Conditions(NumConditions).Condition = NPCCHAT_COND_SAY  'Notice we UCase$() the next line - this is so we can ignore the case
                            Conditions(NumConditions).ValueStr = UCase$(Replace$(TempSplit(1), "_", " "))   'Replace underscores with spaces
                            ConditionFlags = ConditionFlags Or NPCCHAT_COND_SAY
                        Else
                            For F = 1 To NumConditions
                                If Conditions(F).Condition = NPCCHAT_COND_SAY Then
                                    Conditions(F).ValueStr = UCase$(Replace$(TempSplit(1), "_", " "))
                                    Exit For
                                End If
                            Next F
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
                NPCChat(Index).ChatLine(ChatLine).Text = Replace$(Trim$(TempSplit(2)), "/r", vbNewLine)
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
            
        '*** Look for a STARTASK line ***
        If UCase$(Left$(ln, 9)) = "STARTASK " Then
            NPCChat(Index).Ask.StartAsk = Val(Right$(ln, Len(ln) - 9))
            If NPCChat(Index).Ask.StartAsk <= 0 Then
                ErrTxt = "STARTASK is <= 0"
                GoTo ErrOut
            End If
        End If
        
        '*** Look for an ASK line ***
        If UCase$(Left$(ln, 4)) = "ASK " Then
            
            'Split up the information (0 = "ASK", 1 = ID, 2 = Question text)
            TempSplit() = Split(ln, " ", 3)
            
            'Update the ask information
            AskIndex = Val(TempSplit(1))
            If HighAskIndex < AskIndex Then
                HighAskIndex = AskIndex
                ReDim Preserve NPCChat(Index).Ask.Ask(1 To AskIndex)
                NPCChat(Index).Ask.Ask(AskIndex).Question = Replace$(Trim$(TempSplit(2)), "/r", vbNewLine)
            End If

            'Get the answers
            AnswerIndex = 0
            Do
                Line Input #FileNum, ln2
                ln2 = Trim$(ln2)
                If ln2 <> vbNullString Then
                    If UCase$(Left$(ln2, 6)) = "ASKEND" Then Exit Do
                    If UCase$(Left$(ln2, 7)) = "ANSWER " Then
                        TempSplit() = Split(ln2, " ", 3)
                        If UBound(TempSplit) < 2 Then
                            ErrTxt = "Invalid number of ANSWER parameters!" & """ & ln2 & """
                            GoTo ErrOut
                        End If
                        AnswerIndex = AnswerIndex + 1
                        With NPCChat(Index).Ask.Ask(AskIndex)
                            .NumAnswers = AnswerIndex
                            ReDim Preserve .Answer(1 To AnswerIndex)
                            .Answer(AnswerIndex).Text = Trim$(TempSplit(2))
                            .Answer(AnswerIndex).GotoID = Val(TempSplit(1))
                        End With
                    Else
                        ErrTxt = "Unknown command in ASK block!" & vbNewLine & """ ln2 & """
                        GoTo ErrOut
                    End If
                End If
            Loop
            
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

Sub Engine_Init_ParticleEngine(Optional ByVal SkipToTextures As Boolean = False)
'*****************************************************************
'Loads all particles into memory - unlike normal textures, these stay in memory. This isn't
'done for any reason in particular, they just use so little memory since they are so small
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_ParticleEngine
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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_RenderStates
'************************************************************

    With D3DDevice
        
        'Set the shader to be used
        D3DDevice.SetVertexShader FVF
    
        'Set the render states
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

        'Particle engine settings
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        'Set the texture stage stats (filters)
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
        
    End With

End Sub

Sub Engine_Init_Texture(ByVal TextureNum As Integer)
'*****************************************************************
'Loads a texture into memory
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_Texture
'*****************************************************************
Dim UseTextureFormat As CONST_D3DFORMAT
Dim TexInfo As D3DXIMAGE_INFO_A
Dim FilePath As String

    'Check for a valid texture
    If TextureNum < 1 Then Exit Sub

    'Make sure we even need to load the texture
    If SurfaceTimer(TextureNum) > timeGetTime Then Exit Sub
    
    'Set the texture timer
    SurfaceTimer(TextureNum) = timeGetTime + SurfaceTimerMax

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
        Set SurfaceDB(TextureNum) = D3DX.CreateTextureFromFileEx(D3DDevice, FilePath, D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, TextureCompress, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, TexInfo, ByVal 0)
        SurfaceSize(TextureNum).X = TexInfo.Width
        SurfaceSize(TextureNum).Y = TexInfo.Height
        
    Else
        
        'Set the texture (without getting the dimensions)
        Set SurfaceDB(TextureNum) = D3DX.CreateTextureFromFileEx(D3DDevice, FilePath, SurfaceSize(TextureNum).X, SurfaceSize(TextureNum).Y, 0, 0, TextureCompress, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)
    
    End If

End Sub

Sub Engine_Init_FontTextures()
'*****************************************************************
'Init the custom font textures
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontTextures
'*****************************************************************
Dim TexInfo As D3DXIMAGE_INFO_A

    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    '*** Default font ***
    
    'Set the texture
    Set Font_Default.Texture = D3DX.CreateTextureFromFileEx(D3DDevice, DataPath & "texdefault.png", D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, TexInfo, ByVal 0)
    
    'Store the size of the texture
    Font_Default.TextureSize.X = TexInfo.Width
    Font_Default.TextureSize.Y = TexInfo.Height

End Sub

Sub Engine_Init_FontSettings()
'*****************************************************************
'Init the custom font settings
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontSettings
'*****************************************************************
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single

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
    
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - Font_Default.HeaderInfo.BaseCharOffset) \ Font_Default.RowPitch
        u = ((LoopChar - Font_Default.HeaderInfo.BaseCharOffset) - (Row * Font_Default.RowPitch)) * Font_Default.ColFactor
        v = Row * Font_Default.RowFactor

        'Set the verticies
        With Font_Default.HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).Rhw = 1
            .Vertex(0).tU = u
            .Vertex(0).tV = v
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).Z = 0
            
            .Vertex(1).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).Rhw = 1
            .Vertex(1).tU = u + Font_Default.ColFactor
            .Vertex(1).tV = v
            .Vertex(1).X = Font_Default.HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).Z = 0
            
            .Vertex(2).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).Rhw = 1
            .Vertex(2).tU = u
            .Vertex(2).tV = v + Font_Default.RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = Font_Default.HeaderInfo.CellHeight
            .Vertex(2).Z = 0
            
            .Vertex(3).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).Rhw = 1
            .Vertex(3).tU = u + Font_Default.ColFactor
            .Vertex(3).tV = v + Font_Default.RowFactor
            .Vertex(3).X = Font_Default.HeaderInfo.CellWidth
            .Vertex(3).Y = Font_Default.HeaderInfo.CellHeight
            .Vertex(3).Z = 0
        End With
        
    Next LoopChar

End Sub

Public Sub Engine_Init_TileEngine()
'*****************************************************************
'Init Tile Engine
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_TileEngine
'*****************************************************************
Dim t As Long

    'Size the form
    frmMain.Width = ScreenWidth * Screen.TwipsPerPixelX
    frmMain.Height = ScreenHeight * Screen.TwipsPerPixelY
    
    'Get some engine settings
    UseSfx = Val(Var_Get(DataPath & "Game.ini", "INIT", "UseSfx"))
    If UseSfx <> 0 Then UseSfx = 1      'Force to 1 or 0
    
    UseMusic = Val(Var_Get(DataPath & "Game.ini", "INIT", "UseMusic"))
    If UseMusic <> 0 Then UseMusic = 1  'Force to 1 or 0
    
    UseVSync = Val(Var_Get(DataPath & "Game.ini", "INIT", "VSync"))
    If UseVSync <> 0 Then UseVSync = 1  'Force to 1 or 0

    t = Val(Var_Get(DataPath & "Game.ini", "INIT", "Windowed"))
    If t = 0 Then Windowed = False Else Windowed = True
    
    ReverseSound = Val(Var_Get(DataPath & "Game.ini", "INIT", "ReverseSound"))
    If ReverseSound <> 0 Then ReverseSound = -1 Else ReverseSound = 1   'Force to -1 or 1
    
    TextureCompress = Val(Var_Get(DataPath & "Game.ini", "INIT", "TextureCompression"))
    If TextureCompress <> 0 Then TextureCompress = D3DFMT_DXT5  'Force to 0 or D3DFMT_DXT5
    
    Bit32 = Val(Var_Get(DataPath & "Game.ini", "INIT", "32bit"))
    If Bit32 <> 0 Then Bit32 = 1        'Force to 1 or 0
    
    FPSCap = Val(Var_Get(DataPath & "Game.ini", "INIT", "FPSCap"))
    If FPSCap < 0 Then FPSCap = 0
    If FPSCap > 0 Then FPSCap = 1000 \ FPSCap
    
    DisableChatBubbles = Val(Var_Get(DataPath & "Game.ini", "INIT", "DisableChatBubbles"))
    If DisableChatBubbles <> 0 Then DisableChatBubbles = 1        'Force to 1 or 0
    
    UseWeather = Val(Var_Get(DataPath & "Game.ini", "INIT", "UseWeather"))
    If UseWeather <> 0 Then UseWeather = 1
    
    UseMotionBlur = Val(Var_Get(DataPath & "Game.ini", "INIT", "UseMotionBlur"))
    If UseMotionBlur <> 0 Then UseMotionBlur = 1
    
    'Load the key definitions
    Input_Keys_LoadDefinitions
    
    '****** INIT DirectX ******
    ' Create the root D3D objects
    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate()
    Set D3DX = New D3DX8
    Input_Init
    Sound_Init

    'Create the D3D Device
    If Not Engine_Init_D3DDevice(D3DCREATE_PUREDEVICE) Then
        If Not Engine_Init_D3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
            If Not Engine_Init_D3DDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                If Not Engine_Init_D3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                    MsgBox "Could not init D3DDevice. Exiting..."
                    Engine_Init_UnloadTileEngine
                    Engine_UnloadAllForms
                    End
                End If
            End If
        End If
    End If
    Engine_Init_RenderStates
    
    'Load the rest of the tile engine stuff
    Engine_Init_FontTextures
    Engine_Init_ParticleEngine
    
    'Create the needed information for the motion bluring
    If UseMotionBlur Then
        Set DeviceBuffer = D3DDevice.GetRenderTarget
        Set DeviceStencil = D3DDevice.GetDepthStencilSurface
        Set BlurStencil = D3DDevice.CreateDepthStencilSurface(BufferWidth, BufferHeight, D3DFMT_D16, D3DMULTISAMPLE_NONE)
        Set BlurTexture = D3DX.CreateTexture(D3DDevice, BufferWidth, BufferHeight, 0, D3DUSAGE_RENDERTARGET, DispMode.Format, D3DPOOL_DEFAULT)
        Set BlurSurf = BlurTexture.GetSurfaceLevel(0)
        
        'Create the motion-blur vertex array
        For t = 0 To 3
            BlurTA(t).Color = D3DColorXRGB(255, 255, 255)
            BlurTA(t).Rhw = 1
        Next t
        BlurTA(1).X = ScreenWidth
        BlurTA(2).Y = ScreenHeight
        BlurTA(3).X = ScreenWidth
        BlurTA(3).Y = ScreenHeight
        
    End If

    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    'Set the ending time to now (to prevent the client thinking there was a huge FPS jump)
    EndTime = timeGetTime

    'Get the AlternateRender flag
    AlternateRender = Val(Var_Get(DataPath & "Game.ini", "INIT", "AlternateRender"))
    AlternateRenderMap = Val(Var_Get(DataPath & "Game.ini", "INIT", "AlternateRenderMap"))
    AlternateRenderText = Val(Var_Get(DataPath & "Game.ini", "INIT", "AlternateRenderText"))
    If AlternateRender <> 0 Then AlternateRender = 1
    If AlternateRenderMap <> 0 Then AlternateRenderMap = 1
    If AlternateRenderText <> 0 Then AlternateRenderText = 1
    AlternateRenderDefault = AlternateRender
    
    'Set the blur to off
    BlurIntensity = 255

    If AlternateRender = 1 Or AlternateRenderMap = 1 Or AlternateRenderText = 1 Then

        'If using alternate rendering, create the sprite object
        Set Sprite = D3DX.CreateSprite(D3DDevice)
        
        'Set the scaling to default aspect ratio
        SpriteScaleVector.X = 1
        SpriteScaleVector.Y = 1
        
    End If
    
    'Start the engine
    EngineRun = True

End Sub

Public Sub Engine_Init_UnloadTileEngine()
'*****************************************************************
'Shutsdown engine
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_UnloadTileEngine
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
    For X = 1 To MapInfo.Width
        For Y = 1 To MapInfo.Height
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
    
    'Clear motion blur objects
    If Not BlurTexture Is Nothing Then
        Set BlurTexture = Nothing
        Set BlurSurf = Nothing
        Set BlurStencil = Nothing
        Set DeviceStencil = Nothing
        Set DeviceBuffer = Nothing
    End If
    
    'Clear arrays
    Erase BlurTA
    Erase SurfaceTimer
    Erase SoundBufferTimer
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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_WeaponData
'*****************************************************************
Dim LoopC As Long
    
    'Get number of weapons
    NumWeapons = CLng(Var_Get(DataPath & "Weapon.dat", "INIT", "NumWeapons"))
    
    'Resize array
    ReDim WeaponData(0 To NumWeapons) As WeaponData
    
    'Fill list
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

Sub Engine_Weather_UpdateFog()
'*****************************************************************
'Update the fog effects
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Weather_UpdateFog
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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Weather_UpdateLightning
'*****************************************************************
Dim X As Long
Dim Y As Long
Dim i As Long

    'Check if we are in the middle of a flash
    If FlashTimer > 0 Then
        FlashTimer = FlashTimer - ElapsedTime
        
        'The flash has run out
        If FlashTimer <= 0 Then
        
            'Change the light of all the tiles back
            For X = 1 To MapInfo.Width
                For Y = 1 To MapInfo.Height
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
            
            'Sound effect
            Sound_Play WeatherSfx2, DSBPLAY_DEFAULT  'BAM!
            
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

Sub Engine_Weather_Update()
'*****************************************************************
'Initializes the weather effects
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Weather_Update
'*****************************************************************

    'Check if we're using weather
    If UseWeather = 0 Then Exit Sub

    'Only update the weather settings if it has changed!
    If LastWeather <> MapInfo.Weather Then
    
        'Set the lastweather to the current weather
        LastWeather = MapInfo.Weather
        
        'Erase sounds
        Sound_Erase WeatherSfx1
        Sound_Erase WeatherSfx2
    
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
            LightningTimer = 15000 + (Rnd * 15000)
            WeatherDoLightning = 1  'We take our rain with a bit of lightning on top >:D
            WeatherDoFog = 0
            Sound_Set WeatherSfx1, 3
            Sound_Set WeatherSfx2, 2
            Sound_Play WeatherSfx1, DSBPLAY_LOOPING
            
        Case 3  'Inside of a house in a storm (lightning + muted rain sound)
            If WeatherEffectIndex > 0 Then  'Kill the weather effect if used
                If Effect(WeatherEffectIndex).Used Then Effect_Kill WeatherEffectIndex
            End If
            LightningTimer = 15000 + (Rnd * 15000)
            WeatherDoLightning = 1
            WeatherDoFog = 0
            Sound_Set WeatherSfx1, 4
            Sound_Set WeatherSfx2, 6
            Sound_Play WeatherSfx1, DSBPLAY_LOOPING
            
        Case 4  'Inside of a cave in a storm (lightning + muted rain sound + fog)
            If WeatherEffectIndex > 0 Then  'Kill the weather effect if used
                If Effect(WeatherEffectIndex).Used Then Effect_Kill WeatherEffectIndex
            End If
            LightningTimer = 15000 + (Rnd * 15000)
            WeatherDoLightning = 1
            WeatherDoFog = 10    'This will make it nice and spooky! >:D
            Sound_Set WeatherSfx1, 4
            Sound_Set WeatherSfx2, 6
            Sound_Play WeatherSfx1, DSBPLAY_LOOPING
            
        Case Else   'None
            If WeatherEffectIndex > 0 Then  'Kill the weather effect if used
                If Effect(WeatherEffectIndex).Used Then Effect_Kill WeatherEffectIndex
                Sound_Erase WeatherSfx1  'Remove the sounds
                Sound_Erase WeatherSfx2
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

Sub Engine_ShowNPCChatWindow(ByVal NPCName As String, ByVal ChatIndex As Byte, ByVal AskIndex As Byte)
'*****************************************************************
'Shows the NPC chat window
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_ShowNPCChatWindow
'*****************************************************************
Dim i As Long
Dim Offset As Long

    'Set the window values
    ActiveAsk.AskIndex = AskIndex
    ActiveAsk.ChatIndex = ChatIndex
    ActiveAsk.AskName = NPCName
    ActiveAsk.QuestionTxt = NPCName & ": " & vbNewLine & Engine_WordWrap(NPCChat(ChatIndex).Ask.Ask(AskIndex).Question, GameWindow.NPCChat.Screen.Width - 10)

    'Set the window information
    With GameWindow.NPCChat
        .NumAnswers = NPCChat(ChatIndex).Ask.Ask(AskIndex).NumAnswers
        ReDim .Answer(1 To .NumAnswers)
        
        Offset = .Screen.Height - 5
        For i = .NumAnswers To 1 Step -1
            Offset = Offset - Font_Default.CharHeight
            .Answer(i).Y = Offset
            .Answer(i).Height = Font_Default.CharHeight
            .Answer(i).X = 5
            .Answer(i).Width = Engine_GetTextWidth(Font_Default, i & ". " & NPCChat(ChatIndex).Ask.Ask(AskIndex).Answer(i).Text)
        Next i
        
    End With
    
    ShowGameWindow(NPCChatWindow) = 1
    LastClickedWindow = NPCChatWindow
    SelGameWindow = NPCChatWindow

End Sub

Function Engine_LegalPos(ByVal X As Integer, ByVal Y As Integer, ByVal Heading As Byte) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_LegalPos
'*****************************************************************
Dim i As Integer

    'Check that it is in the map
    If X < 1 Then Exit Function
    If X > MapInfo.Width Then Exit Function
    If Y < 1 Then Exit Function
    If Y > MapInfo.Height Then Exit Function

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
                If CharList(i).Pos.Y = Y Then
                    If CharList(i).OwnerChar <> UserCharIndex Then
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i

    'The position is legal
    Engine_LegalPos = True

End Function

Sub Engine_MoveScreen(ByVal Heading As Byte)
'*****************************************************************
'Starts the screen moving in a direction
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_MoveScreen
'*****************************************************************
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
    
    If tX < 1 Then tX = 1: If X < 0 Then X = 0
    If tX > MapInfo.Width Then tX = MapInfo.Width: If X > 0 Then X = 0
    If tY < 1 Then tY = 1: If Y < 0 Then Y = 0
    If tY > MapInfo.Height Then tY = MapInfo.Height: If Y > 0 Then Y = 0

    'Start moving... MainLoop does the rest
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = True

End Sub

Sub Engine_MoveUser(ByVal Direction As Byte)
'*****************************************************************
'Move user in appropriate direction
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_MoveUser
'*****************************************************************
Dim ax As Integer
Dim ay As Integer
Dim aX2 As Integer
Dim aY2 As Integer
Dim aX3 As Integer
Dim aY3 As Integer
Dim Direction2 As Byte
Dim Direction3 As Byte

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
            ax = 1
            ay = -1
            aX2 = 0
            aY2 = -1
            aX3 = 1
            aY3 = 0
            Direction2 = NORTH
            Direction3 = EAST
        Case NORTHWEST
            ax = -1
            ay = -1
            aX2 = 0
            aY2 = -1
            aX3 = -1
            aY3 = 0
            Direction2 = NORTH
            Direction3 = WEST
        Case SOUTHEAST
            ax = 1
            ay = 1
            aX2 = 0
            aY2 = 1
            aX3 = 1
            aY3 = 0
            Direction2 = SOUTH
            Direction3 = EAST
        Case SOUTHWEST
            ax = -1
            ay = 1
            aX2 = 0
            aY2 = 1
            aX3 = -1
            aY3 = 0
            Direction2 = SOUTH
            Direction3 = WEST
        Case NORTH
            ax = 0
            ay = -1
        Case EAST
            ax = 1
            ay = 0
        Case SOUTH
            ax = 0
            ay = 1
        Case WEST
            ax = -1
            ay = 0
    End Select

    'If the shop, mailbox or read mail window are showing, hide them
    ShowGameWindow(MailboxWindow) = 0
    ShowGameWindow(ShopWindow) = 0
    ShowGameWindow(ViewMessageWindow) = 0
    ShowGameWindow(AmountWindow) = 0
    ShowGameWindow(BankWindow) = 0
    If LastClickedWindow = MailboxWindow Or LastClickedWindow = ShopWindow Or LastClickedWindow = ViewMessageWindow Or _
        LastClickedWindow = AmountWindow Or LastClickedWindow = BankWindow Then LastClickedWindow = 0
    AmountWindowUsage = 0
    AmountWindowItemIndex = 0
    AmountWindowValue = vbNullString

    'Try the first movement
    If Engine_LegalPos(UserPos.X + ax, UserPos.Y + ay, Direction) Then
        Engine_SendMovePacket Direction
        Exit Sub
    End If
    
    'If the first movement failed, use the second and third if a diagonal direction
    If Direction2 > 0 Then
        If Engine_LegalPos(UserPos.X + aX2, UserPos.Y + aY2, Direction) Then
            Engine_SendMovePacket Direction2
            Exit Sub
        End If
        If Engine_LegalPos(UserPos.X + aX3, UserPos.Y + aY3, Direction3) Then
            Engine_SendMovePacket Direction3
            Exit Sub
        End If
    End If

    'Movement failed, rotate the user to face the direction if needed
    'Only rotate if the user is not already facing that direction
    If CharList(UserCharIndex).Heading <> Direction Then
        sndBuf.Allocate 2
        sndBuf.Put_Byte DataCode.User_Rotate
        sndBuf.Put_Byte Direction
    End If

End Sub

Sub Engine_SendMovePacket(ByVal Direction As Byte)
'*****************************************************************
'Sends the user's movement packet to the server
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_SendMovePacket
'*****************************************************************
Dim Running As Byte

    'If running
    If GetAsyncKeyState(vbKeyShift) Then
    
        'Check if the user has enough stamina to run
        If BaseStats(SID.MinSTA) > RunningCost Then Running = 1
        
    End If

    'Send the information to the server
    sndBuf.Allocate 2
    sndBuf.Put_Byte DataCode.User_Move
    
    'Running or not
    If Running = 1 Then sndBuf.Put_Byte Direction Or 128 Else sndBuf.Put_Byte Direction

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
    Sound_UpdateMap
    
End Sub

Sub Engine_OBJ_Create(ByVal GrhIndex As Long, ByVal X As Byte, ByVal Y As Byte)
'*****************************************************************
'Create an object on the map and update LastOBJ value
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_OBJ_Create
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

Function Engine_OBJ_AtTile(ByVal X As Byte, ByVal Y As Byte) As Boolean
'*****************************************************************
'Checks for an object at tile (X,Y)
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_OBJ_AtTile
'*****************************************************************
Dim i As Long

    'Check if any objects exist
    If LastObj = 0 Then Exit Function

    'Loop through all the objects
    For i = 1 To LastObj
        
        'Check if the object is located at the tile
        If OBJList(i).Pos.X = X Then
            If OBJList(i).Pos.Y = Y Then
                
                'We have an object here!
                Engine_OBJ_AtTile = True
                Exit Function
                
            End If
        End If
        
    Next i

End Function

Sub Engine_OBJ_Erase(ByVal ObjIndex As Integer)
'*****************************************************************
'Erase an object from the map and update the LastOBJ value
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_OBJ_Erase
'*****************************************************************

    'Check for a valid object
    If ObjIndex > LastObj Then Exit Sub
    If ObjIndex <= 0 Then Exit Sub

    'Erase the object
    ZeroMemory OBJList(ObjIndex), Len(OBJList(ObjIndex))

    'Update LastOBJ
    If ObjIndex = LastObj Then
        Do Until OBJList(LastObj).Grh.GrhIndex > 1
            'Move down one object
            LastObj = LastObj - 1
            If LastObj = 0 Then Exit Do
        Loop
        If ObjIndex <> LastObj Then
            'We still have objects, resize the array to end at the last used slot
            If LastObj <> 0 Then
                ReDim Preserve OBJList(1 To LastObj)
            Else
                Erase OBJList
            End If
        End If
    End If

End Sub

Function Engine_PixelPosX(ByVal X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_PixelPosX
'*****************************************************************

    Engine_PixelPosX = (X - 1) * TilePixelWidth

End Function

Function Engine_PixelPosY(ByVal Y As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_PixelPosY
'*****************************************************************

    Engine_PixelPosY = (Y - 1) * TilePixelHeight

End Function

Private Function Engine_Collision_Between(ByVal Value As Single, ByVal Bound1 As Single, ByVal Bound2 As Single) As Byte
'*****************************************************************
'Find if a value is between two other values (used for line collision)
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Between
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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Line
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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_LineRect
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

Function Engine_Collision_Rect(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal Width1 As Integer, ByVal Height1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer, ByVal Width2 As Integer, ByVal Height2 As Integer) As Boolean
'*****************************************************************
'Check for collision between two rectangles
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Rect
'*****************************************************************

    If x1 + Width1 >= x2 Then
        If x1 <= x2 + Width2 Then
            If Y1 + Height1 >= Y2 Then
                If Y1 <= Y2 + Height2 Then
                    Engine_Collision_Rect = True
                End If
            End If
        End If
    End If

End Function

Private Sub Engine_Render_Char(ByVal CharIndex As Long, ByVal PixelOffsetX As Single, ByVal PixelOffsetY As Single)
'*****************************************************************
'Draw a character to the screen by the CharIndex
'First variables are set, then all shadows drawn, then character drawn, then extras (emoticons, icons, etc)
'Any variables not handled in "Set the variables" are set in Shadow calls - do not call a second time in the
'normal character rendering calls
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Render_Char
'*****************************************************************
Dim TempGrh As Grh
Dim Moved As Boolean
Dim IconCount As Byte
Dim IconOffset As Integer
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
    If CharList(CharIndex).Pos.Y > 1 Then
        TempBlock2 = MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y - 1)
    Else
        TempBlock2 = TempBlock
    End If
    
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
    
    'Update aggressive timer
    If CharList(CharIndex).Aggressive > 0 Then
        If CharList(CharIndex).AggressiveCounter < timeGetTime Then
            CharList(CharIndex).Aggressive = 0
            CharList(CharIndex).AggressiveCounter = 0
        End If
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
    Engine_Render_Text Font_Default, CharList(CharIndex).Name, PixelOffsetX + 16 - CharList(CharIndex).NameOffset, PixelOffsetY - 40, RenderColor(1)

    'Count the number of icons that will be needed to draw
    With CharList(CharIndex).CharStatus
        IconCount = 0
        IconCount = .Blessed + .Protected + .Strengthened + .Cursed + .WarCursed + .IronSkinned + .Exhausted
    End With
    
    'Health/Mana bars
    Engine_Render_Rectangle PixelOffsetX - 4, PixelOffsetY + 34, (CharList(CharIndex).HealthPercent / 100) * 40, 4, 1, 1, 1, 1, 1, 1, 0, 0, HealthColor, HealthColor, HealthColor, HealthColor, 0, False
    Engine_Render_Rectangle PixelOffsetX - 4, PixelOffsetY + 38, (CharList(CharIndex).ManaPercent / 100) * 40, 4, 1, 1, 1, 1, 1, 1, 0, 0, ManaColor, ManaColor, ManaColor, ManaColor, 0, False

    'Draw the icons
    If IconCount > 0 Then

        'Calculate the icon offset
        IconOffset = PixelOffsetX + 16 - (IconCount * 8)

        If CharList(CharIndex).CharStatus.Blessed Then
            Engine_Init_Grh TempGrh, 15
            Engine_Render_Grh TempGrh, IconOffset, PixelOffsetY - 50, 0, 0, False
            IconOffset = IconOffset + 16
        End If
        If CharList(CharIndex).CharStatus.Protected Then
            Engine_Init_Grh TempGrh, 20
            Engine_Render_Grh TempGrh, IconOffset, PixelOffsetY - 50, 0, 0, False
            IconOffset = IconOffset + 16
        End If
        If CharList(CharIndex).CharStatus.Strengthened Then
            Engine_Init_Grh TempGrh, 17
            Engine_Render_Grh TempGrh, IconOffset, PixelOffsetY - 50, 0, 0, False
            IconOffset = IconOffset + 16
        End If
        If CharList(CharIndex).CharStatus.Cursed Then
            Engine_Init_Grh TempGrh, 18
            Engine_Render_Grh TempGrh, IconOffset, PixelOffsetY - 50, 0, 0, False
            IconOffset = IconOffset + 16
        End If
        If CharList(CharIndex).CharStatus.WarCursed Then
            Engine_Init_Grh TempGrh, 19
            Engine_Render_Grh TempGrh, IconOffset, PixelOffsetY - 50, 0, 0, False
            IconOffset = IconOffset + 16
        End If
        If CharList(CharIndex).CharStatus.IronSkinned Then
            Engine_Init_Grh TempGrh, 16
            Engine_Render_Grh TempGrh, IconOffset, PixelOffsetY - 50, 0, 0, False
            IconOffset = IconOffset + 16
        End If
        If CharList(CharIndex).CharStatus.Exhausted Then
            Engine_Init_Grh TempGrh, 22
            Engine_Render_Grh TempGrh, IconOffset, PixelOffsetY - 50, 0, 0, False
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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Render_ChatTextBuffer
'************************************************************
Dim SrcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim i As Long

    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
    
    'Assign the alternate rendering value
    AlternateRender = AlternateRenderText

    'Check if using alternate rendering
    If AlternateRender Then

        'Check if the texture is no longer the font texture
        If LastTexture <> Font_Default_TextureNum Then
        
            'End the old sprite we had going
            If SpriteBegun = 1 Then
                Sprite.End
                Sprite.Begin
            End If
            
        End If
        
        'Loop through all the characters
        For i = 0 To UBound(ChatVA) Step 6

            'Create the source rectangle
            With SrcRect
                .Left = ChatVA(i).tU * Font_Default.TextureSize.X
                .Top = ChatVA(i).tV * Font_Default.TextureSize.Y
                .Right = ChatVA(i + 5).tU * Font_Default.TextureSize.X
                .bottom = ChatVA(i + 5).tV * Font_Default.TextureSize.Y
            End With
            
            'Set the translation (location on the screen)
            v3.X = ChatVA(i).X
            v3.Y = ChatVA(i).Y
            
            'Draw the character
            Sprite.Draw Font_Default.Texture, SrcRect, SpriteScaleVector, v2, 0, v3, ChatVA(i).Color
    
        Next i

    Else
        
        'Clear the LastTexture, letting the rest of the engine know that the texture needs to be changed for next rect render
        D3DDevice.SetTexture 0, Font_Default.Texture
        LastTexture = Font_Default_TextureNum
    
        'Set up the vertex buffer
        If ShowGameWindow(ChatWindow) Then
            If ChatArrayUbound > 0 Then
                D3DDevice.SetStreamSource 0, ChatVB, FVF_Size
                D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
            End If
        End If
    
    End If
    
    'Retreive the default alternate render value
    AlternateRender = AlternateRenderDefault

End Sub

Private Function Engine_UpdateGrh(ByRef Grh As Grh, Optional ByVal LoopAnim As Boolean = True) As Boolean
'*****************************************************************
'Updates the grh's animation
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_UpdateGrh
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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Render_Grh
'*****************************************************************
Dim CurrGrhIndex As Long    'The grh index we will be working with (acquired after updating animations)
Dim FileNum As Integer

    'Check to make sure it is legal
    If Grh.GrhIndex < 1 Then Exit Sub
    If GrhData(Grh.GrhIndex).NumFrames < 1 Then Exit Sub
    If Grh.FrameCounter < 1 Then
        'Grh has a delay, so just update the frame and then leave
        Engine_UpdateGrh Grh, LoopAnim
        Exit Sub
    End If
    If Int(Grh.FrameCounter) > GrhData(Grh.GrhIndex).NumFrames Then Grh.FrameCounter = 1
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrGrhIndex = GrhData(Grh.GrhIndex).Frames(Int(Grh.FrameCounter))

    'Check for in-bounds
    If X + GrhData(CurrGrhIndex).pixelWidth > 0 Then
        If Y + GrhData(CurrGrhIndex).pixelHeight > 0 Then
            If X < ScreenWidth Then
                If Y < ScreenHeight Then
                
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
                    If AlternateRender = 0 Then
                    
                        'Render the texture with 2 triangles on a triangle strip
                        Engine_Render_Rectangle X, Y, GrhData(CurrGrhIndex).pixelWidth, GrhData(CurrGrhIndex).pixelHeight, GrhData(CurrGrhIndex).SX, _
                            GrhData(CurrGrhIndex).SY, GrhData(CurrGrhIndex).pixelWidth, GrhData(CurrGrhIndex).pixelHeight, , , Angle, FileNum, Light1, Light2, Light3, Light4, Shadow, False
                        
                    Else
                        
                        'Render the texture as a D3DXSprite
                        Engine_Render_D3DXSprite X, Y, GrhData(CurrGrhIndex).pixelWidth, GrhData(CurrGrhIndex).pixelHeight, GrhData(CurrGrhIndex).SX, GrhData(CurrGrhIndex).SY, Light1, FileNum, Angle
                        
                    End If
                    
                End If
            End If
        End If
    End If

End Sub

Private Sub Engine_Render_D3DXSprite(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal SrcX As Single, ByVal SrcY As Single, ByVal Light As Long, ByVal TextureNum As Long, ByVal Degrees As Single)
'*****************************************************************
'Renders a Grh in the form of a D3DXSprite instead of a rectangle (slower, less flexibility)
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Render_D3DXSprite
'*****************************************************************
Dim SrcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2

    'End the old sprite we had going (only if the texture changed)
    If TextureNum <> LastTexture Then
        If SpriteBegun = 1 Then
            Sprite.End
            Sprite.Begin
        End If
    End If
    
    'Ready the texture
    Engine_ReadyTexture TextureNum
    
    'Create the source rectangle
    With SrcRect
        .Left = SrcX
        .Top = SrcY
        .Right = .Left + Width
        .bottom = .Top + Height
    End With
    
    'Create the rotation point
    If Degrees Then
        Degrees = ((Degrees + 180) * DegreeToRadian)
        If Degrees > 360 Then Degrees = Degrees - 360
        With v2
            .X = (Width * 0.5)
            .Y = (Height * 0.5)
        End With
    End If
    
    'Set the translation (location on the screen)
    v3.X = X
    v3.Y = Y

    'Draw the sprite
    If TextureNum > 0 Then
        Sprite.Draw SurfaceDB(TextureNum), SrcRect, SpriteScaleVector, v2, Degrees, v3, Light
    Else
        Sprite.Draw Nothing, SrcRect, SpriteScaleVector, v2, 0, v3, Light
    End If
    
End Sub

Private Sub Engine_Render_ChatBubble(ByVal Text As String, ByVal X As Integer, ByVal Y As Integer)
'*****************************************************************
'Renders a chat bubble and the text for the given text and co-ordinates
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Render_ChatBubble
'*****************************************************************
Const BubbleSectionSize As Long = 6 'The width/height of each "sector" of the bubble in the graphic file
Const RenderColor As Long = -1761607681
Dim TempGrh As Grh
Dim BubbleWidth As Long
Dim BubbleHeight As Long
Dim TempSplit() As String
Dim i As Long
Dim j As Long

    If DisableChatBubbles Then Exit Sub

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
        BubbleWidth = Engine_GetTextWidth(Font_Default, Text)
        BubbleHeight = Font_Default.CharHeight
        
    End If
    
    'Round the width and height to the nearest BubbleSectionSize (the size of each chat bubble side section)
    BubbleWidth = BubbleWidth + BubbleSectionSize
    If BubbleWidth Mod BubbleSectionSize Then BubbleWidth = BubbleWidth + (BubbleSectionSize - (BubbleWidth Mod BubbleSectionSize))
    If BubbleHeight Mod BubbleSectionSize Then BubbleHeight = BubbleHeight + (BubbleSectionSize - (BubbleHeight Mod BubbleSectionSize))

    'Modify the X and Y values the center the bubble
    X = X - (BubbleWidth * 0.5) + 16    'Center
    Y = Y - BubbleHeight - 20           'Align above the head

    '*** Draw the bubble ***
    'Top-left corner
    TempGrh.GrhIndex = 109
    Engine_Render_Grh TempGrh, X, Y, 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    
    'Top-right corner
    TempGrh.GrhIndex = 111
    Engine_Render_Grh TempGrh, X + BubbleWidth + BubbleSectionSize, Y, 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    
    'Bottom-left corner
    TempGrh.GrhIndex = 115
    Engine_Render_Grh TempGrh, X, Y + BubbleHeight + BubbleSectionSize, 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    
    'Bottom-right corner
    TempGrh.GrhIndex = 117
    Engine_Render_Grh TempGrh, X + BubbleWidth + BubbleSectionSize, Y + BubbleHeight + BubbleSectionSize, 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    
    'Top side
    TempGrh.GrhIndex = 110
    For i = 0 To (BubbleWidth \ BubbleSectionSize) - 1
        Engine_Render_Grh TempGrh, X + ((i + 1) * BubbleSectionSize), Y, 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    Next i
    
    'Left side
    TempGrh.GrhIndex = 112
    For i = 0 To (BubbleHeight \ BubbleSectionSize) - 1
        Engine_Render_Grh TempGrh, X, Y + ((i + 1) * BubbleSectionSize), 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    Next i
    
    'Right side
    TempGrh.GrhIndex = 114
    For i = 0 To (BubbleHeight \ BubbleSectionSize) - 1
        Engine_Render_Grh TempGrh, X + BubbleWidth + BubbleSectionSize, Y + ((i + 1) * BubbleSectionSize), 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    Next i

    'Bottom side
    TempGrh.GrhIndex = 116
    For i = 0 To (BubbleWidth \ BubbleSectionSize) - 1
        Engine_Render_Grh TempGrh, X + ((i + 1) * BubbleSectionSize), Y + BubbleHeight + BubbleSectionSize, 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
    Next i

    'Middle
    TempGrh.GrhIndex = 113
    For i = 1 To (BubbleWidth \ BubbleSectionSize)
        For j = 1 To (BubbleHeight \ BubbleSectionSize)
            Engine_Render_Grh TempGrh, X + (i * BubbleSectionSize), Y + (j * BubbleSectionSize), 0, 0, False, RenderColor, RenderColor, RenderColor, RenderColor
        Next j
    Next i

    'Render the text (finally!)
    Engine_Render_Text Font_Default, Text, X + BubbleSectionSize, Y + BubbleSectionSize, D3DColorARGB(255, 0, 0, 0)

End Sub

Private Sub Engine_Render_GUI()
'*****************************************************************
'Render the GUI
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Render_GUI
'*****************************************************************
Dim TempGrh As Grh
Dim i As Long

    'Render the rest of the windows
    For i = NumGameWindows To 1 Step -1
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
    If Not Windowed Then
        TempGrh.FrameCounter = 1
        TempGrh.GrhIndex = 69
        Engine_Render_Grh TempGrh, MousePos.X, MousePos.Y, 0, 0, False
    End If
    
    'Draw item description
    Engine_Render_ItemDesc

End Sub

Private Sub Engine_Render_GUI_Window(ByVal WindowIndex As Byte)
'*****************************************************************
'Render a GUI window
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Render_GUI_Window
'*****************************************************************
Dim User1RenderColor As Long
Dim User2RenderColor As Long
Dim TempGrh As Grh
Dim TempGrh2 As Grh
Dim t As String
Dim s() As String
Dim i As Byte
Dim j As Long

    TempGrh.FrameCounter = 1
    TempGrh2.FrameCounter = 1

    Select Case WindowIndex
        Case TradeWindow
            With GameWindow.Trade
                Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue, GUIColorValue
                
                If TradeTable.User1Accepted Then User1RenderColor = D3DColorARGB(255, 0, 255, 0) Else User1RenderColor = D3DColorARGB(255, 255, 255, 255)
                If TradeTable.User2Accepted Then User2RenderColor = D3DColorARGB(255, 0, 255, 0) Else User2RenderColor = D3DColorARGB(255, 255, 255, 255)

                Engine_Render_Text Font_Default, TradeTable.User1Name, .Screen.X + .User1Name.X, .Screen.Y + .User1Name.Y, User1RenderColor
                Engine_Render_Text Font_Default, TradeTable.User2Name, .Screen.X + .User2Name.X, .Screen.Y + .User2Name.Y, User2RenderColor
                
                For j = 1 To 9
                
                    TempGrh.GrhIndex = TradeTable.Trade1(j).Grh
                    TempGrh2.GrhIndex = TradeTable.Trade2(j).Grh
                
                    Engine_Render_Grh TempGrh, .Screen.X + .Trade1(j).X, .Screen.Y + .Trade1(j).Y, 0, 0, False, User1RenderColor, User1RenderColor, User1RenderColor, User1RenderColor
                    Engine_Render_Grh TempGrh2, .Screen.X + .Trade2(j).X, .Screen.Y + .Trade2(j).Y, 0, 0, False, User2RenderColor, User2RenderColor, User2RenderColor, User2RenderColor
                    
                    Engine_Render_Text Font_Default, TradeTable.Gold1, .Screen.X + .Gold1.X, .Screen.Y + .Gold1.Y, User1RenderColor
                    Engine_Render_Text Font_Default, TradeTable.Gold2, .Screen.X + .Gold2.X, .Screen.Y + .Gold2.Y, User2RenderColor
                
                Next j
                        
            End With
        
        Case NPCChatWindow
            With GameWindow.NPCChat
                Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue, GUIColorValue
                Engine_Render_Text Font_Default, ActiveAsk.QuestionTxt, .Screen.X + 5, .Screen.Y + 5, D3DColorARGB(255, 255, 255, 255)
                For i = 1 To .NumAnswers
                    Engine_Render_Text Font_Default, i & ". " & NPCChat(ActiveAsk.ChatIndex).Ask.Ask(ActiveAsk.AskIndex).Answer(i).Text, .Screen.X + .Answer(i).X, .Screen.Y + .Answer(i).Y, D3DColorARGB(255, 0, 255, 0)
                Next i
            End With
        
        Case StatWindow
            With GameWindow.StatWindow
                Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue, GUIColorValue
                Engine_Render_Text Font_Default, "Str: " & BaseStats(SID.Str) & " + " & ModStats(SID.Str) - BaseStats(SID.Str) & " (" & ModStats(SID.Str) & ")", .Screen.X + .Str.X, .Screen.Y + .Str.Y, -1
                Engine_Render_Text Font_Default, "Agi: " & BaseStats(SID.Agi) & " + " & ModStats(SID.Agi) - BaseStats(SID.Agi) & " (" & ModStats(SID.Agi) & ")", .Screen.X + .Agi.X, .Screen.Y + .Agi.Y, -1
                Engine_Render_Text Font_Default, "Mag: " & BaseStats(SID.Mag) & " + " & ModStats(SID.Mag) - BaseStats(SID.Mag) & " (" & ModStats(SID.Mag) & ")", .Screen.X + .Mag.X, .Screen.Y + .Mag.Y, -1
                If BaseStats(SID.Points) > 0 Then
                    Engine_Render_Grh .AddGrh, .Screen.X + .AddStr.X, .Screen.Y + .AddStr.Y, 0, 1
                    Engine_Render_Grh .AddGrh, .Screen.X + .AddAgi.X, .Screen.Y + .AddAgi.Y, 0, 1
                    Engine_Render_Grh .AddGrh, .Screen.X + .AddMag.X, .Screen.Y + .AddMag.Y, 0, 1
                End If
                Engine_Render_Text Font_Default, "Points: " & BaseStats(SID.Points), .Screen.X + .Points.X, .Screen.Y + .Points.Y, -1
                Engine_Render_Text Font_Default, "Gold: " & BaseStats(SID.Gold), .Screen.X + .Gold.X, .Screen.Y + .Gold.Y, -1
                Engine_Render_Text Font_Default, "Def: " & BaseStats(SID.DEF) & " + " & ModStats(SID.DEF) - BaseStats(SID.DEF) & " (" & ModStats(SID.DEF) & ")", .Screen.X + .DEF.X, .Screen.Y + .DEF.Y, -1
                Engine_Render_Text Font_Default, "Dmg: " & BaseStats(SID.MinHIT) & "+" & ModStats(SID.MinHIT) - BaseStats(SID.MinHIT) & " ~ " & BaseStats(SID.MaxHIT) & "+" & ModStats(SID.MaxHIT) - BaseStats(SID.MaxHIT) & " (" & ModStats(SID.MinHIT) & " ~ " & ModStats(SID.MaxHIT) & ")", .Screen.X + .Dmg.X, .Screen.Y + .Dmg.Y, -1
            End With
        
         Case ChatWindow
            With GameWindow.ChatWindow
                Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue
            End With
            
            'Render the chat text
            Engine_Render_ChatTextBuffer
            
            'Draw entered text
            If EnterText = True Then
                If EnterTextBufferWidth = 0 Then EnterTextBufferWidth = 1   'Dividing by 0 is never good
                If LenB(ShownText) <> 0 Then Engine_Render_Text Font_Default, ShownText, GameWindow.ChatWindow.Screen.X + GameWindow.ChatWindow.Text.X, GameWindow.ChatWindow.Screen.Y + GameWindow.ChatWindow.Text.Y, FontColor_Talk
                If timeGetTime Mod CursorFlashRate * 2 < CursorFlashRate Then
                    TempGrh.GrhIndex = 39
                    TempGrh.FrameCounter = 1
                    TempGrh.Started = 1
                    Engine_Render_Grh TempGrh, GameWindow.ChatWindow.Screen.X + GameWindow.ChatWindow.Text.X + Engine_GetTextWidth(Font_Default, ShownText), GameWindow.ChatWindow.Screen.Y + GameWindow.ChatWindow.Text.Y, 0, 0, False
                End If
            End If
            
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
                Engine_Render_Text Font_Default, MailboxListBuffer, .Screen.X + .List.X, .Screen.Y + .List.Y, -1
                Engine_Render_Text Font_Default, "Read", .Screen.X + .ReadLbl.X, .Screen.Y + .ReadLbl.Y, -1
                Engine_Render_Text Font_Default, "Write", .Screen.X + .WriteLbl.X, .Screen.Y + .WriteLbl.Y, -1
                Engine_Render_Text Font_Default, "Delete", .Screen.X + .DeleteLbl.X, .Screen.Y + .DeleteLbl.Y, -1
                If SelMessage > 0 Then Engine_Render_Rectangle .Screen.X + .List.X, .Screen.Y + .List.Y + ((SelMessage - 1) * Font_Default.CharHeight), .List.Width, Font_Default.CharHeight, 1, 1, 1, 1, 1, 1, 0, 0, 2097217280, 2097217280, 2097217280, 2097217280, , False   'ARGB: 125/0/255/0
            End With
    
        Case ViewMessageWindow
            With GameWindow.ViewMessage
                Engine_Render_Grh .SkinGrh, .Screen.X, .Screen.Y, 0, 1, True, GUIColorValue, GUIColorValue, GUIColorValue, GUIColorValue
                Engine_Render_Text Font_Default, ReadMailData.WriterName, .Screen.X + .From.X, .Screen.Y + .From.Y, -1
                Engine_Render_Text Font_Default, ReadMailData.Subject, .Screen.X + .Subject.X, .Screen.Y + .Subject.Y, -1
                Engine_Render_Text Font_Default, ReadMailData.Message, .Screen.X + .Message.X, .Screen.Y + .Message.Y, -1
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
                If LenB(WriteMailData.RecieverName) <> 0 Then Engine_Render_Text Font_Default, WriteMailData.RecieverName, .Screen.X + .From.X, .Screen.Y + .From.Y, -1
                If WMSelCon = wmFrom Then
                    If timeGetTime Mod CursorFlashRate * 2 < CursorFlashRate Then
                        TempGrh.GrhIndex = 39
                        Engine_Render_Grh TempGrh, .Screen.X + .From.X + Engine_GetTextWidth(Font_Default, WriteMailData.RecieverName), .Screen.Y + .From.Y, 0, 0, False
                    End If
                End If
                'Subject text box
                If LenB(WriteMailData.Subject) <> 0 Then Engine_Render_Text Font_Default, WriteMailData.Subject, .Screen.X + .Subject.X, .Screen.Y + .Subject.Y, -1
                If WMSelCon = wmSubject Then
                    If timeGetTime Mod CursorFlashRate * 2 < CursorFlashRate Then
                        TempGrh.GrhIndex = 39
                        Engine_Render_Grh TempGrh, .Screen.X + .Subject.X + Engine_GetTextWidth(Font_Default, WriteMailData.Subject), .Screen.Y + .Subject.Y, 0, 0, False
                    End If
                End If
                'Message body text box
                t = Engine_WordWrap(WriteMailData.Message, GameWindow.WriteMessage.Message.Width)
                If LenB(WriteMailData.Message) <> 0 Then Engine_Render_Text Font_Default, t, .Screen.X + .Message.X, .Screen.Y + .Message.Y, -1
                If WMSelCon = wmMessage Then
                    If timeGetTime Mod CursorFlashRate * 2 < CursorFlashRate Then
                        If InStr(1, t, vbNewLine) Then
                            s = Split(t, vbNewLine)
                            i = UBound(s)
                            j = Engine_GetTextWidth(Font_Default, s(i))
                        Else
                            i = 0   'Ubound
                            j = Engine_GetTextWidth(Font_Default, t) 'Size
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
                If LenB(AmountWindowValue) <> 0 Then Engine_Render_Text Font_Default, AmountWindowValue, .Screen.X + .Value.X, .Screen.Y + .Value.Y, -1
            End With

    End Select

End Sub

Public Sub Engine_Render_Inventory(Optional ByVal InventoryType As Long = 1)
'*****************************************************************
'Renders the inventory
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Render_Inventory
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
                'If UserInventory(LoopC).Amount > 1 Then
                    Engine_Render_Text Font_Default, UserInventory(LoopC).Amount, DestX, DestY, -1
                'End If
                If UserInventory(LoopC).Equipped Then Engine_Render_Text Font_Default, "E", DestX + (30 - Engine_GetTextWidth(Font_Default, "E")), DestY, -16711936
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
                If UserBank(LoopC).Amount <> -1 Then Engine_Render_Text Font_Default, UserBank(LoopC).Amount, DestX, DestY, -1
            End If
        Next LoopC
    End Select

End Sub

Private Sub Engine_Render_ItemDesc()
'************************************************************
'Draw description text
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Render_ItemDesc
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
    If X + ItemDescWidth > ScreenWidth Then X = ScreenWidth - ItemDescWidth
    If Y < 0 Then Y = 0
    If Y + (ItemDescLines * Font_Default.CharHeight) > ScreenHeight Then Y = ScreenHeight - (ItemDescLines * Font_Default.CharHeight)

    'Draw backdrop
    Engine_Render_Rectangle X - 5, Y - 5, ItemDescWidth + 10, (Font_Default.CharHeight * ItemDescLines) + 10, 1, 1, 1, 1, 1, 1, 0, 0, -1761607681, -1761607681, -1761607681, -1761607681, , False

    'Draw text
    For i = 1 To ItemDescLines
        Engine_Render_Text Font_Default, ItemDescLine(i), X, Y + ((i - 1) * Font_Default.CharHeight), -16777216
    Next i

End Sub

Private Sub Engine_ReadyTexture(ByVal TextureNum As Long)
'************************************************************
'Gets a texture ready to for usage
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_ReadyTexture
'************************************************************

    'Load the surface into memory if it is not in memory and reset the timer
    If TextureNum > 0 Then
        If SurfaceTimer(TextureNum) < timeGetTime Then Engine_Init_Texture TextureNum
        SurfaceTimer(TextureNum) = timeGetTime + SurfaceTimerMax
    End If
    
    'Check what render method we're using
    If AlternateRender Then
    
        'Set the texture
        LastTexture = TextureNum
        If TextureNum <= 0 Then D3DDevice.SetTexture 0, Nothing
        
    Else
    
        'Set the texture
        If TextureNum <= 0 Then
            D3DDevice.SetTexture 0, Nothing
            LastTexture = 0
        Else
            If LastTexture <> TextureNum Then
                D3DDevice.SetTexture 0, SurfaceDB(TextureNum)
                LastTexture = TextureNum
            End If
        End If
        
    End If

End Sub

Sub Engine_Render_Rectangle(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal SrcX As Single, ByVal SrcY As Single, ByVal SrcWidth As Single, ByVal SrcHeight As Single, Optional ByVal SrcBitmapWidth As Long = -1, Optional ByVal SrcBitmapHeight As Long = -1, Optional ByVal Degrees As Single = 0, Optional ByVal TextureNum As Long, Optional ByVal Color0 As Long = -1, Optional ByVal Color1 As Long = -1, Optional ByVal Color2 As Long = -1, Optional ByVal Color3 As Long = -1, Optional ByVal Shadow As Byte = 0, Optional ByVal InBoundsCheck As Boolean = True)
'************************************************************
'Render a square/rectangle based on the specified values then rotate it if needed
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Render_Rectangle
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
Dim L As Single

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
        VertexArray(0).tU = (SrcX / SrcBitmapWidth)
        VertexArray(0).tV = (SrcY / SrcBitmapHeight)
        
        VertexArray(1).X = VertexArray(0).X + Width
        VertexArray(1).tU = ((SrcX + Width) / SrcBitmapWidth)

        VertexArray(2).X = X
        VertexArray(2).tU = (SrcX / SrcBitmapWidth)

        VertexArray(3).X = X + Width
        VertexArray(3).tU = (SrcX + SrcWidth + ShadowAdd) / SrcBitmapWidth

    Else
        
        '------------------------------------------------------------------------------------------------------
        '------------------------------------------------------------------------------------------------------
        'If the image is partially outside of the screen, it is trimmed so only that which is in the screen is drawn
        'This provides for quite a decent FPS boost if you have lots of tiles that stretch outside of the view area
        'Important: Something about this doesn't seem to be functioning correctly. It is supposed to crop down the
        'image and only draw that which is going to be in the screen, but it doesn't work right and I have no
        'idea why. Uncomment the lines to see what happens. I have given up on this since the FPS boost really isn't
        'significant for me to put any more work into it, but if someone could fix it, it would definitely be
        'added back into the engine.
        '------------------------------------------------------------------------------------------------------
        '------------------------------------------------------------------------------------------------------
        'If X < 0 Then
        '    SrcX = SrcX - X
        '    SrcWidth = SrcWidth + X
        '    Width = Width + X
        '    X = 0
        'End If
        'If Y < 0 Then
        '    SrcY = SrcY - Y
        '    SrcHeight = SrcHeight + Y
        '    Height = Height + Y
        '    Y = 0
        'End If
        'If X + Width > ScreenWidth Then
        '    L = X + Width - ScreenWidth
        '    Width = Width - L
        '    SrcWidth = SrcWidth - L
        'End If
        'If Y + Height > ScreenHeight Then
        '    L = Y + Height - ScreenHeight
        '    Height = Height - L
        '    SrcHeight = SrcHeight - L
        'End If
        '------------------------------------------------------------------------------------------------------
        '------------------------------------------------------------------------------------------------------
        
        'If we are NOT using shadows, then we add +1 to the width/height (trust me, just do it)
        ShadowAdd = 1

        'Find the left side of the rectangle
        VertexArray(0).X = X
        VertexArray(0).tU = (SrcX / SrcBitmapWidth)

        'Find the top side of the rectangle
        VertexArray(0).Y = Y
        VertexArray(0).tV = (SrcY / SrcBitmapHeight)
    
        'Find the right side of the rectangle
        VertexArray(1).X = X + Width
        VertexArray(1).tU = (SrcX + SrcWidth + ShadowAdd) / SrcBitmapWidth

        'These values will only equal each other when not a shadow
        VertexArray(2).X = VertexArray(0).X
        VertexArray(3).X = VertexArray(1).X
        
    End If
    
    'Find the bottom of the rectangle
    VertexArray(2).Y = Y + Height
    VertexArray(2).tV = (SrcY + SrcHeight + ShadowAdd) / SrcBitmapHeight

    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(1).Y = VertexArray(0).Y
    VertexArray(1).tV = VertexArray(0).tV
    VertexArray(2).tU = VertexArray(0).tU
    VertexArray(3).Y = VertexArray(2).Y
    VertexArray(3).tU = VertexArray(1).tU
    VertexArray(3).tV = VertexArray(2).tV
    
    'Check if a rotation is required
    If Degrees <> 0 And Degrees <> 360 Then

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

Public Sub Engine_CreateTileLayers()
'************************************************************
'Creates the tile layers used for rendering the tiles so they can be drawn faster
'Has to happen every time the user warps or moves a whole tile
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_CreateTileLayers
'************************************************************
Dim Layer As Byte
Dim ScreenX As Long
Dim ScreenY As Long
Dim tBuf As Integer
Dim pX As Integer
Dim pY As Integer
Dim X As Long
Dim Y As Long
    
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
        
End Sub

Sub Engine_Render_Screen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'************************************************************
'Draw current visible to scratch area based on TileX and TileY
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Render_Screen
'************************************************************
Dim FrameUseMotionBlur As Boolean   'Lets us know if this frame is using motion blur so we don't have to leave support for it on
Dim LightOffset As Long
Dim ChrID() As Integer
Dim ChrY() As Integer
Dim Y As Long           'Keeps track of where on map we are
Dim X As Long
Dim j As Long
Dim Angle As Single
Dim Layer As Byte

    'Check for valid positions
    If UserPos.X = 0 Then Exit Sub
    If UserPos.Y = 0 Then Exit Sub
    If UserCharIndex = 0 Then Exit Sub
    
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
    ParticleOffsetX = (Engine_PixelPosX(ScreenMinX) - PixelOffsetX)
    ParticleOffsetY = (Engine_PixelPosY(ScreenMinY) - PixelOffsetY)

    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then
        
        'The worst we can do at this point is avoid an error we can't fix!
        On Error Resume Next
        
        'Do a loop while device is lost
        If D3DDevice.TestCooperativeLevel = D3DERR_DEVICELOST Then Exit Sub
            
        'Clear all the textures
        LastTexture = -999
        For j = 1 To NumGrhFiles
            Set SurfaceDB(j) = Nothing
            SurfaceTimer(j) = 0
            SurfaceSize(j).X = 0
            SurfaceSize(j).Y = 0
        Next j
        
        'Clear the D3DXSprite
        If AlternateRenderDefault = 1 Or AlternateRenderMap = 1 Or AlternateRenderText = 1 Then
            SpriteBegun = 0
            Set Sprite = Nothing
            Set Sprite = D3DX.CreateSprite(D3DDevice)
        End If
        
        Set DeviceBuffer = Nothing
        Set DeviceStencil = Nothing
        Set BlurStencil = Nothing
        Set BlurTexture = Nothing
        Set BlurSurf = Nothing
        
        'Make sure the scene is ended
        D3DDevice.EndScene
        
        'Reset the device
        D3DDevice.Reset D3DWindow
        
        Set DeviceBuffer = D3DDevice.GetRenderTarget
        Set DeviceStencil = D3DDevice.GetDepthStencilSurface
        Set BlurStencil = D3DDevice.CreateDepthStencilSurface(BufferWidth, BufferHeight, D3DFMT_D16, D3DMULTISAMPLE_NONE)
        Set BlurTexture = D3DX.CreateTexture(D3DDevice, BufferWidth, BufferHeight, 0, D3DUSAGE_RENDERTARGET, DispMode.Format, D3DPOOL_DEFAULT)
        Set BlurSurf = BlurTexture.GetSurfaceLevel(0)
        
        'Reset the render states
        Engine_Init_RenderStates
        
        'Load the particle textures
        Engine_Init_ParticleEngine True
        
        On Error GoTo 0

    Else
    
        'We have to bypass the present the first time through here or else we get an error
        If NotFirstRender Then
        
            'Close off the last sprite
            If SpriteBegun Then
                Sprite.End
                SpriteBegun = 0
                LastTexture = -101
            End If

            With D3DDevice
                
                'End the rendering (scene)
                .EndScene
                
                'Flip the backbuffer to the screen
                .Present ByVal 0, ByVal 0, 0, ByVal 0
                
            End With
                
        Else
        
            'Set NotFirstRender to True so we can start displaying
            NotFirstRender = True
            
        End If
    
    End If
    
    'Check if running (turn on motion blur)
    If UseMotionBlur Then
        If UserCharIndex > 0 Then
            If CharList(UserCharIndex).Moving = 1 And CharList(UserCharIndex).Running Then
                BlurIntensity = 45
            Else
                If BlurIntensity < 255 Then
                    BlurIntensity = BlurIntensity + (ElapsedTime * 0.8)
                    If BlurIntensity > 255 Then BlurIntensity = 255
                End If
            End If
        End If
    End If
    
    'Set the motion blur if needed
    If UseMotionBlur Then
        If BlurIntensity < 255 Or ZoomLevel > 0 Then
            FrameUseMotionBlur = True
            D3DDevice.SetRenderTarget BlurSurf, BlurStencil, 0
        End If
    End If

    'Begin the scene
    D3DDevice.BeginScene
    
    'Clear the screen with a solid color (to prevent artifacts)
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    
    '************** Layer 1 to 3 **************

    'Set the alternate rendering for the map on / off
    AlternateRender = AlternateRenderMap
    
    'Loop through the lower 3 layers
    For Layer = 1 To 3
        LightOffset = ((Layer - 1) * 4) + 1
        
        'Loop through all the tiles we know we will draw for this layer
        For j = 1 To TileLayer(Layer).NumTiles
            With TileLayer(Layer).Tile(j)
                
                'Check if we have to draw with a shadow or not (slighty changes because we have to animate on the shadow, not the main render)
                If MapData(.TileX, .TileY).Shadow(Layer) = 1 Then
                    Engine_Render_Grh MapData(.TileX, .TileY).Graphic(Layer), .PixelPosX + PixelOffsetX, .PixelPosY + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
                    Engine_Render_Grh MapData(.TileX, .TileY).Graphic(Layer), .PixelPosX + PixelOffsetX, .PixelPosY + PixelOffsetY, 0, 0, True, MapData(.TileX, .TileY).Light(LightOffset), MapData(.TileX, .TileY).Light(LightOffset + 1), MapData(.TileX, .TileY).Light(LightOffset + 2), MapData(.TileX, .TileY).Light(LightOffset + 3)
                Else
                    Engine_Render_Grh MapData(.TileX, .TileY).Graphic(Layer), .PixelPosX + PixelOffsetX, .PixelPosY + PixelOffsetY, 0, 1, True, MapData(.TileX, .TileY).Light(LightOffset), MapData(.TileX, .TileY).Light(LightOffset + 1), MapData(.TileX, .TileY).Light(LightOffset + 2), MapData(.TileX, .TileY).Light(LightOffset + 3)
                End If
                
            End With
        Next j
        
    Next Layer
    
    'Set the alternate rendering back to what it was before
    AlternateRender = AlternateRenderDefault

    '************** Objects **************
    For j = 1 To LastObj
        If OBJList(j).Grh.GrhIndex Then
            X = Engine_PixelPosX(OBJList(j).Pos.X - minX) + PixelOffsetX + OBJList(j).Offset.X + TileBufferOffset
            Y = Engine_PixelPosY(OBJList(j).Pos.Y - minY) + PixelOffsetY + OBJList(j).Offset.Y + TileBufferOffset
            If Y >= -32 Then
                If Y <= (ScreenHeight + 32) Then
                    If X >= -32 Then
                        If X <= (ScreenWidth + 32) Then
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
            X = Engine_PixelPosX(CharList(ChrID(j)).Pos.X - minX) + PixelOffsetX + TileBufferOffset
            Y = Engine_PixelPosY(CharList(ChrID(j)).Pos.Y - minY) + PixelOffsetY + TileBufferOffset
            
            If Y >= -32 And Y <= (ScreenHeight + 32) And X >= -32 And X <= (ScreenWidth + 32) Then
                
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

    '************** Layer 4 to 6 **************
    AlternateRender = AlternateRenderMap
    For Layer = 4 To 6
        LightOffset = ((Layer - 1) * 4) + 1
        For j = 1 To TileLayer(Layer).NumTiles
            With TileLayer(Layer).Tile(j)
                If MapData(.TileX, .TileY).Shadow(Layer) = 1 Then
                    Engine_Render_Grh MapData(.TileX, .TileY).Graphic(Layer), .PixelPosX + PixelOffsetX, .PixelPosY + PixelOffsetY, 0, 1, True, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
                    Engine_Render_Grh MapData(.TileX, .TileY).Graphic(Layer), .PixelPosX + PixelOffsetX, .PixelPosY + PixelOffsetY, 0, 0, True, MapData(.TileX, .TileY).Light(LightOffset), MapData(.TileX, .TileY).Light(LightOffset + 1), MapData(.TileX, .TileY).Light(LightOffset + 2), MapData(.TileX, .TileY).Light(LightOffset + 3)
                Else
                    Engine_Render_Grh MapData(.TileX, .TileY).Graphic(Layer), .PixelPosX + PixelOffsetX, .PixelPosY + PixelOffsetY, 0, 1, True, MapData(.TileX, .TileY).Light(LightOffset), MapData(.TileX, .TileY).Light(LightOffset + 1), MapData(.TileX, .TileY).Light(LightOffset + 2), MapData(.TileX, .TileY).Light(LightOffset + 3)
                End If
            End With
        Next j
    Next Layer
    AlternateRender = AlternateRenderDefault
    
    '************** Effects **************
    
    'Loop to do drawing
    If LastEffect > 0 Then
        For j = 1 To LastEffect
            If EffectList(j).Grh.GrhIndex Then
                X = Engine_PixelPosX(EffectList(j).Pos.X - minX) + PixelOffsetX + TileBufferOffset
                Y = Engine_PixelPosY(EffectList(j).Pos.Y - minY) + PixelOffsetY + TileBufferOffset
                If EffectList(j).Time <> 0 And EffectList(j).Time < timeGetTime Then
                
                    'Timer ran out
                    Engine_Effect_Erase j
                    
                ElseIf Y >= -32 And Y <= (ScreenHeight + 32) And X >= -32 And X <= (ScreenWidth + 32) Then
                
                    'Timer or animation going, render
                    Engine_Render_Grh EffectList(j).Grh, X, Y, 1, 1, 0, , , , , , EffectList(j).Angle
                    
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
                ProjectileList(j).X = ProjectileList(j).X + (Sin(Angle) * ElapsedTime * 0.63)
                ProjectileList(j).Y = ProjectileList(j).Y - (Cos(Angle) * ElapsedTime * 0.63)
                
                'Update the rotation
                If ProjectileList(j).RotateSpeed > 0 Then
                    ProjectileList(j).Rotate = ProjectileList(j).Rotate + (ProjectileList(j).RotateSpeed * ElapsedTime * 0.01)
                    Do While ProjectileList(j).Rotate > 360
                        ProjectileList(j).Rotate = ProjectileList(j).Rotate - 360
                    Loop
                End If

                'Draw if within range
                X = ((-minX - 1) * 32) + ProjectileList(j).X + PixelOffsetX + TileBufferOffset
                Y = ((-minY - 1) * 32) + ProjectileList(j).Y + PixelOffsetY + TileBufferOffset
                If Y >= -32 Then
                    If Y <= (ScreenHeight + 32) Then
                        If X >= -32 Then
                            If X <= (ScreenWidth + 32) Then
                                If ProjectileList(j).Rotate = 0 Then
                                    Engine_Render_Grh ProjectileList(j).Grh, X, Y, 0, 1, 1, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1
                                    Engine_Render_Grh ProjectileList(j).Grh, X, Y, 0, 0, 1
                                Else
                                    Engine_Render_Grh ProjectileList(j).Grh, X, Y, 0, 0, 1, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1, ProjectileList(j).Rotate
                                    Engine_Render_Grh ProjectileList(j).Grh, X, Y, 0, 1, 1, , , , , , ProjectileList(j).Rotate
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
            X = Engine_PixelPosX(BloodList(j).Pos.X - minX) + PixelOffsetX + TileBufferOffset
            Y = Engine_PixelPosY(BloodList(j).Pos.Y - minY) + PixelOffsetY + TileBufferOffset
            If Y >= -32 Then
                If Y <= (ScreenHeight + 32) Then
                    If X >= -32 Then
                        If X <= (ScreenWidth + 32) Then
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
    
    'Do the general weather updating
    Engine_Weather_Update

    '************** Chat bubbles **************
    'Loop through the chars
    For j = 1 To LastChar
        If CharList(j).Active Then
            If LenB(CharList(j).BubbleStr) <> 0 Then
                If CharList(j).RealPos.X > -25 Then
                    If CharList(j).RealPos.X < ScreenWidth + 25 Then
                        If CharList(j).RealPos.Y > -25 Then
                            If CharList(j).RealPos.Y < ScreenHeight + 25 Then
                                Engine_Render_ChatBubble CharList(j).BubbleStr, CharList(j).RealPos.X, CharList(j).RealPos.Y
                                CharList(j).BubbleTime = CharList(j).BubbleTime - ElapsedTime
                                If CharList(j).BubbleTime <= 0 Then
                                    CharList(j).BubbleTime = 0
                                    CharList(j).BubbleStr = vbNullString
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
            X = (((DamageList(j).Pos.X - minX) - 1) * TilePixelWidth) + PixelOffsetX + TileBufferOffset
            Y = (((DamageList(j).Pos.Y - minY) - 1) * TilePixelHeight) + PixelOffsetY + TileBufferOffset
            If Y >= -32 Then
                If Y <= (ScreenHeight + 32) Then
                    If X >= -32 Then
                        If X <= (ScreenWidth + 32) Then
                            Engine_Render_Text Font_Default, DamageList(j).Value, X, Y, D3DColorARGB(255, 255, 0, 0)
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
    
    '************** Mini-map **************
    Const tS As Single = 3  'Size of the mini-map dots
    
    'Check if the mini-map is being shownquit
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
                If CharList(X).Active Then
                
                    'The user's character
                    If X = UserCharIndex Then
                        j = D3DColorARGB(200, 0, 255, 0)    'User's character
                        Engine_Render_Rectangle CharList(X).Pos.X * tS, CharList(X).Pos.Y * tS, tS, tS, 1, 1, 1, 1, 1, 1, 0, 0, j, j, j, j, , False
                        GoTo NextChar
                    End If
                    
                    'Part of the user's group or one of the user's slaves
                    If CharList(X).CharType = ClientCharType_Grouped Or (CharList(X).CharType = ClientCharType_Slave And UserCharIndex = CharList(X).OwnerChar) Then
                        If X <> UserCharIndex Then
                            j = D3DColorARGB(200, 100, 220, 100)    'PC (grouped) or the user's slave
                            Engine_Render_Rectangle CharList(X).Pos.X * tS, CharList(X).Pos.Y * tS, tS, tS, 1, 1, 1, 1, 1, 1, 0, 0, j, j, j, j, , False
                            GoTo NextChar
                        End If
                    End If
                
                    'Check if the character is in screen, since the only characters drawn outside of the screen are grouped characters
                    If CharList(X).Pos.X > ScreenMinX Then
                        If CharList(X).Pos.X < ScreenMaxX Then
                            If CharList(X).Pos.Y > ScreenMinY Then
                                If CharList(X).Pos.Y < ScreenMaxY Then
                                    
                                    'Character is a PC
                                    If CharList(X).CharType = ClientCharType_PC Then
                                        j = D3DColorARGB(200, 0, 255, 255)  'PC (not grouped)
                                    'Character is a NPC
                                    Else
                                        j = D3DColorARGB(200, 0, 150, 150)  'NPC
                                    End If
                                    
                                    'Any character but one part of the user's group
                                    Engine_Render_Rectangle CharList(X).Pos.X * tS, CharList(X).Pos.Y * tS, tS, tS, 1, 1, 1, 1, 1, 1, 0, 0, j, j, j, j, , False

                                End If
                            End If
                        End If
                    End If
                    
                End If
                
NextChar:
                
            Next X
            
        End If
        
    End If

    'Show FPS
    Engine_Render_Text Font_Default, "FPS: " & FPS, ScreenWidth - 80, 2, -1

    'Check if using motion blur / zooming
    If UseMotionBlur Then
        If FrameUseMotionBlur Then
            With D3DDevice
            
                'Perform the zooming calculations
                ' * 1.333... maintains the aspect ratio
                ' ... / 1024 is to factor in the buffer size
                BlurTA(0).tU = ZoomLevel * 1.333333333
                BlurTA(0).tV = ZoomLevel
                BlurTA(1).tU = ((ScreenWidth + 1) / 1024) - (ZoomLevel * 1.333333333)
                BlurTA(1).tV = ZoomLevel
                BlurTA(2).tU = ZoomLevel * 1.333333333
                BlurTA(2).tV = ((ScreenHeight + 1) / 1024) - ZoomLevel
                BlurTA(3).tU = BlurTA(1).tU
                BlurTA(3).tV = BlurTA(2).tV
                
                'Draw what we have drawn thus far since the last .Clear
                LastTexture = -100
                .SetRenderTarget DeviceBuffer, DeviceStencil, 0
                .SetTexture 0, BlurTexture
                .SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(BlurIntensity, 255, 255, 255)
                .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR
                .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BlurTA(0), FVF_Size
                .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
            
            End With
        End If
    End If

End Sub

Public Sub Engine_BuildMiniMap()
'************************************************************
'Builds the array for the minimap. Theres multiple styles available, but only one
'is used in the demo, so experiment with them and check which one you like!
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_BuildMiniMap
'************************************************************
Dim NumMiniMapTiles As Long         'UBound of the MiniMapTile array
Dim MiniMapTile() As MiniMapTile    'Color of each tile and their position
Dim MMC_Blocked As Long
Dim MMC_Exit As Long
Dim MMC_Sign As Long
Dim Offset As Long
Dim tVA() As TLVERTEX
Dim X As Long
Dim Y As Long
Dim j As Long

    'Change to the type of map you want
    Const UseOption As Byte = 2
    
    'The size of the tiles
    'If you change this value, you must also change the "tS" constant in Engine_Render_Screen to the same value!
    Const MiniMapSize As Single = 3

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

Private Function Engine_NPCChat_MeetsConditions(ByVal NPCIndex As Integer, ByVal LineIndex As Byte, Optional ByVal SayLine As String = vbNullString) As Byte
'*****************************************************************
'Checks if the conditions have been satisfied for a chat line
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_NPCChat_MeetsConditions
'*****************************************************************
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
        If LenB(SayLine) <> 0 Then   'If the string is not empty
            SayLine = " " & UCase$(SayLine) & " "   'We compair it in UCase$(), since case doesn't matter
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
                    If LenB(SayLine) = 0 Then Exit Function     'No chance it can be right if theres no text!
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
'*****************************************************************
'Checks for a NPC chat triggers
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_NPCChat_CheckForChatTriggers
'*****************************************************************
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
                    If Engine_RectDistance(.RealPos.X, .RealPos.Y, .RealPos.X - ((ScreenWidth - 50) \ 2), .RealPos.Y - ((ScreenHeight - 50) \ 2), ((ScreenWidth - 50) \ 2) + 1, ((ScreenHeight - 50) \ 2) + 1) Then
                    
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
'*****************************************************************
'Updates the automated NPC chatting
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_NPCChat_Update
'*****************************************************************
Dim i As Byte

    'We're going to be using this object a hell of a lot...
    With CharList(CharIndex)
        
        'Make sure the NPC has automated chat
        If .NPCChatIndex > 0 Then
            
            'Check for a valid distance
            If Engine_RectDistance(.RealPos.X, .RealPos.Y, .RealPos.X - ((ScreenWidth - 50) \ 2), .RealPos.Y - ((ScreenHeight - 50) \ 2), ((ScreenWidth - 50) \ 2) + 1, ((ScreenHeight - 50) \ 2) + 1) Then
            
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
'*****************************************************************
'Adds the NPCChat text according to the style
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_NPCChat_AddText
'*****************************************************************
    
    With CharList(CharIndex)

        'Check for text before adding it
        If LenB(NPCChat(.NPCChatIndex).ChatLine(.NPCChatLine).Text) <> 0 Then
    
            'Find out the style used, and add the chat according to the style
            Select Case NPCChat(.NPCChatIndex).ChatLine(.NPCChatLine).Style
                Case NPCCHAT_STYLE_BUBBLE
                    Engine_MakeChatBubble CharIndex, Engine_WordWrap(.Name & ": " & NPCChat(.NPCChatIndex).ChatLine(.NPCChatLine).Text, BubbleMaxWidth)
                Case NPCCHAT_STYLE_BOX
                    Engine_AddToChatTextBuffer .Name & ": " & NPCChat(.NPCChatIndex).ChatLine(.NPCChatLine).Text, FontColor_Talk
                Case NPCCHAT_STYLE_BOTH
                    Engine_MakeChatBubble CharIndex, Engine_WordWrap(.Name & ": " & NPCChat(.NPCChatIndex).ChatLine(.NPCChatLine).Text, BubbleMaxWidth)
                    Engine_AddToChatTextBuffer .Name & ": " & NPCChat(.NPCChatIndex).ChatLine(.NPCChatLine).Text, FontColor_Talk
            End Select
            
        End If
            
        'Add the chat delay (we do the delay even if theres no text)
        .NPCChatDelay = NPCChat(.NPCChatIndex).ChatLine(.NPCChatLine).Delay
        
    End With

End Sub

Private Function Engine_NPCChat_NextLine(ByVal CharIndex As Integer, Optional ByVal ChatTxt As String) As Byte
'*****************************************************************
'Gets the next free line to use for the NPC chat (0 if none found)
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_NPCChat_NextLine
'*****************************************************************
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
'*****************************************************************
'Check if the path is clear from the user to the target of blocked tiles
'For the line-rect collision, we pretend that each tile is 2 units wide
'so we can give them a width of 1 to center things
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_ClearPath
'*****************************************************************
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
'*****************************************************************
'Render the spells list
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Render_Skills
'*****************************************************************
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

Public Sub Engine_Render_Text(ByRef UseFont As CustomFont, ByVal Text As String, ByVal X As Long, ByVal Y As Long, ByVal Color As Long)
'*****************************************************************
'Render text with a custom font
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Render_Text
'*****************************************************************
Dim TempVA(0 To 3) As TLVERTEX
Dim TempStr() As String
Dim Count As Integer
Dim Ascii() As Byte
Dim Row As Integer
Dim u As Single
Dim v As Single
Dim i As Long
Dim j As Long
Dim KeyPhrase As Byte
Dim TempColor As Long
Dim ResetColor As Byte
Dim SrcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim YOffset As Single
    
    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
    
    'Assign the alternate rendering value
    AlternateRender = AlternateRenderText

    'Get the text into arrays (split by vbCrLf)
    TempStr = Split(Text, vbCrLf)
    
    'Set the temp color (or else the first character has no color)
    TempColor = Color
    
    'Check for alternate rendering
    If AlternateRender Then

        'End the old sprite we had going
        If SpriteBegun = 1 Then
            Sprite.End
            Sprite.Begin
        End If
        
    Else
        
        'Set the texture
        D3DDevice.SetTexture 0, UseFont.Texture

    End If
    
    'Clear the LastTexture, letting the rest of the engine know that the texture needs to be changed for next rect render
    LastTexture = -(Rnd * 10000)
    
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(TempStr)
        If Len(TempStr(i)) > 0 Then
            YOffset = i * UseFont.CharHeight
            Count = 0
        
            'Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(i), vbFromUnicode)
        
            'Loop through the characters
            For j = 1 To Len(TempStr(i))

                'Check for a key phrase
                If Ascii(j - 1) = 124 Then 'If Ascii = "|"
                    KeyPhrase = (Not KeyPhrase)  'TempColor = ARGB 255/255/0/0
                    If KeyPhrase Then TempColor = -65536 Else ResetColor = 1
                Else

                    'Render with triangles
                    If AlternateRender = 0 Then

                        'Copy from the cached vertex array to the temp vertex array
                        CopyMemory TempVA(0), UseFont.HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), FVF_Size * 4

                        'Set up the verticies
                        TempVA(0).X = X + Count
                        TempVA(0).Y = Y + YOffset
                        
                        TempVA(1).X = TempVA(1).X + X + Count
                        TempVA(1).Y = TempVA(0).Y

                        TempVA(2).X = TempVA(0).X
                        TempVA(2).Y = TempVA(2).Y + TempVA(0).Y

                        TempVA(3).X = TempVA(1).X
                        TempVA(3).Y = TempVA(2).Y
                        
                        'Set the colors
                        TempVA(0).Color = TempColor
                        TempVA(1).Color = TempColor
                        TempVA(2).Color = TempColor
                        TempVA(3).Color = TempColor

                        'Draw the verticies
                        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TempVA(0), FVF_Size
                        
                    'Render with D3DXSprite
                    Else
                    
                        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
                        Row = (Ascii(j - 1) - UseFont.HeaderInfo.BaseCharOffset) \ UseFont.RowPitch
                        u = ((Ascii(j - 1) - UseFont.HeaderInfo.BaseCharOffset) - (Row * UseFont.RowPitch)) * UseFont.ColFactor
                        v = Row * UseFont.RowFactor

                        'Create the source rectangle
                        With SrcRect
                            .Left = u * UseFont.TextureSize.X
                            .Top = v * UseFont.TextureSize.Y
                            .Right = .Left + (UseFont.ColFactor * UseFont.TextureSize.X)
                            .bottom = .Top + (UseFont.RowFactor * UseFont.TextureSize.Y)
                        End With
                        
                        'Set the translation (location on the screen)
                        v3.X = X + Count
                        v3.Y = Y + (UseFont.CharHeight * i)
                    
                        'Draw the sprite
                        Sprite.Draw UseFont.Texture, SrcRect, SpriteScaleVector, v2, 0, v3, Color
  
                    End If
  
                    'Shift over the the position to render the next character
                    Count = Count + UseFont.HeaderInfo.CharWidth(Ascii(j - 1))
                
                End If
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = Color
                End If
                
            Next j
            
        End If
    Next i
    
    'Retreive the default alternate render value
    AlternateRender = AlternateRenderDefault

End Sub

Public Sub Engine_SetItemDesc(ByVal Name As String, Optional ByVal Amount As Integer = 0, Optional ByVal Value As Long = 0)
'*****************************************************************
'Set item description values
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_SetItemDesc
'*****************************************************************
Dim i As Byte
Dim X As Long

    'Set the item values
    ItemDescLine(1) = Name
    ItemDescLines = 1
    If Amount <> 0 Then
        ItemDescLines = ItemDescLines + 1
        ItemDescLine(ItemDescLines) = "Amount: " & Amount
    End If
    If Value <> 0 Then
        ItemDescLines = ItemDescLines + 1
        ItemDescLine(ItemDescLines) = "Value: " & Value
    End If

    'Get the largest size
    ItemDescWidth = Engine_GetTextWidth(Font_Default, ItemDescLine(1))
    If ItemDescLines > 1 Then
        For i = 2 To ItemDescLines
            X = Engine_GetTextWidth(Font_Default, ItemDescLine(i))
            If X > ItemDescWidth Then ItemDescWidth = X
        Next i
    End If

End Sub

Sub Engine_ShowNextFrame()
'*****************************************************************
'Updates and draws next frame to screen
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_ShowNextFrame
'*****************************************************************

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
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_SkillIDtoGRHID
'*****************************************************************

    Select Case SkillID
        Case SkID.Bless: Engine_SkillIDtoGRHID = 46
        Case SkID.IronSkin: Engine_SkillIDtoGRHID = 47
        Case SkID.Strengthen: Engine_SkillIDtoGRHID = 48
        Case SkID.Warcry: Engine_SkillIDtoGRHID = 49
        Case SkID.Protection: Engine_SkillIDtoGRHID = 50
        Case SkID.SpikeField: Engine_SkillIDtoGRHID = 62
        Case SkID.Heal: Engine_SkillIDtoGRHID = 63
        Case SkID.SummonBandit: Engine_SkillIDtoGRHID = 1
    End Select

End Function

Public Function Engine_SkillIDtoSkillName(ByVal SkillID As Byte) As String
'*****************************************************************
'Takes in a SkillID and returns the name of that skill
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_SkillIDtoSkillName
'*****************************************************************

    Select Case SkillID
        Case SkID.Bless: Engine_SkillIDtoSkillName = "Bless"
        Case SkID.IronSkin: Engine_SkillIDtoSkillName = "Iron Skin"
        Case SkID.Strengthen: Engine_SkillIDtoSkillName = "Strengthen"
        Case SkID.Warcry: Engine_SkillIDtoSkillName = "War Cry"
        Case SkID.Protection: Engine_SkillIDtoSkillName = "Protection"
        Case SkID.SpikeField: Engine_SkillIDtoSkillName = "Spike Field"
        Case SkID.Heal: Engine_SkillIDtoSkillName = "Heal"
        Case SkID.SummonBandit: Engine_SkillIDtoSkillName = "Summon Bandit"
        Case Else: Engine_SkillIDtoSkillName = "Unknown Skill"
    End Select

End Function

Public Sub Engine_SortIntArray(TheArray() As Integer, TheIndex() As Integer, ByVal LowerBound As Integer, ByVal UpperBound As Integer)
'*****************************************************************
'Sort an array of integers
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_SortIntArray
'*****************************************************************
Dim indxt As Long   'Stored index
Dim swp As Integer  'Swap variable
Dim i As Integer    'Subarray Low  Scan Index
Dim j As Integer    'Subarray High Scan Index

    'Start the loop
    For j = LowerBound + 1 To UpperBound
        indxt = TheIndex(j)
        swp = TheArray(indxt)
        For i = j - 1 To LowerBound Step -1
            If TheArray(TheIndex(i)) <= swp Then Exit For
            TheIndex(i + 1) = TheIndex(i)
        Next i
        TheIndex(i + 1) = indxt
    Next j

End Sub

Sub Engine_UnloadAllForms()
'*****************************************************************
'Unloads all forms
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_UnloadAllForms
'*****************************************************************
Dim frm As Form

    For Each frm In VB.Forms
        Unload frm
    Next

End Sub

Function Engine_Distance(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer) As Single
'*****************************************************************
'Finds the distance between two points
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Distance
'*****************************************************************

    Engine_Distance = Sqr(((Y1 - Y2) ^ 2 + (x1 - x2) ^ 2))
    
End Function

Sub Engine_UseQuickBar(ByVal Slot As Byte)
'*****************************************************************
'Use the object in the quickbar slot
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_UseQuickBar
'*****************************************************************

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
                    sndBuf.Allocate 5
                    sndBuf.Put_Byte DataCode.User_CastSkill
                    sndBuf.Put_Byte QuickBarID(Slot).ID
                    sndBuf.Put_Integer TargetCharIndex
                    sndBuf.Put_Byte CharList(UserCharIndex).Heading
                End If
            End If
        End If

    End Select

End Sub

Public Function Engine_GetBlinkTime() As Long
'*****************************************************************
'Return a value on how long until the next blink happens
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetBlinkTime
'*****************************************************************

    Engine_GetBlinkTime = 4000 + Int(Rnd * 5000)
    
End Function

Public Function Engine_RectDistance(ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long, ByVal MaxXDist As Long, ByVal MaxYDist As Long) As Byte
'*****************************************************************
'Check if two tile points are in the same area
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_RectDistance
'*****************************************************************

    If Abs(x1 - x2) < MaxXDist + 1 Then
        If Abs(Y1 - Y2) < MaxYDist + 1 Then
            Engine_RectDistance = True
        End If
    End If

End Function

Public Function Engine_FindDirection(Pos As Position, Target As Position) As Byte
'*****************************************************************
'Returns the direction in which the Target is from the Pos, 0 if equal
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_FindDirection
'*****************************************************************
Dim pX As Integer
Dim pY As Integer
Dim tX As Integer
Dim tY As Integer
Dim X As Integer
Dim Y As Integer

    'Put the bytes into integer variables (causes overflows for negatives, even if the return is an integer)
    pX = Pos.X
    pY = Pos.Y
    tX = Target.X
    tY = Target.Y
    
    'Find the difference
    X = pX - tX
    Y = pY - tY
    Log "Server_FindDirection: Position difference (X:" & X & " Y:" & Y & ") found", CodeTracker '//\\LOGLINE//\\

    'Same position
    If X = 0 Then
        If Y = 0 Then
            Engine_FindDirection = NORTH
            Exit Function
        End If
    End If

    'NE
    If X <= -1 Then
        If Y >= 1 Then
            Engine_FindDirection = NORTHEAST
            Exit Function
        End If
    End If

    'NW
    If X >= 1 Then
        If Y >= 1 Then
            Engine_FindDirection = NORTHWEST
            Exit Function
        End If
    End If

    'SW
    If X >= 1 Then
        If Y <= -1 Then
            Engine_FindDirection = SOUTHWEST
            Exit Function
        End If
    End If

    'SE
    If X <= -1 Then
        If Y <= -1 Then
            Engine_FindDirection = SOUTHEAST
            Exit Function
        End If
    End If

    'South
    If Y <= -1 Then
        Engine_FindDirection = SOUTH
        Exit Function
    End If

    'North
    If Y >= 1 Then
        Engine_FindDirection = NORTH
        Exit Function
    End If

    'West
    If X >= 1 Then
        Engine_FindDirection = WEST
        Exit Function
    End If

    'East
    If X <= -1 Then
        Engine_FindDirection = EAST
        Exit Function
    End If

End Function
