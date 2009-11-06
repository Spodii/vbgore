Attribute VB_Name = "Declares"
Public Const ServerId As Integer = 0


'Holds a position on a 2d grid
Public Type Position
    X As Long
    Y As Long
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

'Holds a position on a 2d grid in floating variables (singles)
Public Type FloatPos
    X As Single
    Y As Single
End Type

'**************************************************************
'** Below is where you add some info on the new types you add **
'**************************************************************

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

'Holds info about a object
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
    DropC As Integer
End Type
Public DropObjs() As Obj
Public ShopObjs() As Obj

'Holds data for a character - used for saving/loading NPCs (not displaying the char)
Type CharShort
    CharIndex As Integer    'Character's index
    Body As Integer         'Body index
    Weapon As Integer       'Weapon index
    Pants As Integer
    Armor As Integer        'Armor index
    Shoe As Integer
    Eye As Integer         'eye index
    Helm As Integer
    Shield As Integer
    Hair As Integer         'Hair index
    Heading As Byte         'Current direction facing
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

'Simplified npc type
Type NPC
    Name As String      'Name of the NPC
    Npcnumber As Integer    'The NPC index within NPC.dat
End Type

'**************************************************************
'** This is where you add some info on the new types you add **
'**************************************************************
'Somewhat simplified char info
Public Type Char
    Active As Byte
    Heading As Byte
    HeadHeading As Byte
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
    ScrollDirectionX As Integer
    ScrollDirectionY As Integer
    ActionIndex As Byte
End Type

'Very simplified Object info
Public Type ObjData
    Name As String              'Name
    GrhIndex As Integer         'Graphic index
End Type
Public ObjData() As ObjData

'Running speed - make sure you have the same value on the server!
Public Const RunningSpeed As Byte = 5

'The NPC we are editing
Public OpenNPC As NPC
Public Npcnumber As Integer
Public CharList(1) As Char
