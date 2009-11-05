Attribute VB_Name = "Declares"
Option Explicit

'Path of the loaded file
Public FilePath As String

Public Const NumStats As Byte = 31
Public Const NumSkills As Byte = 8
Public Const MAX_INVENTORY_SLOTS = 49
Public Const MaxMailPerUser As Byte = 50

Type WorldPos
    Map As Integer  'Map
    X As Integer    'X coordinate
    Y As Integer    'Y coordinate
End Type

'Holds data for a user or NPC character
Type Char
    CharIndex As Integer    'Character's index
    Hair As Integer            'Hair index
    Head As Integer            'Head index
    Body As Integer            'Body index
    Weapon As Integer          'Weapon index
    Heading As Byte         'Current direction facing
    HeadHeading As Byte     'Direction char's head is facing
    Desc As String          'Description
End Type

'Flags for a user
Type UserFlags
    UserLogged As Byte      'If the user is logged in
    SwitchingMaps As Byte   'If the user is switching maps
    DownloadingMap As Byte  'If the user is downloading a map update
    LastViewedMail As Byte  'The last mail index which the user viewed
    TradeWithNPC As Integer 'NPC the user is trading with
    DevMode As Byte         'If the user is in Dev Mode
    SetTileX As Byte        'Tile which the dev is setting
    SetTileY As Byte
    TargetIndex As Integer  'Index of the NPC or Player targeted
    Target As Byte          'Type of targeting - 0 for none, 1 for player, 2 for NPC
    AdminID As Byte         'What type of admin the user is: 0 = None, 1 = GM, 2 = Dev, 3 = GM/Dev
End Type

Type UserCounters
    IdleCount As Long           'Stores last time the user made a move
    AttackCounter As Long       'Stores last time user attacked
    MoveCounter As Long         'Stores last time the user moved
    SendMapCounter As WorldPos  'Stores map counter information
    BlinkCounter As Long        'How long until the user has to blink automatically
    AggressiveCounter As Long   'How long the user will remain aggressive-faced

    SpellExhaustion As Long     'Time until another spell can be casted
    BlessCounter As Long        'Time left on bless
    ProtectCounter As Long      'Time left on protection
    StrengthenCounter As Long   'Time left on strengthen
    WarCurseCounter As Long     'Time left on warcry-curse
End Type

Type UserOBJ
    ObjIndex As Long    'Index of the object
    Amount As Long      'Amount of the objects
    Equipped As Byte    'If the object is equipted
End Type

'Holds the things you are asked to do to finish a quest, but oly those that are strictly needed
Type QuestRequirements
    TargetNPC As Integer        'The index of the NPC you have to kill
    TargetNPCNumber As Integer  'The number left to be killed
End Type

'User skills casted
Type Skills
    IronSkin As Byte
    Bless As Integer
    Protect As Integer
    Strengthen As Integer
    WarCurse As Integer
End Type

'Known skills
Type KnownSkills
    IronSkin As Byte
    Bless As Byte
    Protect As Byte
    Strengthen As Byte
    Warcry As Byte
    Heal As Byte
    SpikeField As Byte
    Spike As Byte
End Type

'Holds data for a user
Type User
    Name As String      'Name of the user
    Password As String  'User's password
    Char As Char        'Defines users looks
    Desc As String      'User's description
    Pos As WorldPos     'User's current position
    Gold As Long        'How much gold the user has

    IP As String            'User Ip
    ConnID As Long          'Connection ID
    SendBuffer() As Byte    'Buffer for sending data

    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ 'The user's inventory
    WeaponEqpObjIndex As Integer    'The index of the equipted weapon
    WeaponEqpSlot As Byte           'Slot of the equipted weapon
    WeaponType As Byte              'Type of weapon the user is using
    ArmorEqpObjIndex As Integer     'The index of the equipted armor
    ArmorEqpSlot As Byte            'Slot of the equipted armor

    Counters As UserCounters    'Declares the user counters
    Stats As UserStats          'Declares the user stats
    Flags As UserFlags          'Declares the user flags
    Skills As Skills            'Declares the skills casted on the user
    KnownSkills(1 To NumSkills) As Byte 'Declares the skills known by the user

    CompletedQuests As String   'The string contains the indexes of all completed quests in order
    Quest As Integer            'The quest index of the current quest if any
    QuestRequirements As QuestRequirements  'Requirements for current quest
    MailID(1 To MaxMailPerUser) As Integer  'ID of the user's mail
    MailboxPos As WorldPos                  'Position of the last-used mailbox
End Type

Public UserChar As User

Public Declare Function getprivateprofilestring Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Function writeprivateprofilestring Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:44)  Decl: 187  Code: 0  Total: 187 Lines
':) CommentOnly: 62 (33.2%)  Commented: 62 (33.2%)  Empty: 19 (10.2%)  Max Logic Depth: 1
