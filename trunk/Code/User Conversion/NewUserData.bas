Attribute VB_Name = "NewUserData"
'*******************************************************************************
'Place in this module the new variables (so that way it can save correctly)
'*******************************************************************************
Option Explicit
Public Const MaxQuests As Byte = 20         'How many quests a user can accept at once
Public Const NumSkills As Byte = 8
Public Const NumStats As Byte = 31
Public Const MAX_INVENTORY_SLOTS = 49   'Maximum number of slots
Public Const MaxMailPerUser As Byte = 50    'How much mail each user may have maximum
Public Type QuestStatus 'Status of user's current quests
    NPCKills As Integer     'How many of the targeted NPCs the user has killed
End Type
Type UserOBJ    'Objects the user has
    ObjIndex As Long    'Index of the object
    Amount As Long      'Amount of the objects
    Equipped As Byte    'If the object is equipted
End Type
Type WorldPos   'Holds placement information
    Map As Integer  'Map
    X As Integer    'X coordinate
    Y As Integer    'Y coordinate
End Type
Type UserCounters   'Counters for a user
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
Type Char   'Holds data for a user or NPC character
    CharIndex As Integer        'Character's index
    Hair As Integer             'Hair index
    Head As Integer             'Head index
    Body As Integer             'Body index
    Weapon As Integer           'Weapon index
    Wings As Integer            'Wings index
    Heading As Byte             'Current direction facing
    HeadHeading As Byte         'Direction char's head is facing
    Desc As String              'Description
End Type
Type UserFlags  'Flags for a user
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
    Disconnecting As Byte   'If the user will be disconnected after data is sent
    QuestNPC As Integer     'The ID of the NPC that the user is talking to about a quest
End Type
Type Skills 'User skills casted
    IronSkin As Byte
    Bless As Integer
    Protect As Integer
    Strengthen As Integer
    WarCurse As Integer
End Type
Type User   'Holds data for a user
    Name As String      'Name of the user
    Password As String  'User's password
    Char As Char        'Defines users looks
    Desc As String      'User's description
    Pos As WorldPos     'User's current position
    Gold As Long        'How much gold the user has
    IP As String            'User Ip
    ConnID As Long          'Connection ID
    SendBuffer() As Byte    'Buffer for sending data
    Object(1 To MAX_INVENTORY_SLOTS) As NewUserData.UserOBJ 'The user's inventory
    WeaponEqpObjIndex As Integer    'The index of the equipted weapon
    WeaponEqpSlot As Byte           'Slot of the equipted weapon
    WeaponType As Byte              'Type of weapon the user is using
    ArmorEqpObjIndex As Integer     'The index of the equipted armor
    ArmorEqpSlot As Byte            'Slot of the equipted armor
    Counters As UserCounters    'Declares the user counters
    BaseStats(1 To NumStats) As Long    '--IMPORTANT - THIS REPLACES THE USERSTATS CLASS MODULE!--
    Flags As UserFlags          'Declares the user flags
    Skills As Skills            'Declares the skills casted on the user
    KnownSkills(1 To NumSkills) As Byte 'Declares the skills known by the user
    CompletedQuests As String   'The string contains the indexes of all completed quests in order
    Quest(1 To MaxQuests) As Integer            'The quest index of the current quests if any
    QuestStatus(1 To MaxQuests) As QuestStatus  'Counts certain parts of quests that require being counted (ie NPC kills)
    MailID(1 To MaxMailPerUser) As Integer      'ID of the user's mail
    MailboxPos As WorldPos                      'Position of the last-used mailbox
End Type
Public UserList() As User   'Holds data for each user

