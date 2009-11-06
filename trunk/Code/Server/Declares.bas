Attribute VB_Name = "Declares"
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
'************            Official Release: Version 0.3.1            ************
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
'**        http://www.vbgore.com/index.php?title=Donate                       **
'**  *Contribute - Check out our forums, contribute ideas, report bugs, or    **
'**        help expend the wiki pages!                                        **
'**  *Link To Us - Creating a link to vbGORE, whether it is on your own web   **
'**        page or a link to vbGORE in a forum you visit, every link helps    **
'**        spread the word of vbGORE's existance! Buttons and banners for     **
'**        linking to vbGORE can be found on the following page:              **
'**        http://www.vbgore.com/index.php?title=Buttons_and_Banners          **
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
'** Chase and Nex666: Help with mapping, graphics, bug reports, hosting, etc  **
'** Graphics (Avatar): Supplied the character sprite graphics, + a few more   **
'**   http://www.zidev.com/                                                   **
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
'**                                                                           **
'** Also, all the members of the vbGORE community who have submitted          **
'** tutorials, bugs, suggestions, criticism and have just stuck around!!      **
'**                                                                           **
'** If you feel you belong in these credits, please contact Spodi (above).    **
'*******************************************************************************
'*******************************************************************************

Option Explicit

'Used to record the number of packets coming in/out and what command ID they have
Public Const DEBUG_UseLogging As Boolean = True                  '//\\LOGLINE//\\
Public Const DEBUG_RecordPacketsOut As Boolean = False

'********** Public CONSTANTS ***********

'Change to 1 to enable database optimization on runtime
Public Const OptimizeDatabase As Byte = 0

'If we run the server in high priority (recommended to use unless on a test server)
Public Const RunHighPriority As Byte = 0

'How long objects can be on the ground (in miliseconds) before being removed
Public Const GroundObjLife As Long = 300000 '5 minutes

'How long the maps last in memory when no users are on it
Public Const EmptyMapLife As Long = 1800    '3 minutes

'Amount of time that must elapse for certain user events (in miliseconds)
Public Const DelayTimeMail As Long = 3000   'Sending messages (has to be updated client-side, too)
Public Const DelayTimeTalk As Long = 500    'Talking (in any form)

'Change this value to add a cost to sending mail
Public Const MailCost As Long = 0

'Maximum amount of objects allowed on a single tile
Public Const MaxObjsPerTile As Byte = 5

'Running information
Public Const RunningSpeed As Byte = 5   'How much to increase speed when running
Public Const RunningCost As Long = 2    'How much stamina it cost to run

'How long char remains aggressive-faced after being attacked
Public Const AGGRESSIVEFACETIME = 4000

'Calculate the data in/out per sec or ont
Public Const CalcTraffic As Boolean = True

'Sound constants
Public Const SOUND_SWING As Byte = 7
Public Const SOUND_WARP As Byte = 1

'Stat constants
Public Const STAT_ATTACKWAIT As Long = 1000 'How many ms a user has to wait till he can attack again

'How many quests a user can accept at once
Public Const MaxQuests As Byte = 20

'Time that must elapse for NPC to make another action (in miliseconds) after attacking
Public Const NPCDelayFight As Long = 1000

'Blocked directions - take the blocked byte and OR these values (If Blocked OR <Direction> Then...)
Public Const BlockedNorth As Byte = 1
Public Const BlockedEast As Byte = 2
Public Const BlockedSouth As Byte = 4
Public Const BlockedWest As Byte = 8
Public Const BlockedAll As Byte = 15

'Character types for CharList()
Public Const CharType_PC As Byte = 1
Public Const CharType_NPC As Byte = 2

'Max distance for two chars being on the same screen (for the rect distance)
Public Const MaxServerDistanceX As Long = 13
Public Const MaxServerDistanceY As Long = 10

'************ Positioning ************
Type WorldPos   'Holds placement information
    Map As Integer  'Map
    X As Byte       'X coordinate
    Y As Byte       'Y coordinate
End Type

'************ Object types ************
Public Const MAX_INVENTORY_OBJS As Integer = 9999   'Maximum number of objects per slot (same obj)
Public Const MAX_INVENTORY_SLOTS As Byte = 49       'Maximum number of slots
Public Type ObjData
    name As String                  'Name
    ObjType As Byte                 'Type (armor, weapon, item, etc)
    GrhIndex As Long                'Graphic index
    SpriteBody As Integer           'Index of the body sprite to change to
    SpriteWeapon As Integer         'Index of the weapon sprite to change to
    SpriteHair As Integer           'Index of the hair sprite to change to
    SpriteHead As Integer           'Index of the head sprite to change to
    SpriteWings As Integer          'Index of the wings sprite to change to
    WeaponType As Byte              'What type of weapon, if it is a weapon
    WeaponRange As Byte             'Range of the weapon (only applicable if a ranged WeaponType)
    UseGrh As Long                  'Grh of the object when used (projectile for ranged, slash for melee, effects for use-once)
    ProjectileRotateSpeed As Byte   'How fast the projectile rotates (if at all)
    Price As Long                   'Price of the object
    RepHP As Long                   'How much HP to replenish
    RepMP As Long                   'How much MP to replenish
    RepSP As Long                   'How much SP to replenish
    RepHPP As Integer               'Percentage of HP to replenish
    RepMPP As Integer               'Percentage of MP to replenish
    RepSPP As Integer               'Percentage of SP to replenish
    AddStat(1 To NumStats) As Long  'How much to add to the stat by the SID
End Type
Public ObjData() As ObjData
Public Type Obj 'Holds info about a object
    ObjIndex As Integer     'Index of the object
    Amount As Integer       'Amount of the object
End Type

'************ Map Tiles/Information ************
Type NPCLoadData    'Used to load NPCs from the .temp map files
    NPCNum As Integer
    X As Byte
    Y As Byte
End Type
Type MapBlock   'Information for each map block
    '*** IMPORTANT! *** ADDING ARRAYS TO THIS UDT WILL BREAK THE LOADING!
    'If you must add an array, try adding it to a different UDT, like I did with the objects
    TileExitMap As Integer      'Warp location when user touches the tile
    TileExitX As Byte
    TileExitY As Byte
    Blocked As Byte             'If the tile is blocked
    Mailbox As Byte             'If there is a mailbox on the tile
    UserIndex As Integer        'Index of the user on the tile
    NPCIndex As Integer         'Index of the NPC on the tile
End Type
Type ObjBlock   'Information on an object on a map block
    NumObjs As Byte             'Number of objects on the tile
    ObjLife() As Long           'When the object was created (used to determine it's life)
    ObjInfo() As Obj            'Information of the object on the tile
End Type
Type MapInfo    'Map information
    NumUsers As Integer     'Number of users on the map
    name As String          'Name of the map
    MapVersion As Integer   'Version of the map
    Weather As Byte         'What weather effects the map has going
    Music As Byte           'The music file number of the map
    DataLoaded As Byte      'If the map data is loaded
    UnloadTimer As Long     'How long until the surface unloads
    Data() As MapBlock      'Holds the information on each tile; Data(TileX, TileY)
    ObjTile() As ObjBlock   'Holds the information on the objects on the tiles; Obj(TileX, TileY)
End Type
Public MapInfo() As MapInfo

'************ Mailing System ************
Public Const MaxMailPerUser As Byte = 50    'How much mail each user may have maximum
Public Const MaxMailObjs As Byte = 10       'How many objects can be attached to a message maximum
Type MailData   'Mailing system information
    Subject As String
    WriterName As String
    RecieveDate As Date
    Message As String
    New As Byte
    Obj(1 To MaxMailObjs) As Obj
End Type

'************ Generic Character Data ************
Type CharData   'Charlist types (for reverting from CharIndex to PC/NPC index)
    Index As Integer
    CharType As Byte    '0 = Unused, 1 = PC, 2 = NPC
End Type
Public CharList() As CharData

'************ Quest ************
Public Type Quest
    name As String                  'The quest's name
    StartTxt As String              'What the NPC says to the player to explain the crisis
    AcceptTxt As String             'What the NPC says when the player accepts the quest
    IncompleteTxt As String         'What the NPC says to the player when they return without completing the quest
    FinishTxt As String             'What the NPC says when the player finishes the quest
    AcceptReqLvl As Long            'What level the user is required to be to accept
    AcceptReqObj As Integer         'Index of the object the user is required to have to accept
    AcceptReqObjAmount As Integer   'How much of the object the user must have before accepting
    AcceptRewExp As Long            'Amount of Exp the user gets for accepting the quest
    AcceptRewGold As Long           'Amount of gold the user gets for accepting the quest
    AcceptRewObj As Integer         'Object the user gets for accepting the quest
    AcceptRewObjAmount As Integer   'Amount of the object the user gets for accepting the quest
    AcceptLearnSkill As Byte        'Skill the user learns for accepting the quest (by SkID value)
    FinishReqObj As Integer         'Object the user must have to finish the quest
    FinishReqObjAmount As Integer   'Amount of the object the user must have to finish the quest
    FinishReqNPC As Integer         'Index of the NPC the user must kill to finish the quest
    FinishReqNPCAmount As Integer   'How many of the NPCs the user must kill to finish the quest
    FinishRewExp As Long            'Exp the user gets for finishing the quest
    FinishRewGold As Long           'How much gold the user gets for finishing the quest
    FinishRewObj As Integer         'The index of the object the user gets for finishing the quest
    FinishRewObjAmount As Integer   'The amount of the object the user gets for finishing the quest
    FinishLearnSkill As Byte        'Skill the user learns for finishing the quest (by SkID value)
    Redoable As Byte                'If the quest can be done infinite times
End Type
Public QuestData() As Quest

'************ NPC/Character types ************
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
Public Type QuestStatus 'Status of user's current quests
    NPCKills As Integer     'How many of the targeted NPCs the user has killed
End Type
Type UserFlags  'Flags for a user
    UserLogged As Byte      'If the user is logged in
    SwitchingMaps As Byte   'If the user is switching maps
    LastViewedMail As Byte  'The last mail index which the user viewed
    TradeWithNPC As Integer 'NPC the user is trading with
    TargetIndex As Integer  'Index of the NPC or Player targeted
    Target As Byte          'Type of targeting - 0 for none, 1 for player, 2 for NPC
    GMLevel As Byte         'What type of admin the user is: 0 = None, 1 = GM
    Disconnecting As Byte   'If the user will be disconnected after data is sent
    QuestNPC As Integer     'The ID of the NPC that the user is talking to about a quest
End Type
Type UserCounters   'Counters for a user
    IdleCount As Long           'Stores last time the user sent an action packet
    LastPacket As Long          'Stores last time the user sent ANY packet
    AttackCounter As Long       'Stores last time user attacked
    MoveCounter As Long         'Stores last time the user moved
    AggressiveCounter As Long   'How long the user will remain aggressive-faced
    SpellExhaustion As Long     'Time until another spell can be casted
    BlessCounter As Long        'Time left on bless
    ProtectCounter As Long      'Time left on protection
    StrengthenCounter As Long   'Time left on strengthen
    WarCurseCounter As Long     'Time left on warcry-curse
    DelayTimeMail As Long       'Mail write delay time
    DelayTimeTalk As Long       'Talk delay time
    PacketsInCount As Long      'Packets in per second (used to prevent packet flooding)
    PacketsInTime As Long       'When the packet counting started
End Type
Type UserOBJ    'Objects the user has
    ObjIndex As Long    'Index of the object
    Amount As Long      'Amount of the objects
    Equipped As Byte    'If the object is equipted
End Type
Type Skills 'Skills casted on a user / NPC (the value holds how powerful it is, which is often based off of the magic stat)
    IronSkin As Byte
    Bless As Integer
    Protect As Integer
    Strengthen As Integer
    WarCurse As Integer
End Type
Type KnownSkills    'Known skills by the user
    IronSkin As Byte
    Bless As Byte
    Protect As Byte
    Strengthen As Byte
    Warcry As Byte
    Heal As Byte
    SpikeField As Byte
    Spike As Byte
End Type
Type User   'Holds data for a user
    name As String      'Name of the user
    Char As Char        'Defines users looks
    Desc As String      'User's description
    Pos As WorldPos     'User's current position
    Gold As Long        'How much gold the user has
    ConnID As Long          'Connection ID
    SendBuffer() As Byte    'Buffer for sending data
    BufferSize As Long      'Size of the buffer
    HasBuffer As Byte       'If there is anything in the buffer
    HasMovePacket As Byte   'If there is a move packet in the buffer (if there is, buffer is forced to PP_High upon movement)
    PPValue As Byte         'Packet priority value
    PPCount As Long         'Packet priority count-down (only valid if PPValue = PP_Low)
    PacketWait As Long      'Packet wait count-down (not to be confused with the packet priority - this one is for Packet_WaitTime)
    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ 'The user's inventory
    Bank(1 To MAX_INVENTORY_SLOTS) As Obj       'The user's bank items
    WeaponEqpObjIndex As Integer    'The index of the equipted weapon
    WeaponEqpSlot As Byte           'Slot of the equipted weapon
    WeaponType As Byte              'Type of weapon the user is using
    ArmorEqpObjIndex As Integer     'The index of the equipted armor
    ArmorEqpSlot As Byte            'Slot of the equipted armorn
    WingsEqpObjIndex As Integer     'The index of the equipted Wings
    WingsEqpSlot As Byte            'Slot of the equipted Wings
    Counters As UserCounters    'Declares the user counters
    Stats As UserStats          'Declares the user stats
    flags As UserFlags          'Declares the user flags
    Skills As Skills            'Declares the skills casted on the user
    KnownSkills(1 To NumSkills) As Byte 'Declares the skills known by the user
    CompletedQuests As String   'The string contains the indexes of all completed quests in order
    Quest(1 To MaxQuests) As Integer            'The quest index of the current quests if any
    QuestStatus(1 To MaxQuests) As QuestStatus  'Counts certain parts of quests that require being counted (ie NPC kills)
    MailID(1 To MaxMailPerUser) As Long         'ID of the user's mail
    MailboxPos As WorldPos                      'Position of the last-used mailbox
End Type
Public UserList() As User   'Holds data for each user
Type NPCFlags   'Flags for a NPC
    NPCAlive As Byte        'If the NPC is alive and visible
    NPCActive As Byte       'If the NPC is active
    Thralled As Byte        'If the NPC is thralled (if so, it does not get saved or respawn)
    UpdateStats As Byte     'If to update the mod stats
End Type
Type NPCCounters    'Counters for a NPC
    RespawnCounter As Long      'Stores the death time to respawn later
    AggressiveCounter As Long   'How long the NPC will remain aggressive-faced
    SpellExhaustion As Long     'Time until another spell can be casted
    BlessCounter As Long        'Time left on bless
    ProtectCounter As Long      'Time left on protection
    StrengthenCounter As Long   'Time left on strengthen
    WarCurseCounter As Long     'Time left on warcry-curse
    ActionDelay As Long         'How long until the NPC can perform another action
End Type
Type NPC    'Holds all the NPC variables
    name As String  'Name of the NPC
    Char As Char    'Defines NPC looks
    Desc As String  'Description
    Pos As WorldPos         'Current NPC Postion
    StartPos As WorldPos    'Spawning location of the NPC
    NPCNumber As Integer    'The NPC index within NPC.dat
    AI As Byte              'Used AI algorithm
    ChatID As Byte          'Chat ID used
    RespawnWait As Long     'How long for the NPC to respawn
    Attackable As Byte      'If the NPC is attackable
    Hostile As Byte         'If the NPC is hostile
    GiveEXP As Long         'How much exp given on death
    GiveGLD As Long         'How much gold given on death
    Quest As Integer        'Quest index
    AttackRange As Byte     'The NPC's attack range
    AttackGrh As Long       'Grh used when the NPC attacks
    ProjectileRotateSpeed As Byte   'If a projectile, how fast it rotates
    Skills As Skills                'Declares the skills casted on the NPC
    flags As NPCFlags               'Declares the NPC's flags
    Counters As NPCCounters         'Declares the NPC's counters
    NumVendItems As Byte            'Number of items the NPC is vending
    NumDropItems As Byte            'Number of items the NPC is dropping
    BaseStat(1 To NumStats) As Long 'Declares the NPC's stats
    ModStat(1 To NumStats) As Long  'Declares the NPC's stats
    
    'THESE ARRAYS MUST STAY DOWN HERE AT THE BOTTOM OF THE UDT!
    VendItems() As Obj              'Information on the item the NPC is vending
    DropItems() As Obj              'Information on the item to drop
    DropRate() As Single            'The drop rate of the item in the DropItems() array sharing the same index
End Type
Public NPCList() As NPC     'Holds data for each NPC

'Two bytes put together (used for the NPC loading/saving of vending/drop item amounts)
Public Type NPCBytes
    Vend As Byte
    Drop As Byte
End Type

'***********************************
'********** Misc Values ************
'***********************************
'All the below can be changed without worry of conversion

'Weapon type constants
Public Enum WeaponType
    Hand = 0        'Weapon is hand-based
    Staff = 1       'Weapon is a staff
    Dagger = 2      'Weapon is a dagger
    Sword = 3       'Weapon is a sword
    Throwing = 4    'Weapon is thrown (ninja stars, throwing knives, etc)
End Enum
#If False Then
Private Hand, Staff, Dagger, Sword, Throwing
#End If

'Object types
Public Const OBJTYPE_USEONCE As Byte = 1    'Objects that can be used only once
Public Const OBJTYPE_WEAPON As Byte = 2     'Weapons of all types
Public Const OBJTYPE_ARMOR As Byte = 3      'Body armors
Public Const OBJTYPE_WINGS As Byte = 4      'Wings

'Constants for headings
Public Const NORTH As Byte = 1
Public Const EAST As Byte = 2
Public Const SOUTH As Byte = 3
Public Const WEST As Byte = 4
Public Const NORTHEAST As Byte = 5
Public Const SOUTHEAST As Byte = 6
Public Const SOUTHWEST As Byte = 7
Public Const NORTHWEST As Byte = 8

'Map sizes
Public Const XMaxMapSize As Byte = 100  'Maximum width of the map in tiles
Public Const XMinMapSize As Byte = 1    'Minimum width of the map in tiles
Public Const YMaxMapSize As Byte = 100  'Maximum height of the map in tiles
Public Const YMinMapSize As Byte = 1    'Minimum height of the map in tiles

'Window size in tiles
Public Const XWindow As Byte = 25   'Size of the window's width in tiles
Public Const YWindow As Byte = 18   'Size of the window's height in tiles

'********** Public VARS ***********

'Where the map borders are.. Set during load
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte
Public ResPos As WorldPos
Public StartPos As WorldPos
Public NumUsers As Integer  'Current number of users
Public LastUser As Integer  'Index of the last user
Public LastChar As Integer
Public LastNPC As Integer
Public NumNPCs As Integer
Public NumMaps As Integer
Public NumQuests As Integer
Public NumObjDatas As Integer

Public IdleLimit As Long
Public LastPacket As Long
Public MaxUsers As Integer

'All the users located on a map
Public Type MapUsersType
    Index() As Long
End Type
Public MapUsers() As MapUsersType

'Names of NPCs (for NPCs involved in quests)
Public NPCName() As String

'Number of connections (used just for displaying purposes)
Public CurrConnections As Long

'The time the server started (in system time)
Public ServerStartTime As Long

'States the server is running
Public ServerRunning As Byte

'ID of the local socket
Public LocalSoxID As Long

'Buffer used for conversions to send to Data_Send
Public ConBuf As DataBuffer

'Traffic information (bytes are converted to kbytes to allow larger numbers)
Public DataIn As Long
Public DataOut As Long
Public DataKBIn As Long
Public DataKBOut As Long

'The number of bytes we need to send all of our known skills
Public NumBytesForSkills As Long

'Server is unloading
Public UnloadServer As Byte

'Help variables
Public Const NumHelpLines As Byte = 3
Public HelpLine(1 To NumHelpLines) As String    'These are filled in on frmMain.StartServer

Public DebugPacketsOut() As Long

Public LastTimeGetTime As Long

'********** EXTERNAL FUNCTIONS ***********
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
