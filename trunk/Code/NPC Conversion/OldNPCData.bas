Attribute VB_Name = "OldNPCData"
'*******************************************************************************
'Place in this module the old variables (so that way it can load correctly)
'*******************************************************************************
Option Explicit
Public Const NumStats As Byte = 18
Public Type Obj 'Holds info about a object
    ObjIndex As Integer     'Index of the object
    Amount As Integer       'Amount of the object
End Type
Type Skills 'User skills casted
    IronSkin As Byte
    Bless As Integer
    Protect As Integer
    Strengthen As Integer
    WarCurse As Integer
End Type
Type WorldPos   'Holds placement information
    Map As Integer  'Map
    X As Integer    'X coordinate
    Y As Integer    'Y coordinate
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
Type NPCFlags   'Flags for a NPC
    NPCAlive As Byte        'If the NPC is alive and visible
    NPCActive As Byte       'If the NPC is active
    ActionDelay As Long     'How long until the NPC can perform another action
    WalkPath() As WorldPos  'The position the NPC will be traveling
    HasPath As Byte         'If the NPC has a path they are following
    PathPos As Integer      'The index in the WalkPath() the NPC is currently on
    GoalX As Byte           'The position the NPC is trying to get to with the walkpath
    GoalY As Byte
End Type
Type NPCCounters    'Counters for a NPC
    RespawnCounter As Long  'Stores the death time to respawn later
    BlinkCounter As Long    'How long until the NPC blinks again
    AggressiveCounter As Long   'How long the NPC will remain aggressive-faced
    SpellExhaustion As Long     'Time until another spell can be casted
    BlessCounter As Long        'Time left on bless
    ProtectCounter As Long      'Time left on protection
    StrengthenCounter As Long   'Time left on strengthen
    WarCurseCounter As Long     'Time left on warcry-curse
End Type
Type NPC    'Holds all the NPC variables
    Name As String  'Name of the NPC
    Char As Char    'Defines NPC looks
    Desc As String  'Description
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
    NumVendItems As Integer         'Number of items the NPC is vending
    VendItems() As Obj              'Information on the item the NPC is vending
End Type
Public NPCList() As NPC     'Holds data for each NPC
