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
'************            Official Release: Version 0.1.2            ************
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
'**        http://www.vbgore.com/en/index.php?title=Donate                    **
'**  *Contribute - Check out our forums, contribute ideas, report bugs, or    **
'**        help expend the wiki pages!                                        **
'**  *Link To Us - Creating a link to vbGORE, whether it is on your own web   **
'**        page or a link to vbGORE in a forum you visit, every link helps    **
'**        spread the word of vbGORE's existance! Buttons and banners for     **
'**        linking to vbGORE can be found on the following page:              **
'**        http://www.vbgore.com/en/index.php?title=Buttons_and_Banners       **
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
'**                                                                           **
'** Also, all the members of the vbGORE community who have submitted          **
'** tutorials, bugs, suggestions, criticism and have just stuck around!!      **
'**                                                                           **
'** If you feel you belong in these credits, please contact Spodi (above).    **
'*******************************************************************************
'*******************************************************************************

Option Explicit

Public Const NumStats As Byte = 31
Public Const NumSkills As Byte = 8

Public Const DegreeToRadian As Single = 0.0174532925

'Position displayed in the form's caption
Public HoverX As Long
Public HoverY As Long
Public HovertX As Long
Public HovertY As Long

'Which layer is selected in frmTile
Public SelectedLayer As Integer

'States if the project is unloading (has to give Sox time to unload)
Public IsUnloading As Byte

'Control
Public prgRun As Boolean 'When true the program ends

'Tells us which box we are going to set the value we get from SetTile to
Public stBoxID As Byte

'********** Map display variables **********
Public WeatherChkValue As Byte
Public CharsChkValue As Byte
Public ObjChkValue As Byte
Public BrightChkValue As Byte
Public GridChkValue As Byte
Public InfoChkValue As Byte

Public SetTilesChkValue As Byte
Public ViewTilesChkValue As Byte
Public ShowMapInfoChkValue As Byte
Public ShowNPCsChkValue As Byte
Public PartChkValue As Byte
Public FloodsChkValue As Byte
Public ObjEditChkValue As Byte
Public ExitsChkValue As Byte
Public BlocksChkValue As Byte
Public SfxChkValue As Byte

Public L1ChkValue As Byte
Public L2ChkValue As Byte
Public L3ChkValue As Byte
Public L4ChkValue As Byte
Public L5ChkValue As Byte
Public L6ChkValue As Byte

'********** Map optimization variables **********
Public Enum MapOptType
    None = 0
    ObjOnBlocked = 1
    NPCOnBlocked = 2
    DuplicateGrhLayers = 3
    EmptyLight = 4
End Enum
Public Type MapOpt
    Type As MapOptType
    tX As Byte
    tY As Byte
    Layer As Byte
    Layer2 As Byte
End Type
Public MapOpt() As MapOpt

'********** Object types **********
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

'********** NPC Types **********
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
    Name As String          'Name of the NPC
    Char As CharShort       'Defines NPC looks
    Desc As String          'Description

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
    VendItems() As OBJ              'Information on the item the NPC is vending
End Type

'********** OUTSIDE FUNCTIONS ***********
'For Get and Write Var
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:36)  Decl: 285  Code: 0  Total: 285 Lines
':) CommentOnly: 72 (25.3%)  Commented: 45 (15.8%)  Empty: 18 (6.3%)  Max Logic Depth: 1
