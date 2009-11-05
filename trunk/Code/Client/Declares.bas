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


'********** Debug/Display Settings ************
'These are your key constants - reccomended you turn off ALL debug constants before
' compiling your code for public usage just speed reasons

'These two are mostly used for checking to make sure the encryption works
Public Const DEBUG_PrintPacket_In As Boolean = False     'Shows packets coming in in chat box
Public Const DEBUG_PrintPacket_Out As Boolean = False    'Shows packets going out in chat box

'********** Object types ************
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

'********** Emoticons ************
Public Const NumEmotes As Byte = 10
Public Type EmoID
    Dots As Byte
    Exclimation As Byte
    Question As Byte
    Surprised As Byte
    Heart As Byte
    Hearts As Byte
    HeartBroken As Byte
    Utensils As Byte
    Meat As Byte
    ExcliQuestion As Byte
End Type
Public EmoID As EmoID

'********** Character Stats/Skills ************
Public Const NumStats As Byte = 31
Public Const NumSkills As Byte = 8
Public BaseStats(1 To NumStats) As Long
Public ModStats(1 To NumStats) As Long
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
    WeaponSkill As Byte
End Type
Public SID As StatOrder 'Stat ID
Public Type SkillID
    Bless As Byte
    Protection As Byte
    Strengthen As Byte
    Warcry As Byte
    Heal As Byte
    IronSkin As Byte
    Curse As Byte
    SpikeField As Byte
End Type
Public SkID As SkillID  'Skill IDs

'Item description variables
Public ItemDescWidth As Long
Public ItemDescLine(1 To 10) As String  'Allow 10 lines maximum
Public ItemDescLines As Byte

'Object constants
Public Const MAX_INVENTORY_SLOTS As Byte = 49

'User's inventory
Type Inventory
    ObjIndex As Long
    Name As String
    GrhIndex As Integer
    Amount As Integer
    Equipped As Boolean
End Type

'Known user skills/spells
Public UserKnowSkill(1 To NumSkills)

'User status vars
Public UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory

'Server stuff
Public SendNewChar As Boolean 'Used during login
Public DownloadingMap As Boolean 'Currently downloading a map from server

'Control
Public prgRun As Boolean 'When true the program ends

'Game Dev - check if tile has changed on map tile
Public MapTileChanged(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As Byte

'Data String Codenames (Reduces all data transfers to 1 byte tags)
Public Type DataCode
    Comm_Talk As Byte                   'Normal chat - "@"
    Comm_UMsgbox As Byte                'Urgent Messagebox - "!!"
    Comm_Shout As Byte                  'Shout A Message - "/SHOUT"
    Comm_Emote As Byte                  'Emote A Message - ":"
    Comm_Whisper As Byte                'Whisper A Private Message - "\"
    Comm_FontType_Talk As Byte
    Comm_FontType_Fight As Byte
    Comm_FontType_Info As Byte
    Comm_FontType_Quest As Byte
    Server_MailMessage As Byte
    Server_MailBox As Byte
    Server_MailItemInfo As Byte
    Server_MailItemTake As Byte
    Server_MailItemRemove As Byte
    Server_MailDelete As Byte
    Server_MailCompose As Byte
    Server_UserCharIndex As Byte        'User Character Index - "SUC"
    Server_SetUserPosition As Byte      'Set User's Position - "SUP"
    Server_MakeChar As Byte             'Create New Character From Map - "MAC"
    Server_EraseChar As Byte            'Erase Character From Map - "ERC"
    Server_MoveChar As Byte             'Move Character On Map - "MOC"
    Server_ChangeChar As Byte           'Change Character Apperance - "CHC"
    Server_MakeObject As Byte           'Create An Object On Map - "MOB"
    Server_EraseObject As Byte          'Erase An Object On Map - "EOB"
    Server_PlaySound As Byte            'Play A Sound On Client - "PLW"
    Server_Who As Byte                  'Who Is Currently Online - "/WHO"
    Server_CharHP As Byte
    Server_CharMP As Byte
    Server_IconCursed As Byte
    Server_IconWarCursed As Byte
    Server_IconBlessed As Byte
    Server_IconStrengthened As Byte
    Server_IconProtected As Byte
    Server_IconIronSkin As Byte
    Server_IconSpellExhaustion As Byte
    Server_SetCharDamage As Byte
    Server_Ping As Byte
    Server_Help As Byte
    Server_Disconnect As Byte
    Map_LoadMap As Byte                 'Load Map - "SCM"
    Map_DoneLoadingMap As Byte          'Done Loading Map - "DLM"
    Map_RequestUpdate As Byte           'Request Map Update - "RMU"
    Map_StartTransfer As Byte           'Start Map Transfer - "SMT"
    Map_EndTransfer As Byte             'Done Transfering Map - "EMT"
    Map_DoneSwitching As Byte           'Done Switching Maps - "DSM"
    Map_SendName As Byte                'Request Map Name - "SMN"
    Map_RequestPositionUpdate As Byte   'Request Position Update - "RPU"
    Map_UpdateTile As Byte
    User_Target As Byte
    User_KnownSkills As Byte            'Request Known Skills
    User_Attack As Byte                 'User Attack - "ATT"
    User_SetInventorySlot As Byte       'Set User Inventory Slot - "SIS"
    User_Desc As Byte
    User_Login  As Byte                 'Log In Existing User - "LOGIN"
    User_NewLogin As Byte               'Create A New User - "NLOGIN"
    User_Get As Byte                    'User Get An Item Off Ground - "GET"
    User_Drop As Byte                   'User Drop An Item - "DRP"
    User_Use As Byte                    'User Use An Item - "USE"
    User_Move As Byte                   'Move User Character - "M"
    User_Rotate As Byte                 'Rotate User Character - "SUH"
    User_LeftClick As Byte              'User Left-Clicked Tile - "LC"
    User_RightClick As Byte             'User Right-Clicked Tile - "RC"
    User_LookLeft As Byte
    User_LookRight As Byte
    User_AggressiveFace As Byte
    User_Blink As Byte
    User_Trade_StartNPCTrade As Byte
    User_Trade_BuyFromNPC As Byte
    User_Trade_SellToNPC As Byte
    User_BaseStat As Byte
    User_ModStat As Byte
    User_CastSkill As Byte
    User_ChangeInvSlot As Byte
    User_Emote As Byte
    User_StartQuest As Byte
    Dev_SetSurface As Byte
    Dev_SetBlocked As Byte
    Dev_SetExit As Byte
    Dev_SetLight As Byte
    Dev_SetNPC As Byte
    Dev_SetMailbox As Byte
    Dev_SetObject As Byte
    Dev_SetMapInfo As Byte
    Dev_UpdateTile As Byte
    Dev_SaveMap As Byte
    Dev_SetMode As Byte
    Dev_SetTile As Byte
    GM_Approach As Byte
    GM_Summon As Byte
    GM_Kick As Byte
    GM_Raise As Byte
End Type
Public DataCode As DataCode

Public sndBuf As DataBuffer
Public ChatBufferChunk As Integer
Public PingSTime As Long
Public Ping As Long
Public SoxID As Long
Public SocketOpen As Byte
Public TargetCharIndex As Integer
Public Const DegreeToRadian As Single = 0.0174532925

'Holds the skin the user is using at the time
Public CurrentSkin As String

'If we are in windowed mode or not
Public Const Windowed As Boolean = False

'Blocked directions - take the blocked byte and OR these values (If Blocked OR <Direction> Then...)
Public Const BlockedNorth As Byte = 1
Public Const BlockedEast As Byte = 2
Public Const BlockedSouth As Byte = 4
Public Const BlockedWest As Byte = 8
Public Const BlockedAll As Byte = 15

'How many pings we have set with no return
Public NonRetPings As Byte

'States if the project is unloading (has to give Sox time to unload)
Public IsUnloading As Byte

'User login information
Public UserPassword As String
Public UserName As String

'********** OUTSIDE FUNCTIONS ***********
'For Get and Write Var
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:36)  Decl: 285  Code: 0  Total: 285 Lines
':) CommentOnly: 72 (25.3%)  Commented: 45 (15.8%)  Empty: 18 (6.3%)  Max Logic Depth: 1
