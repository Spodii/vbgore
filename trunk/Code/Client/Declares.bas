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
'************            Official Release: Version 0.5.5            ************
'************                 http://www.vbgore.com                 ************
'*******************************************************************************
'*******************************************************************************
'***** License Information For General Users: **********************************
'*******************************************************************************
'** vbGORE comes completely free. You may charge for people to play your game **
'** along with you may accept donations for your game. The only rules you     **
'** must follow when using vbGORE are:                                        **
'**  - You may not claim the engine as your own creation.                     **
'**  - You may not sell the code to vbGORE in any way or form, whether it is  **
'**    the original vbGORE code or a modified version of it. Selling your game**
'**    may be an exception - if you wish to do this, you must first acquire   **
'**    permission from Spodi directly.                                        **
'**  - If you are distributing vbGORE or modified code of it, read the        **
'**    section "Source Distrubution Information" below.                       **
'** This license is subject to change at any time. Only the most current      **
'** version of the license applies, not the copy of the license that came with**
'** your copy of vbGORE. This means if rules are added for version 1.0, you   **
'** can not avoid them by using a previous version such as 0.3.               **
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
'**  *Spread The Word - The more people who know about vbGORE, the more people**
'**        there is to report bugs and suggestions to improve the engine!     **
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
'** Chase: Help with programming, bug reports, and adding the trading system  **
'** Nex666: Help with mapping, graphics, bug reports, hosting, etc            **
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
'**                                                                           **
'** Also, all the members of the vbGORE community who have submitted          **
'** tutorials, bugs, suggestions, criticism and have just stuck around!!      **
'**                                                                           **
'** If you feel you belong in these credits, please contact Spodi (above).    **
'*******************************************************************************
'*******************************************************************************

Option Explicit

'********** Debug/Display Settings ************
'These are your key constants - reccomended you turn off ALL debug constants before
' compiling your code for public usage just speed reasons

'These two are mostly used for checking to make sure the encryption works
Public Const DEBUG_PrintPacketReadErrors As Boolean = False 'Shows command IDs that arn't used being processed and such
Public Const DEBUG_PrintPacket_In As Boolean = False        'Shows packets coming in in chat box

'Set this to true to force updater check
Public Const ForceUpdateCheck As Boolean = False

'Running speed - make sure you have the same value on the server!
Public Const RunningSpeed As Byte = 5
Public Const RunningCost As Long = 1    'How much stamina it cost to run

'Max chat bubble width
Public Const BubbleMaxWidth As Long = 140

'Word filter - use by "word-filterto,nextword-nextfilterto"... etc
Public Const FilterString As String = "fuck-****,shit-****,ass-***,bitch-*****"
Public FilterFind() As String
Public FilterReplace() As String

'********** NPC chat info ************
Public Type NPCChatLineCondition
    Condition As Byte           'The condition used (see NPCCHAT_COND_)
    Value As Long               'Used to hold a numeric condition value
    ValueStr As String          'Used to hold a value for SAY conditions
End Type
Public Type NPCChatLine
    NumConditions As Byte       'Total number of conditions
    Conditions() As NPCChatLineCondition
    Text As String              'The text that will be said
    Style As Byte               'The style used for the text (see NPCCHAT_STYLE_)
    Delay As Integer            'The delay time applied after saying this line
End Type
Public Type NPCChatAskAnswer    'The individual chat input answers
    Text As String              'The answer string
    GotoID As Byte              'ID the answer will move to
End Type
Public Type NPCChatAskLine      'Individual chat input lines
    Question As String          'The question text
    NumAnswers As Byte          'Number of answers that can be used
    Answer() As NPCChatAskAnswer
End Type
Public Type NPCChatAsk          'Chat input information (ASK parameters)
    StartAsk As Byte            'ID to start the asking on
    Ask() As NPCChatAskLine     'Holds all the ASK questions/responses
End Type
Public Type NPCChat
    Format As Byte              'Format of the chat (see NPCCHAT_FORMAT_)
    ChatLine() As NPCChatLine   'The information on the chat line
    NumLines As Byte            'The number of chat lines
    Distance As Long            'The distance the user must be from the NPC to activate the chat
    Ask As NPCChatAsk           'All the ASK information
End Type
Public NPCChat() As NPCChat

'Conditions (this are used as bit-flags, so only use powers of 2!)
Public Const NPCCHAT_COND_NUMCONDITIONS As Byte = 7
Public Const NPCCHAT_COND_LEVELLESSTHAN As Long = 2 ^ 0
Public Const NPCCHAT_COND_LEVELMORETHAN As Long = 2 ^ 1
Public Const NPCCHAT_COND_HPLESSTHAN As Long = 2 ^ 2
Public Const NPCCHAT_COND_HPMORETHAN As Long = 2 ^ 3
Public Const NPCCHAT_COND_KNOWSKILL As Long = 2 ^ 4
Public Const NPCCHAT_COND_DONTKNOWSKILL As Long = 2 ^ 5
Public Const NPCCHAT_COND_SAY As Long = 2 ^ 6

'Chat formats
Public Const NPCCHAT_FORMAT_RANDOM As Byte = 0
Public Const NPCCHAT_FORMAT_LINEAR As Byte = 1

'Chat sytles
Public Const NPCCHAT_STYLE_BOTH As Byte = 0
Public Const NPCCHAT_STYLE_BOX As Byte = 1
Public Const NPCCHAT_STYLE_BUBBLE As Byte = 2

'Client character types
Public Const ClientCharType_PC As Byte = 1
Public Const ClientCharType_NPC As Byte = 2
Public Const ClientCharType_Grouped As Byte = 3
Public Const ClientCharType_Slave As Byte = 4

'********** Object info ************
Public Type ObjData
    Name As String              'Name
    ObjType As Byte             'Type (armor, weapon, item, etc)
    GrhIndex As Long            'Graphic index
    MinHP As Integer            'Bonus HP regenerated
    MaxHP As Integer            'Bonus Max HP raised
    MinHIT As Integer           'Bonus minimum hit
    MaxHIT As Integer           'Bonus maximum hit
    DEF As Integer              'Bonus defence
    ArmorIndex As Byte          'Index of the body sprite
    WeaponIndex As Byte         'Index of the weapon sprite
    WeaponType As Byte          'What type of weapon, if it is a weapon
    Value As Long               'Value of the object
End Type

'********** Trade table ************
Public Type TradeObj
    Name As String
    Grh As Long
    Amount As Integer
    Value As Long
End Type
Public Type TradeTable
    User1Name As String              'The name of the table
    User2Name As String
    User1Accepted As Byte
    User2Accepted As Byte
    Trade1(1 To 9) As TradeObj  'The objects both indexes have entered
    Trade2(1 To 9) As TradeObj
    Gold1 As Long               'The gold both indexes have entered
    Gold2 As Long
    MyIndex As Byte             'States whether this client is index 1 or 2
End Type
Public TradeTable As TradeTable

'********** Other stuff ************
Public BaseStats(1 To NumStats) As Long
Public ModStats(FirstModStat To NumStats) As Long

'Delay timers for packet-related actions (so to not spam the server)
Public Const AttackDelay As Long = 200  'These constants are client-side only
Public Const LootDelay As Long = 500    ' - changing these lower wont make it faster server-side!
Public LastAttackTime As Long
Public LastLootTime As Long

'If the map is loading (used to be used for the downloading status of maps)
Public DownloadingMap As Boolean

'Item description variables
Public ItemDescWidth As Long
Public ItemDescLine(1 To 10) As String  'Allow 10 lines maximum
Public ItemDescLines As Byte

'Object constants
Public Const MAX_INVENTORY_SLOTS As Byte = 49

'Active ASK information
Public Type ActiveAsk
    AskName As String
    AskIndex As Byte
    ChatIndex As Byte
    QuestionTxt As String
End Type
Public ActiveAsk As ActiveAsk

'User's inventory
Type Inventory
    ObjIndex As Long
    Name As String
    GrhIndex As Long
    Amount As Integer
    Equipped As Boolean
    Value As Long
End Type

'Quest information
Type QuestInfo
    Name As String
    Desc As String
End Type
Public QuestInfo() As QuestInfo
Public QuestInfoUBound As Byte

'Messages
Public NumMessages As Byte
Public Message() As String

'Signs
Public Signs() As String

'Known user skills/spells
Public UserKnowSkill(1 To NumSkills) As Byte

'Attack range
Public UserAttackRange As Byte

'User status vars
Public UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory
Public UserBank(1 To MAX_INVENTORY_SLOTS) As Inventory

'If there is a clear path to the target (if any)
Public ClearPathToTarget As Byte

'Used during login
Public SendNewChar As Boolean

Public sndBuf As DataBuffer
Public ChatBufferChunk As Single
Public SoxID As Long
Public SocketMoveToIP As String
Public SocketMoveToPort As Integer
Public SocketOpen As Byte
Public TargetCharIndex As Integer
Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180
Public Const RadianToDegree As Single = 57.2958279087977 '180 / Pi

'Mail sending spam prevention
Public LastMailSendTime As Long

'Holds the skin the user is using at the time
Public CurrentSkin As String

'Blocked directions - take the blocked byte and OR these values (If Blocked OR <Direction> Then...)
Public Const BlockedNorth As Byte = 1
Public Const BlockedEast As Byte = 2
Public Const BlockedSouth As Byte = 4
Public Const BlockedWest As Byte = 8
Public Const BlockedAll As Byte = 15

Public UseSfx As Byte
Public UseMusic As Byte

'States if the project is unloading (has to give Sox time to unload)
Public IsUnloading As Byte

'User login information
Public UserPassword As String
Public UserName As String
Public UserClass As Byte
Public UserBody As Byte
Public UserHead As Byte

'Holds the name of the last person to whisper to the client
Public LastWhisperName As String

'Zoom level - 0 = No Zoom, > 0 = Zoomed
Public ZoomLevel As Single
Public Const MaxZoomLevel As Single = 0.3

'Cursor flash rate
Public Const CursorFlashRate As Long = 450

'If click-warping is on or not (can only be used by GMs)
Public UseClickWarp As Byte

'Emoticon delay
Public EmoticonDelay As Long

'How long char remains aggressive-faced after being attacked
Public Const AGGRESSIVEFACETIME = 4000

'Maximum variable sizes
Public Const MAXLONG As Long = (2 ^ 31) - 1
Public Const MAXINT As Integer = (2 ^ 15) - 1
Public Const MAXBYTE As Byte = (2 ^ 8) - 1

'********** OUTSIDE FUNCTIONS ***********
Public Declare Function GetKeyState Lib "User32" (ByVal nVirtKey As Long) As Integer
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Public Declare Function GetActiveWindow Lib "User32" () As Long
