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
'************            Official Release: Version 0.2.0            ************
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

'********** Debug/Display Settings ************
'These are your key constants - reccomended you turn off ALL debug constants before
' compiling your code for public usage just speed reasons

'These two are mostly used for checking to make sure the encryption works
Public Const DEBUG_PrintPacketReadErrors As Boolean = False 'Will print the packet read errors in debug.print
Public Const DEBUG_PrintPacket_In As Boolean = False     'Shows packets coming in in chat box
Public Const DEBUG_PrintPacket_Out As Boolean = False    'Shows packets going out in chat box

'Set this to true to force updater check
Public Const ForceUpdateCheck As Boolean = False

'********** Object types ************
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
    Price As Long               'Price of the object
End Type

Public BaseStats(1 To NumStats) As Long
Public ModStats(1 To NumStats) As Long

'If the map is loading (used to be used for the downloading status of maps)
Public DownloadingMap As Boolean

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
    GrhIndex As Long
    Amount As Integer
    Equipped As Boolean
End Type

'Messages
Public NumMessages As Byte
Public Message() As String

'Known user skills/spells
Public UserKnowSkill(1 To NumSkills)

'User status vars
Public UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory

'Used during login
Public SendNewChar As Boolean

'Control
Public prgRun As Boolean 'When true the program ends

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

'Zoom level - 0 = No Zoom, > 0 = Zoomed
Public ZoomLevel As Single
Public Const MaxZoomLevel As Single = 0.25

'********** OUTSIDE FUNCTIONS ***********
'For Get and Write Var
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:36)  Decl: 285  Code: 0  Total: 285 Lines
':) CommentOnly: 72 (25.3%)  Commented: 45 (15.8%)  Empty: 18 (6.3%)  Max Logic Depth: 1
