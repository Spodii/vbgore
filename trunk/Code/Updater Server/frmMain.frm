VERSION 5.00
Object = "{1FBEF8E7-C785-4338-83C0-B3688028151A}#1.0#0"; "vbgoresocketstring.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Update Server"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":1708A
   ScaleHeight     =   275
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Connection Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   3615
      Begin SoxOCX.Sox Sox 
         Height          =   420
         Left            =   2880
         Top             =   240
         Visible         =   0   'False
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin VB.Timer SecondTmr 
         Interval        =   1000
         Left            =   2400
         Top             =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Connections :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   615
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin VB.Label ConnsLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   11
         Top             =   480
         Width           =   90
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Upload Rate (Kb/s) :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label OutLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   90
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Current Transfer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3615
      Begin VB.Label ClientIDLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   16
         Top             =   960
         Width           =   90
      End
      Begin VB.Label ClientIPLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.0.0.0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.Label PercentLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   14
         Top             =   480
         Width           =   210
      End
      Begin VB.Label FileNameLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   13
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   0
         TabIndex        =   7
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Client IP :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   5
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "File Name :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Totals"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   3615
      Begin VB.Label MBoutTxt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1920
         TabIndex        =   20
         Top             =   720
         Width           =   90
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Data Out (Mb) :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   19
         Top             =   720
         Width           =   1860
      End
      Begin VB.Label ConnectionsEstablishedLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1920
         TabIndex        =   18
         Top             =   480
         Width           =   90
      End
      Begin VB.Label FilesUploadedLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1920
         TabIndex        =   17
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Connections Established :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   2
         Top             =   480
         Width           =   1860
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Files Uploaded :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   1860
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

'The update server works in the following manner:
'Step 1:
'-Server starts up, loads files and file information, the listens for connections
'-Client connects to server
'-Server recieves connection
'Step 2:
'-Server sends information of first file
'-Client reads first file's information, checks if the file is up to date
'*If file up to date, the client requests the next file's information (start of Step 2)
'*If file is not up to date, the client requests for the update
' -Server sends file to client
' -When client recieves the end of file, the client changes the file's information to match the server's
' -Client requests the next file (start of Step 2)
'Step 3:
'-Server loops through Step 2 until every file has been checked and updated
'-Server disconnects client

'Local socket ID
Private LocalID As Long

'User information
Private Const MaxConnections As Byte = 10
Private Type User
    ConnID As Integer   'ID of the user's socket
    CurrFile As Long    'File the user is currently downloading
    LastBytesSent As Long   'Used to calculate the amount sent from a file
End Type
Private UserList(1 To MaxConnections) As User

'Upload amount
Private Const UploadBufferSize As Byte = 1  'How many seconds to average over
Private UploadBuffer(1 To UploadBufferSize) As Long
Private UploadCount As Long 'Counts the amount uploaded (reset by timer)

'Used to calculate MBs uploaded
Private upBytes As Long
Private upKBytes As Long
Private upMBytes As Long

'Counter to clear the data (clearing after every transfer makes a lot of flashing - very annoying)
Private ClearDataCount As Byte

'How many people are currently connected
Private CurrConnections As Long

'Totals
Private FilesUploaded As Long   'Total amount of files uploaded
Private ConnectionsEst As Long  'Total amount of connections established

'File list information
Private FileList() As String            'List of files by their complete path on the server
Private FileListShortName() As String   'List of files by their shortened path
Private FileSize() As Long              'Size of the file in bytes
Private NumFiles As Long

'***** Used for changing file times *****
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_DEVICE = &H40
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_SPARSE_FILE = &H200
Private Const FILE_ATTRIBUTE_REPARSE_POINT = &H400
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_OFFLINE = &H1000
Private Const FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
Private Const FILE_ATTRIBUTE_ENCRYPTED = &H4000
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_SHARE_DELETE = &H4
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_ALL = &H10000000
Private Const DELETE = &H10000
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Const CREATE_NEW = 1
Private Const CREATE_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALWAYS = 4
Private Const TRUNCATE_EXISTING = 5

Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&

    'Close form
    If Button = vbLeftButton Then
        If X >= Me.ScaleWidth - 23 Then
            If X <= Me.ScaleWidth - 10 Then
                If Y <= 26 Then
                    If Y >= 11 Then
                        Unload Me
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub Form_Load()

Dim FileNum As Byte
Dim i As Long
Dim j As Integer

'Get the file list

    FileList() = AllFilesInFolders(App.Path & "\Files\", True)
    NumFiles = UBound(FileList())
    
    'Quit if we have no files to update
    If NumFiles = 0 Then
        MsgBox "Error: You must include files to update to run the server!", vbOKOnly
        Unload Me
        Exit Sub
    End If

    'Create the short file list
    ReDim FileListShortName(0 To NumFiles)
    j = Len(App.Path)
    For i = 0 To NumFiles
        FileListShortName(i) = Right$(FileList(i), Len(FileList(i)) - j)
    Next i

    'Create the file size list
    ReDim FileSize(0 To NumFiles)
    FileNum = FreeFile
    For i = 0 To NumFiles
        Open FileList(i) For Append As #FileNum
        FileSize(i) = LOF(FileNum)
        Close #FileNum
    Next i

    'Start up the socket
    LocalID = Sox.Listen("127.0.0.1", 10201)
    Sox.SetOption LocalID, soxSO_TCP_NODELAY, True

    'Show the form
    Me.Show

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case X
    Case MouseMove
    Case LeftUp
    Case LeftDown
        'Return from system tray
        If Me.WindowState = 1 Then
            TrayDelete
            Me.WindowState = 0
            Me.Show
        End If
    Case LeftDbClick
    Case RightUp
    Case RightDown
    Case RightDbClick
    End Select

End Sub

Private Sub Form_Resize()

'If the form becomes minimized, move to system tray

    If WindowState = 1 Then
        TrayAdd Me, "Update Server: " & CurrConnections & " connections", MouseMove
        Me.Hide
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Sox.ShutDown = soxERROR Then
        Let Cancel = True
    Else
        Sox.UnHook
    End If

End Sub

Private Sub SecondTmr_Timer()

Dim TotalValue As Long
Dim i As Integer

'Calculate total uploaded rate

    upBytes = upBytes + UploadCount
    Do While upBytes > 1024
        upKBytes = upKBytes + 1
        upBytes = upBytes - 1024
    Loop
    Do While upKBytes > 1024
        upMBytes = upMBytes + 1
        upKBytes = upKBytes - 1024
    Loop
    MBoutTxt.Caption = upMBytes

    'Place in the new value
    UploadBuffer(1) = UploadCount
    UploadCount = 0

    'Calculate the average
    For i = 1 To UploadBufferSize
        TotalValue = TotalValue + UploadBuffer(i)
    Next i
    OutLbl.Caption = (TotalValue \ UploadBufferSize) \ 1024

    'Figure out how many active connections there is
    CurrConnections = 0
    For i = 1 To MaxConnections
        If UserList(i).ConnID > 0 Then CurrConnections = CurrConnections + 1
    Next i
    ConnsLbl.Caption = CurrConnections
    TrayModify ToolTip, "Update Server: " & CurrConnections & " connections"

End Sub

Private Sub SendFileRequest(ByVal UserIndex As Long, ByVal FileIndex As Long)

Dim SECURITY_ATTRIBUTES As SECURITY_ATTRIBUTES
Dim File_CreationTime As FILETIME
Dim File_LastAccessTime As FILETIME
Dim File_LastWriteTime As FILETIME
Dim FileNum As Byte
Dim Temp As Long

'Check for valid file index

    If FileIndex < 0 Then Exit Sub
    If FileIndex > NumFiles Then Exit Sub

    'Get the file's LastWrite time
    SECURITY_ATTRIBUTES.nLength = Len(SECURITY_ATTRIBUTES)
    SECURITY_ATTRIBUTES.lpSecurityDescriptor = 0
    SECURITY_ATTRIBUTES.bInheritHandle = False
    Temp = CreateFile(FileList(FileIndex) & Chr$(0), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, SECURITY_ATTRIBUTES, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    DoEvents
    GetFileTime Temp, File_CreationTime, File_LastAccessTime, File_LastWriteTime
    DoEvents
    CloseHandle Temp

    'Send the request to start transfering the file
    Sox.SendData UserIndex, "****sendrequest****|" & FileListShortName(FileIndex) & "|" & FileSize(FileIndex) & "|" & File_LastWriteTime.dwLowDateTime & "|" & File_LastWriteTime.dwHighDateTime

End Sub

Private Sub Sox_OnConnection(inSox As Long)

'Assign the inSox ID as the ConnID

    UserList(inSox).ConnID = inSox

    'Start the transfering with the first file
    UserList(inSox).CurrFile = 0
    UserList(inSox).LastBytesSent = 0
    SendFileRequest inSox, UserList(inSox).CurrFile

    'Raise connections count
    ConnectionsEst = ConnectionsEst + 1
    frmMain.ConnectionsEstablishedLbl.Caption = ConnectionsEst

End Sub

Private Sub Sox_OnDataArrival(inSox As Long, inData() As Byte)

Dim SendBuffer As String
Dim FileNum As Byte
Dim rData As String
Dim FileLen As Long

'Turn the data into a string

    rData = StrConv(inData(), vbUnicode)

    'The file was accepted to be transfered
    If rData = "****ok****" Then

        'Send the whole file
        FileNum = FreeFile
        Open FileList(UserList(inSox).CurrFile) For Binary Access Read As #FileNum
        FileLen = LOF(FileNum)
        SendBuffer = Space$(FileLen)
        Get #FileNum, , SendBuffer
        Close #FileNum
        Sox.SendData inSox, SendBuffer & "****ENDOFFILE****"

        'Raise files uploaded count
        FilesUploaded = FilesUploaded + 1
        FilesUploadedLbl.Caption = FilesUploaded

        'Check if the user has finished updating
        If UserList(inSox).CurrFile > NumFiles Then
            'Tell the user they have finished
            Sox.SendData inSox, "****DONE****"
        Else
            'Send the next file
            SendFileRequest inSox, UserList(inSox).CurrFile
        End If
        Exit Sub

    End If

    'The file was not accepted to be transfered - the user must already be up to date
    If rData = "****no****" Then

        'Raise file number
        UserList(inSox).CurrFile = UserList(inSox).CurrFile + 1

        'Check if the user has finished updating
        If UserList(inSox).CurrFile > NumFiles Then
            'Tell the user they have finished
            Sox.SendData inSox, "****DONE****"
            UserList(inSox).ConnID = 0
            UserList(inSox).CurrFile = 0
            UserList(inSox).LastBytesSent = 0
        Else
            'Send the next file
            SendFileRequest inSox, UserList(inSox).CurrFile
        End If
        Exit Sub

    End If

End Sub

Private Sub Sox_OnSendComplete(inSox As Long)

'Clear information

    UserList(inSox).LastBytesSent = 0

End Sub

Private Sub Sox_OnSendProgress(inSox As Long, bytesSent As Long, bytesRemaining As Long)

'Update the sending information

    If UserList(inSox).CurrFile <= NumFiles Then FileNameLbl.Caption = FileListShortName(UserList(inSox).CurrFile)
    PercentLbl.Caption = Round(bytesSent / (bytesSent + bytesRemaining), 2) * 100 & "%"
    ClientIPLbl.Caption = Sox.Address(inSox)
    ClientIDLbl.Caption = inSox

    'Update the bytes remaining info
    UploadCount = UploadCount + (bytesSent - UserList(inSox).LastBytesSent)
    UserList(inSox).LastBytesSent = bytesSent

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:47)  Decl: 167  Code: 249  Total: 416 Lines
':) CommentOnly: 112 (26.9%)  Commented: 10 (2.4%)  Empty: 66 (15.9%)  Max Logic Depth: 3
