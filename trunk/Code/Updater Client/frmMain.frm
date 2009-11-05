VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{1FBEF8E7-C785-4338-83C0-B3688028151A}#1.0#0"; "vbgoresocketstring.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Update Client"
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3045
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":17D2A
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   203
   StartUpPosition =   2  'CenterScreen
   Begin SoxOCX.Sox Sox 
      Height          =   420
      Left            =   2400
      Top             =   720
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSComDlg.CommonDialog C1 
      Left            =   3600
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label ConnectCmd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   960
      TabIndex        =   9
      Top             =   2040
      Width           =   1020
   End
   Begin VB.Label StatusLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connecting..."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   8
      Top             =   1680
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   540
   End
   Begin VB.Label FileLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current File :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   885
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "KB/Sec"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label spid 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   1440
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Speed :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label PercentLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Download Status:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1260
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

'***** Used for changing file times *****
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
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
Private Type da
    FileToSend As String
    FileName As String
    RemoteIP As String
    FileSize As Double
    SaveAs As String
    Pstatus As Double
    lastamount As Double
    CreationTime As FILETIME
    AccessTime As FILETIME
    WriteTime As FILETIME
End Type
Private Info As da

Dim RecFileName As String   'File name (path) that we are recieving
Dim RecFileSize As Long     'Official file size recieved from server
Dim RecFileTime As FILETIME 'Official file time recieved from server
Dim WriteFileNum As Byte    'File number being written to
Dim FileHandle As Long      'File handle for getting/setting file time

'The access and creation time of the file - does not change on update, just stored to keep values same
Dim RecFileAccess As FILETIME
Dim RecFileCreation As FILETIME

Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long

Private Sub Connect()

'Set the status

    ConnectCmd.Enabled = False
    StatusLbl.Caption = "Connecting..."

    'Set up the socket
    LocalID = Sox.Connect("127.0.0.1", 10201)

    'Check for invalid LocalID (did not connect)
    If LocalID = -1 Then
        StatusLbl.Caption = "Unable to connect!"
        ConnectCmd.Enabled = True
    Else
        Sox.SetOption LocalID, soxSO_TCP_NODELAY, True
    End If

End Sub

Private Sub ConnectCmd_Click()

    Connect

End Sub

Function Engine_FileExist(file As String, FileType As VbFileAttribute) As Boolean

'Checks if a file exists

    Engine_FileExist = (Dir$(file, FileType) <> "")

End Function

Private Sub Form_Load()

    Me.Show
    DoEvents
    Connect

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&

    'Close form
    If Button = vbLeftButton Then
        If x >= Me.ScaleWidth - 23 Then
            If x <= Me.ScaleWidth - 10 Then
                If Y <= 26 Then
                    If Y >= 11 Then
                        Unload Me
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If FileHandle <> 0 Then CloseHandle FileHandle
    If Sox.ShutDown = soxERROR Then
        Let Cancel = True
    Else
        Sox.UnHook
    End If

End Sub

Private Sub MakeDir(Path As String)

'Make sure a directory exists - if it does not, it is created

    MkDir Path

End Sub

Private Sub Sox_OnClose(inSox As Long)

    ConnectCmd.Enabled = True

End Sub

Private Sub Sox_OnDataArrival(inSox As Long, inData() As Byte)

    On Error Resume Next
    Dim SECURITY_ATTRIBUTES As SECURITY_ATTRIBUTES
    Dim File_LastWriteTime As FILETIME
    Dim TempArray() As String
    Dim TempStr As String
    Dim AppPath As String
    Dim rData As String

        'Turn the data into a string
        rData = StrConv(inData(), vbUnicode)

        'Check for a send request
        If Left$(rData, 19) = "****sendrequest****" Then

            'Retrieve the file name, size and time
            TempArray = Split(rData, "|")
            RecFileName = TempArray(1)
            RecFileSize = TempArray(2)
            RecFileTime.dwLowDateTime = TempArray(3)
            RecFileTime.dwHighDateTime = TempArray(4)
            FileLbl.Caption = RecFileName & " (" & RecFileSize & ")"

            'If the application path is "x:\", set it to "x:" or else we will have "x:\\" (where x is any letter)
            AppPath = App.Path
            If Len(AppPath) = 3 Then AppPath = Left$(AppPath, 2)

            'Close old file if open
            If WriteFileNum > 0 Then Close #WriteFileNum
            WriteFileNum = 0

            'Create directory if not yet made
            MakeSureDirectoryPathExists AppPath & RecFileName

            'Open file to write to
            WriteFileNum = FreeFile
            Open AppPath & RecFileName For Binary Access Write As #WriteFileNum

            'Check if the file write time is same
            If Engine_FileExist(AppPath & RecFileName, vbNormal) Then
                If FileHandle <> 0 Then CloseHandle FileHandle  'Close file handle if open
                'Recieve the write time
                SECURITY_ATTRIBUTES.nLength = Len(SECURITY_ATTRIBUTES)
                SECURITY_ATTRIBUTES.lpSecurityDescriptor = 0
                SECURITY_ATTRIBUTES.bInheritHandle = False
                FileHandle = CreateFile(AppPath & RecFileName & Chr$(0), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, SECURITY_ATTRIBUTES, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
                GetFileTime FileHandle, RecFileCreation, RecFileAccess, File_LastWriteTime
                If File_LastWriteTime.dwLowDateTime = RecFileTime.dwLowDateTime Then
                    If File_LastWriteTime.dwHighDateTime = RecFileTime.dwHighDateTime Then
                        Debug.Print RecFileSize; LOF(WriteFileNum)
                        If RecFileSize = LOF(WriteFileNum) Then
                            'Write time is the same, so no point in updating
                            CloseHandle FileHandle
                            FileHandle = 0
                            Sox.SendData LocalID, "****no****"
                            Exit Sub
                        End If
                    End If
                End If
            End If

            'File write time was not the same or the file did not exist, so recieve the update
            Sox.SendData LocalID, "****ok****"
            StatusLbl.Caption = "Downloading..."
            Exit Sub

        End If

        'End of file reached
        If Right$(rData, 17) = "****ENDOFFILE****" Then
            TempStr = Left$(rData, Len(rData) - 17) 'Crop out the ENDOFFILE to recieve the last bit of data
            If Len(TempStr) Then Put #WriteFileNum, , TempStr   'Write the last data
            Close #WriteFileNum     'Close the file since we're done
            PercentLbl.Caption = "0%"

            'If the application path is "x:\", set it to "x:" or else we will have "x:\\" (where x is any letter)
            AppPath = App.Path
            If Len(AppPath) = 3 Then AppPath = Left$(AppPath, 2)

            'If we dont have the handle already, then get it
            If FileHandle = 0 Then
                SECURITY_ATTRIBUTES.nLength = Len(SECURITY_ATTRIBUTES)
                SECURITY_ATTRIBUTES.lpSecurityDescriptor = 0
                SECURITY_ATTRIBUTES.bInheritHandle = False
                FileHandle = CreateFile(AppPath & RecFileName & Chr$(0), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, SECURITY_ATTRIBUTES, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
                GetFileTime FileHandle, RecFileCreation, RecFileAccess, File_LastWriteTime
            End If

            'Set the file time
            SetFileTime FileHandle, RecFileCreation, RecFileAccess, RecFileTime

            'Close the handle off
            CloseHandle FileHandle
            FileHandle = 0
            Exit Sub
        End If

        'Done downloading
        If Right$(rData, 12) = "****DONE****" Then
            Sox.Shut LocalID
            Sox.ShutDown
            StatusLbl.Caption = "Download Successful!"
            FileLbl.Caption = ""
            PercentLbl.Caption = "100%"
            Exit Sub
        End If

        'If not the above, then we are *hopefully* recieving the file data
        Put #WriteFileNum, , rData

End Sub

Private Sub Sox_OnRecvProgress(inSox As Long, bytesRecv As Long, bytesRemaining As Long)

'Update recieved percentage

    PercentLbl.Caption = (Round(bytesRecv / (bytesRecv + bytesRemaining), 2) * 100) & "%"

End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub Timer2_Timer()

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:45)  Decl: 167  Code: 199  Total: 366 Lines
':) CommentOnly: 98 (26.8%)  Commented: 9 (2.5%)  Empty: 56 (15.3%)  Max Logic Depth: 6
