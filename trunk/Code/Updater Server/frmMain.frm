VERSION 5.00
Object = "{E4B113F2-DDE2-4AB3-AEA1-60C47D60380C}#1.0#0"; "vbgoresocketstring.ocx"
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
   StartUpPosition =   2  'CenterScreen
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
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   480
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
      Begin VB.Label StatusLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   22
         Top             =   720
         Width           =   390
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   9
         Left            =   1035
         TabIndex        =   21
         Top             =   720
         Width           =   540
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
Private FileHash() As String * 32       'MD5 hash of the file
Private NumFiles As Long

'Path to the compressed file (CompressPath + FileListShortName)
Private CompressPath As String

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
 
Sub Initialize()
Dim FileNum As Byte
Dim i As Long
Dim j As Integer

    InitFilePaths
    
    StatusLbl.Caption = "Loading file list"
    DoEvents
    
    'Set the compressed path
    CompressPath = App.Path & "\_Compressed"

    'Get the file list
    FileList() = AllFilesInFolders(App.Path & "\UpdateFiles\", True)
    NumFiles = UBound(FileList())
    
    'Quit if we have no files to update
    If NumFiles = 0 Then
        MsgBox "Error: You must include files to update to run the server!" & vbCrLf & _
            "Place them in the following path:" & vbCrLf & vbCrLf & _
            App.Path & "\UpdateFiles\", vbOKOnly
        Unload Me
        Exit Sub
    End If

    'Create the short file list
    ReDim FileListShortName(0 To NumFiles)
    j = Len(App.Path & "UpdateFiles\")
    For i = 0 To NumFiles
        FileListShortName(i) = Right$(FileList(i), Len(FileList(i)) - j)
    Next i
    
    StatusLbl.Caption = "Loading file sizes"
    DoEvents

    'Create the file size list
    ReDim FileSize(0 To NumFiles)
    FileNum = FreeFile
    For i = 0 To NumFiles
        Open FileList(i) For Append As #FileNum
        FileSize(i) = LOF(FileNum)
        Close #FileNum
    Next i
    
    StatusLbl.Caption = "Creating MD5 hashes"
    DoEvents
    
    'Create MD5 hashes
    ReDim FileHash(0 To NumFiles)
    For i = 0 To NumFiles
        FileHash(i) = MD5_File(FileList(i))
    Next i
    
    StatusLbl.Caption = "Compressing files"
    DoEvents
    
    'Create compressed files
    For i = 0 To NumFiles
        If Engine_FileExist(App.Path & "\_Compressed" & FileListShortName(i), vbNormal) Then Kill App.Path & "\_Compressed" & FileListShortName(i)
        MakeSureDirectoryPathExists App.Path & "\_Compressed" & FileListShortName(i)
        Compression_Compress FileList(i), App.Path & "\_Compressed" & FileListShortName(i), RLE_Loop
    Next i
    
    StatusLbl.Caption = "Creating socket"
    DoEvents

    'Start up the socket
    LocalID = Sox.Listen(Var_Get(ServerDataPath & "Server.ini", "INIT", "UpdateIP"), Val(Var_Get(ServerDataPath & "Server.ini", "INIT", "UpdatePort")))
    Sox.SetOption LocalID, soxSO_TCP_NODELAY, True
    
    If frmMain.Sox.Address(LocalID) = "-1" Then MsgBox "Error while creating server connection. Please make sure you are connected to the internet and supplied a valid IP" & vbCrLf & "Make sure you use your INTERNAL IP, which can be found by Start -> Run -> 'Cmd' (Enter) -> IPConfig" & vbCrLf & "Finally, make sure you are NOT running another instance of the server, since two applications can not bind to the same port. If problems persist, you can try changing the port.", vbOKOnly

    StatusLbl.Caption = "Loaded!"

End Sub

Private Sub Form_DblClick()

    'Send to system tray
    Me.WindowState = vbMinimized

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

Private Sub Form_Load()

    'Show the form
    Me.Show
    DoEvents
    
    'Load
    Initialize
    
End Sub

Function Engine_FileExist(File As String, FileType As VbFileAttribute) As Boolean

'Checks if a file exists

    Engine_FileExist = (Dir$(File, FileType) <> "")

End Function

Private Function Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

'*****************************************************************
'Gets a variable from a text file
'*****************************************************************

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

    szReturn = vbNullString

    sSpaces = Space$(1000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish

    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File

    Var_Get = RTrim$(sSpaces)
    Var_Get = Left$(Var_Get, Len(Var_Get) - 1)

End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Select Case x
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
Dim FileNum As Byte

    'Check for valid file index
    If FileIndex < 0 Then Exit Sub
    If FileIndex > NumFiles Then Exit Sub

    'Send the request to start transfering the file
    DoEvents
    Sox.SendData UserIndex, "****sendrequest****|" & FileListShortName(FileIndex) & "|" & FileSize(FileIndex) & "|" & FileHash(FileIndex)

End Sub

Private Sub Sox_OnClose(inSox As Long)
Dim i As Long
    
    'Remove the user's ConnID if they are on the socket that just closed
    For i = 1 To UBound(UserList)
        If UserList(i).ConnID = inSox Then
            UserList(i).ConnID = 0
            Exit For
        End If
    Next i

End Sub

Private Sub Sox_OnConnection(inSox As Long)

    'Assign the inSox ID as the ConnID
    UserList(inSox).ConnID = inSox

    'Start the transfering with the first file
    UserList(inSox).CurrFile = 0
    UserList(inSox).LastBytesSent = 0
    DoEvents
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
        Open CompressPath & FileListShortName(UserList(inSox).CurrFile) For Binary Access Read As #FileNum
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
