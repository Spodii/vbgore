VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{598D2D95-4E74-49D9-8F45-E9E53990E851}#1.0#0"; "goresockfull.ocx"
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
   Begin GORESOCKfull.Socket Socket 
      Height          =   660
      Left            =   2280
      Top             =   480
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   1164
   End
   Begin VB.Timer CloseTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2400
      Top             =   1200
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
Private Type da
    FileToSend As String
    FileName As String
    RemoteIP As String
    FileSize As Double
    SaveAs As String
    PStatus As Double
    LastAmount As Double
End Type
Private Info As da

Dim RecFileName As String       'File name (path) that we are recieving
Dim RecFileSize As Long         'Official file size recieved from server
Dim RecFileHash As String * 32  'Hash received from the server

Dim WriteFileNum As Byte    'File number being written to

Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

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

Private Sub Connect()

    'Set the status
    ConnectCmd.Enabled = False
    StatusLbl.Caption = "Connecting..."

    'Set up the socket
    LocalID = Socket.Connect("127.0.0.1", 10201)
    
     'Check for invalid LocalID (did not connect)
    If LocalID = -1 Then
        StatusLbl.Caption = "Unable to connect!"
        ConnectCmd.Enabled = True
    Else
        Socket.SetOption LocalID, soxSO_TCP_NODELAY, True
    End If
    
End Sub

Private Sub CloseTimer_Timer()

    'Quit the updater - we must user a timer since DoEvents wont work (since we're not multithreaded)
    Unload Me
    End

End Sub

Private Sub ConnectCmd_Click()

    Connect

End Sub

Function Engine_FileExist(File As String, FileType As VbFileAttribute) As Boolean

'Checks if a file exists

    Engine_FileExist = (Dir$(File, FileType) <> "")

End Function

Private Sub Form_Load()

    InitFilePaths
    Me.Show
    DoEvents
    Connect

End Sub

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

Private Sub Form_Unload(Cancel As Integer)
Static Cancels As Byte

    If Socket.ShutDown = soxERROR Then
        Cancels = Cancels + 1
        If Cancels < 3 Then
            Let Cancel = True
        Else
            Socket.UnHook  'Force unload
        End If
    Else
        Socket.UnHook
    End If

End Sub

Private Sub socket_OnClose(inSox As Long)

    ConnectCmd.Enabled = True

End Sub

Private Sub socket_OnDataArrival(inSox As Long, inData() As Byte)
Dim TempArray() As String
Dim TempStr As String
Dim AppPath As String
Dim FileNum As Byte
Dim rData As String
Dim i As Integer
Dim j As Long
Dim b() As Byte

    'Turn the data into a string
    rData = StrConv(inData(), vbUnicode)

    'Check for a send request
    If Left$(rData, 19) = "****sendrequest****" Then

        'Retrieve the file name, size and time
        TempArray = Split(rData, "|")
        RecFileName = TempArray(1)
        RecFileSize = TempArray(2)
        RecFileHash = TempArray(3)
        FileLbl.Caption = RecFileName & " (" & RecFileSize & ")"

        'If the application path is "x:\", set it to "x:" or else we will have "x:\\" (where x is any letter)
        AppPath = App.Path
        If Len(AppPath) = 3 Then AppPath = Left$(AppPath, 2)

        'Close old file if open
        If WriteFileNum > 0 Then Close #WriteFileNum
        WriteFileNum = 0

        'Create directory if not yet made
        MakeSureDirectoryPathExists AppPath & RecFileName & ".compressed"

        'Open file to write to
        WriteFileNum = FreeFile
        Open AppPath & RecFileName & ".compressed" For Binary Access Write As #WriteFileNum

        'Check if the file write time is same
        If Engine_FileExist(AppPath & RecFileName, vbNormal) Then
            
            'If the file does exist, compair it to the recieved MD5 hash to see if it is the same file
            TempStr = MD5_File(App.Path & RecFileName)
            If TempStr = RecFileHash Then
            
                'The file is the same, do not update
                b() = StrConv("****no****", vbFromUnicode)
                Socket.SendData LocalID, b()
                Exit Sub
                
            End If

        End If

        'File is different, request the update
        b() = StrConv("****ok****", vbFromUnicode)
        Socket.SendData LocalID, b()
        StatusLbl.Caption = "Downloading..."
        Exit Sub

    End If

    'End of file reached
    If Right$(rData, 17) = "****ENDOFFILE****" Then
        TempStr = Left$(rData, Len(rData) - 17) 'Crop out the ENDOFFILE to recieve the last bit of data
        If Len(TempStr) Then Put #WriteFileNum, , TempStr   'Write the last data
        Close #WriteFileNum     'Close the file since we're done
        Compression_DeCompress App.Path & RecFileName & ".compressed", App.Path & RecFileName, RLE_Loop 'Take the compressed file and decompress it
        'Kill App.Path & RecFileName & ".compressed" 'Kill the compressed file
        PercentLbl.Caption = "0%"
        Exit Sub
    End If

    'Done downloading
    If Right$(rData, 12) = "****DONE****" Then
        Socket.Shut LocalID
        Socket.ShutDown
        StatusLbl.Caption = "Download Successful!"
        FileLbl.Caption = ""
        PercentLbl.Caption = "100%"
        
        'Make sure the file is closed
        On Error Resume Next
        Close #WriteFileNum
        On Error GoTo 0
        
        'Clean up the compressed files
        TempArray = AllFilesInFolders(App.Path, True)
        For j = 1 To UBound(TempArray)
            If Right$(TempArray(j), 11) = ".compressed" Then Kill TempArray(j)
        Next j
        
        'Load the client
        If MsgBox("The update has been completed! Do you wish to run the client now?", vbYesNo) = vbYes Then
            ShellExecute Me.hwnd, vbNullString, App.Path & "\GameClient.exe", "-sdf@041jkdf0)21`~", vbNullString, 1
        End If
        
        'Unload the updater
        Socket.Shut LocalID
        DoEvents
        Socket.UnHook
        DoEvents
        
        'Initiate the closedown (gives the socket time to unload)
        CloseTimer.Enabled = True
        
        Exit Sub
    End If

    'If not the above, then we are *hopefully* recieving the file data
    Put #WriteFileNum, , rData

End Sub

Private Sub socket_OnRecvProgress(inSox As Long, bytesRecv As Long, bytesRemaining As Long)

    'Update recieved percentage
    PercentLbl.Caption = (Round(bytesRecv / (bytesRecv + bytesRemaining), 2) * 100) & "%"

End Sub
