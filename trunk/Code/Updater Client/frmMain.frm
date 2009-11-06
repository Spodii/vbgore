VERSION 5.00
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
   Begin VB.Timer CloseTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2520
      Top             =   600
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

Private Sub Connect()

    'Set the status
    ConnectCmd.Enabled = False
    StatusLbl.Caption = "Connecting..."

    'Set up the socket
    LocalID = GOREsock_Connect("127.0.0.1", 10201)
    
     'Check for invalid LocalID (did not connect)
    If LocalID = -1 Then
        StatusLbl.Caption = "Unable to connect!"
        ConnectCmd.Enabled = True
    Else
        GOREsock_SetOption LocalID, soxSO_TCP_NODELAY, False
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

Private Sub Form_Load()
    
    GOREsock_Initialize Me.hWnd
    InitFilePaths
    Me.Show
    DoEvents
    Connect

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&

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

    If GOREsock_ShutDown = soxERROR Then
        Cancels = Cancels + 1
        If Cancels < 3 Then
            Let Cancel = True
        Else
            GOREsock_UnHook  'Force unload
        End If
    Else
        GOREsock_UnHook
    End If

End Sub
