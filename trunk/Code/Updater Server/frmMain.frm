VERSION 5.00
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

Private Sub Form_DblClick()

    'Send to system tray
    Me.WindowState = vbMinimized

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

Private Sub Form_Load()

    'Show the form
    Me.Show
    DoEvents
    
    'Load
    Initialize
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If X = LeftDown Then
        'Return from system tray
        If Me.WindowState = 1 Then
            TrayDelete
            Me.WindowState = 0
            Me.Show
        End If
    End If

End Sub

Private Sub Form_Resize()

'If the form becomes minimized, move to system tray

    If WindowState = 1 Then
        TrayAdd Me, "Update Server: " & CurrConnections & " connections", MouseMove
        Me.Hide
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Close down the socket
    If GOREsock_ShutDown = soxERROR Then Cancel = 1
    GOREsock_UnHook
    If GOREsock.Loaded Then
        GOREsock_Terminate
        Cancel = 1
    End If
    
    If Cancel <> 1 Then
        End
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
