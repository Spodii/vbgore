VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "Update Server"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4515
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   251
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   301
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4335
      Begin VB.Timer SecondTmr 
         Interval        =   1000
         Left            =   2400
         Top             =   240
      End
      Begin VB.Label StatusLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   90
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
      Begin VB.Label ClientIDLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   4335
      Begin VB.Label MBoutTxt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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

'The update works in the following manner:
' - Server creates overall update list and MD5 hashes for each list, list & hashes is compressed and stored in memory
' - Client connects to the server and downloads the list off the server
' - Client decompresses the list, checks which files it needs to update
' - For every file the client needs to update, it sends a request to the server
' - After each download, the MD5 hash is compared with the one from the server to varify file contents

Private Sub Form_DblClick()

    'Send to system tray
    Me.WindowState = vbMinimized

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
