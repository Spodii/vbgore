VERSION 5.00
Begin VB.Form frmAnimate 
   BackColor       =   &H80000005&
   Caption         =   "Animation Creator"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13110
   Icon            =   "frmAnimate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   13110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      BackColor       =   &H80000005&
      Caption         =   "Frame"
      Height          =   1815
      Index           =   1
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   1335
      Begin VB.TextBox GrhTxt 
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.PictureBox FramePic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   1
         Left            =   120
         ScaleHeight     =   1065
         ScaleWidth      =   1065
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label GrhLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grh:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "General"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         MaxLength       =   6
         TabIndex        =   12
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox SpeedTxt 
         Height          =   285
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2265
         ScaleWidth      =   2505
         TabIndex        =   4
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox FramesTxt 
         Height          =   285
         Left            =   720
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grh Index:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed:"
         Height          =   195
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Animation Preview:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frames:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmAnimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private HighestLoadedIndex As Long

Private Sub Form_Load()
    
    'The first control is already loaded...
    HighestLoadedIndex = 1

End Sub

Private Sub Form_Resize()
Dim NumShortRows As Long
Dim i As Long

    '1440

    'First two rows
    NumShortRows = (frmAnimate.Width - Frame(1).Left - 250 - 120) \ Frame(1).Width
    For i = 1 To NumShortRows
        LoadFrame i, 3000 + ((i - 1) * 1440), 120
    Next i

End Sub

Private Sub LoadFrame(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)

    'Load a frame at the corresponding location
    If HighestLoadedIndex < Index Then
        HighestLoadedIndex = Index
        Load Frame(Index)
        Load GrhTxt(Index)
        Load GrhLbl(Index)
        Load FramePic(Index)
    End If

    Frame(Index).Left = X
    Frame(Index).Top = Y
    Frame(Index).Visible = True

    GrhTxt(Index).Left = 480
    GrhTxt(Index).Top = 240
    GrhTxt(Index).Visible = True
    
    GrhLbl(Index).Left = 120
    GrhLbl(Index).Top = 240
    GrhLbl(Index).Visible = True
    
    FramePic(Index).Left = 120
    FramePic(Index).Top = 600
    FramePic(Index).Visible = True
        
End Sub
