VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Color Conversion"
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5745
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   122
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   383
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox LongTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3960
      TabIndex        =   5
      Text            =   "0"
      Top             =   840
      Width           =   1575
   End
   Begin VB.PictureBox PreviewPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2880
      ScaleHeight     =   255
      ScaleWidth      =   705
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox BTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "0"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox GTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "0"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox RTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "0"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox ATxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Long -> ARGB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3120
      TabIndex        =   10
      Top             =   1440
      Width           =   1470
   End
   Begin VB.Label Command4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Long -> RGB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3120
      TabIndex        =   9
      Top             =   1200
      Width           =   1320
   End
   Begin VB.Label Command2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RGB -> Long"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   840
      TabIndex        =   8
      Top             =   1200
      Width           =   1320
   End
   Begin VB.Label Command1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ARGB -> Long"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   720
      TabIndex        =   7
      Top             =   1440
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alpha  Red  Green  Blue    Preview       Result Long"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5220
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long

Private Sub ATxt_Change()

    If IsNumeric(ATxt.Text) = False Then
        ATxt.Text = ""
        Exit Sub
    End If

End Sub

Private Sub ATxt_GotFocus()

    On Error Resume Next
        ATxt.SelStart = 0
        ATxt.SelLength = Len(ATxt.Text)

End Sub

Private Sub BTxt_Change()

    If IsNumeric(BTxt.Text) = False Then
        BTxt.Text = ""
        Exit Sub
    End If
    UpdatePreview

End Sub

Private Sub BTxt_GotFocus()

    On Error Resume Next
        BTxt.SelStart = 0
        BTxt.SelLength = Len(BTxt.Text)

End Sub

Private Sub Command1_Click()

    On Error Resume Next
        LongTxt.Text = D3DColorARGB(ATxt.Text, RTxt.Text, GTxt.Text, BTxt.Text)

End Sub

Private Sub Command2_Click()

    On Error Resume Next
        LongTxt.Text = RGB(RTxt.Text, GTxt.Text, BTxt.Text)

End Sub

Private Sub Command4_Click()

    SplitRGB LongTxt.Text, RTxt.Text, GTxt.Text, BTxt.Text

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

Private Sub GTxt_Change()

    If IsNumeric(GTxt.Text) = False Then
        GTxt.Text = ""
        Exit Sub
    End If
    UpdatePreview

End Sub

Private Sub GTxt_GotFocus()

    On Error Resume Next
        GTxt.SelStart = 0
        GTxt.SelLength = Len(GTxt.Text)

End Sub

Private Sub Label2_Click()
Dim Dest(3) As Byte

    On Error Resume Next
        CopyMemory Dest(0), CLng(LongTxt.Text), 4
        ATxt.Text = Dest(3)
        RTxt.Text = Dest(2)
        GTxt.Text = Dest(1)
        BTxt.Text = Dest(0)
        
End Sub

Private Sub LongTxt_GotFocus()

    On Error Resume Next
        LongTxt.SelStart = 0
        LongTxt.SelLength = Len(LongTxt.Text)

End Sub

Private Sub RTxt_Change()

    If IsNumeric(RTxt.Text) = False Then
        RTxt.Text = ""
        Exit Sub
    End If
    UpdatePreview

End Sub

Private Sub RTxt_GotFocus()

    On Error Resume Next
        RTxt.SelStart = 0
        RTxt.SelLength = Len(RTxt.Text)

End Sub

Private Sub SplitRGB(ByVal lColor As Long, ByRef lRed As Long, ByRef lGreen As Long, ByRef lBlue As Long)

    lRed = lColor And &HFF
    lGreen = (lColor And &HFF00&) \ &H100&
    lBlue = (lColor And &HFF0000) \ &H10000

End Sub

Private Sub UpdatePreview()

    On Error Resume Next
        PreviewPic.BackColor = RGB(RTxt.Text, GTxt.Text, BTxt.Text)

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:50)  Decl: 57  Code: 117  Total: 174 Lines
':) CommentOnly: 55 (31.6%)  Commented: 0 (0%)  Empty: 29 (16.7%)  Max Logic Depth: 2
