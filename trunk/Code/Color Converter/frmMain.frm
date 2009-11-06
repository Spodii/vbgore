VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "Color Conversion"
   ClientHeight    =   1170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   78
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   403
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cButton3 
      Caption         =   "Long -> ARGB"
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cButton4 
      Caption         =   "Long -> RGB"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cButton1 
      Caption         =   "ARGB -> Long"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cButton2 
      Caption         =   "RGB -> Long"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox LongTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      TabIndex        =   5
      Text            =   "0"
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox PreviewPic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2880
      ScaleHeight     =   255
      ScaleWidth      =   705
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox BTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "0"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox GTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "0"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox RTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "0"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox ATxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      Top             =   360
      Width           =   495
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
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
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

Private Sub cButton1_Click()

    On Error Resume Next
    
    LongTxt.Text = D3DColorARGB(ATxt.Text, RTxt.Text, GTxt.Text, BTxt.Text)

End Sub

Private Sub cButton2_Click()

    On Error Resume Next
    
    LongTxt.Text = RGB(RTxt.Text, GTxt.Text, BTxt.Text)
    
End Sub

Private Sub cButton3_Click()
Dim Dest(3) As Byte

    On Error Resume Next
    
    CopyMemory Dest(0), CLng(LongTxt.Text), 4
    ATxt.Text = Dest(3)
    RTxt.Text = Dest(2)
    GTxt.Text = Dest(1)
    BTxt.Text = Dest(0)
        
End Sub

Private Sub cButton4_Click()

    SplitRGB LongTxt.Text, RTxt.Text, GTxt.Text, BTxt.Text

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim c As Control
    
    For Each c In Me
        If TypeName(c) = "cButton" Then
            c.Refresh
            c.DrawState = 0
        End If
    Next c
    Set c = Nothing
    
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

