VERSION 5.00
Begin VB.Form frmMapInfo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Map Info"
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMapInfo.frx":0000
   ScaleHeight     =   110
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   225
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox MusicTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "ID of the music file to be played in the map. 0 for nothing."
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox WeatherTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2640
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "What kind of weather goes on on the map - 0 for none"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox VersionTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Text            =   "1"
      ToolTipText     =   "Version of the map - if the client's version does not match the server's version, the map will be automatically updated"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox MapNameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Text            =   "Name"
      ToolTipText     =   "Name of the map"
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Music:"
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
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   570
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Weather:"
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
      Height          =   195
      Index           =   2
      Left            =   1800
      TabIndex        =   6
      Top             =   960
      Width           =   795
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
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
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   705
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Map Name:"
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
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmMapInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = 1
    Engine_Var_Write Ini2Path & "MapEditor.ini", "MAPINFO", "X", Me.Left
    Engine_Var_Write Ini2Path & "MapEditor.ini", "MAPINFO", "Y", Me.Top
    HideFrmMapInfo

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

Private Sub MapNameTxt_Change()

    MapInfo.Name = MapNameTxt.Text

End Sub

Private Sub MusicTxt_Change()

    If MusicTxt.Text = "" Then MusicTxt.Text = "0"
    If IsNumeric(MusicTxt.Text) = False Then MusicTxt.Text = "0"
    If Val(MusicTxt.Text) > 255 Then MusicTxt.Text = "255"
    If Val(MusicTxt.Text) < 0 Then MusicTxt.Text = "0"
    MapInfo.Music = Val(MusicTxt.Text)

End Sub

Private Sub WeatherTxt_Change()

    If WeatherTxt.Text = "" Then WeatherTxt.Text = "0"
    If Val(WeatherTxt.Text) < 0 Then WeatherTxt.Text = "0"
    If Val(WeatherTxt.Text) > 255 Then WeatherTxt.Text = "255"
    MapInfo.Weather = WeatherTxt.Text

End Sub

Private Sub VersionTxt_Change()

    If VersionTxt.Text = "" Then VersionTxt.Text = "0"
    If IsNumeric(VersionTxt.Text) = False Then VersionTxt.Text = "0"
    If Val(VersionTxt.Text) > 32767 Then VersionTxt.Text = "32767"
    If Val(VersionTxt.Text) < 0 Then VersionTxt.Text = "0"
    MapInfo.MapVersion = Val(VersionTxt.Text)

End Sub
