VERSION 5.00
Begin VB.Form frmMapInfo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Map Info"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   143
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   209
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton SizeCmd 
      Caption         =   "Apply Resize"
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox HeightTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox WidthTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox MusicTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "ID of the music file to be played in the map. 0 for nothing."
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox WeatherTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "What kind of weather goes on on the map - 0 for none"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox VersionTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "1"
      ToolTipText     =   "Version of the map. Used to determine if the client has the most up-to-date map."
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox MapNameTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "Name"
      ToolTipText     =   "Name of the map"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Height:"
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
      Height          =   195
      Index           =   5
      Left            =   1440
      TabIndex        =   9
      Top             =   1320
      Width           =   630
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width:"
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
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   570
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   840
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   1680
      TabIndex        =   6
      Top             =   480
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   480
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
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

    If IsUnloading = 0 Then Cancel = 1
    Me.Visible = False

End Sub

Private Sub HeightTxt_Change()

    SetInfo "New height of the map (in tiles)."

End Sub

Private Sub HeightTxt_KeyPress(KeyAscii As Integer)

    If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then KeyAscii = 0

End Sub

Private Sub MapNameTxt_Change()

    MapInfo.Name = MapNameTxt.Text

End Sub

Private Sub MapNameTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "The name of the map."

End Sub

Private Sub MusicTxt_Change()

    If MusicTxt.Text = "" Then MusicTxt.Text = "0"
    If IsNumeric(MusicTxt.Text) = False Then MusicTxt.Text = "0"
    If Val(MusicTxt.Text) > 255 Then MusicTxt.Text = "255"
    If Val(MusicTxt.Text) < 0 Then MusicTxt.Text = "0"
    MapInfo.Music = Val(MusicTxt.Text)

End Sub

Private Sub MusicTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub MusicTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "ID of the music file to be played in the map. 0 for nothing."

End Sub

Private Sub SizeCmd_Click()
Dim OldSaveLight() As LightType
Dim OldMap() As MapBlock
Dim OldWidth As Long
Dim OldHeight As Long
Dim X As Long
Dim Y As Long
Dim hX As Long
Dim hY As Long
Dim l As Long

    If MsgBox("If you resize your map down, anything cut off will be lost forever. Are you sure you wish to resize the map?", vbYesNo) = vbNo Then Exit Sub
    
    'Check for valid values
    Select Case Val(WidthTxt.Text)
        Case Is < 5, Is > 254
            SetInfo "Invalid width value defined! 254x254 is largest and 5x5 is smallest map sizes allowed.", 1
            Exit Sub
    End Select
    Select Case Val(HeightTxt.Text)
        Case Is < 5, Is > 254
            SetInfo "Invalid height value defined! 254x254 is largest and 5x5 is smallest map sizes allowed.", 1
            Exit Sub
    End Select
    
    'Resize
    'Goddamn VB won't let us use preserve, so we preserve it ourself!
    
    OldWidth = MapInfo.Width
    OldHeight = MapInfo.Height
    
    MapInfo.Width = Val(WidthTxt.Text)
    MapInfo.Height = Val(HeightTxt.Text)
    
    If OldWidth > MapInfo.Width Or OldHeight > MapInfo.Height Then
        For X = 1 To MapInfo.Width
            For Y = 1 To MapInfo.Height
                If X > MapInfo.Width Or Y > MapInfo.Height Then
                    If MapData(X, Y).NPCIndex > 0 Then
                        Engine_Char_Erase MapData(X, Y).NPCIndex
                    End If
                End If
            Next Y
        Next X
    End If
    
    ReDim OldMap(1 To OldWidth, 1 To OldHeight)
    ReDim OldSaveLight(1 To OldWidth, 1 To OldHeight)
    
    CopyMemory OldMap(1, 1), MapData(1, 1), Len(MapData(1, 1)) * OldWidth * OldHeight
    CopyMemory OldSaveLight(1, 1), SaveLightBuffer(1, 1), Len(SaveLightBuffer(1, 1)) * OldWidth * OldHeight
    
    ReDim MapData(1 To MapInfo.Width, 1 To MapInfo.Height)
    ReDim SaveLightBuffer(1 To MapInfo.Width, 1 To MapInfo.Height)

    If OldWidth < MapInfo.Width Then hX = OldWidth Else hX = MapInfo.Width
    If OldHeight < MapInfo.Height Then hY = OldHeight Else hY = MapInfo.Height
    
    For X = 1 To MapInfo.Width
        For Y = 1 To MapInfo.Height
            If X <= hX And Y <= hY Then
                MapData(X, Y) = OldMap(X, Y)
                CopyMemory SaveLightBuffer(X, Y), OldSaveLight(X, Y), Len(OldSaveLight(X, Y))
            Else
            
                'Default data (so they know the tiles are there)
                Engine_Init_Grh MapData(X, Y).Graphic(1), 2
                
                'Bad to keep lights as 0's
                For l = 1 To 24
                    MapData(X, Y).Light(l) = -1
                    SaveLightBuffer(X, Y).Light(l) = -1
                Next l
                
            End If
        Next Y
    Next X
    
    Engine_BuildMiniMap
    Engine_CreateTileLayers

End Sub

Private Sub SizeCmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Applies the map dimensions change."
    
End Sub

Private Sub VersionTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub VersionTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Version of the map. Used to determine if the client has the most up-to-date map."

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

Private Sub WeatherTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub WeatherTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "What kind of weather the map uses (0 for none)."

End Sub

Private Sub WidthTxt_KeyPress(KeyAscii As Integer)

    If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then KeyAscii = 0

End Sub

Private Sub WidthTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "New width of the map (in tiles)."

End Sub
