VERSION 5.00
Begin VB.Form frmOptimizeStart 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Optimize Map"
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   124
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   184
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MapEditor.cButton OptBtn 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      Caption         =   "Begin Optimization Check"
   End
   Begin MapEditor.cForm cForm 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      MaximizeBtn     =   0   'False
      MinimizeBtn     =   0   'False
      Caption         =   "Map Optimizations"
      CaptionTop      =   0
      AllowResizing   =   0   'False
   End
   Begin VB.CheckBox DuplicateGrhChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Duplicate Grh Layers"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Check for the same graphic placed on multiple layers"
      Top             =   600
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox BlockedNPCChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "NPCs On Blocked Tiles"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      ToolTipText     =   "Check for NPCs placed on tiles that are blocked"
      Top             =   1080
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox BlockedObjChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Objects On Blocked Tiles"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Check for objects placed on tiles that are blocked"
      Top             =   840
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox EmptyLightsChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Lighted Empty Layers"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "Check for lights placed on layers that do not have a graphic"
      Top             =   360
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Optimization Checks:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "frmOptimizeStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddToOpt(ByVal OptType As MapOptType, ByVal tX As Byte, ByVal tY As Byte, Optional ByVal Layer As Byte = 0, Optional ByVal Layer2 As Byte = 0)
Dim Index As Long

    'Add an item to the optimization list
    Index = UBound(MapOpt) + 1
    ReDim Preserve MapOpt(0 To Index)
    MapOpt(Index).tX = tX
    MapOpt(Index).tY = tY
    MapOpt(Index).Layer = Layer
    MapOpt(Index).Layer2 = Layer2
    MapOpt(Index).Type = OptType

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

Private Sub Form_Load()

    cForm.LoadSkin Me
    Skin_Set Me
    Me.Refresh
    
End Sub

Private Sub OptBtn_Click()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim j As Long

    'Set up the list
    ReDim MapOpt(0)

    'Check for objects on blocked tiles
    If BlockedObjChk.Value = 1 Then
        OptBtn.Caption = "Checking Obj tiles..."
        OptBtn.Refresh
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                If MapData(X, Y).ObjInfo.ObjIndex > 0 Then
                    'Check if the obj is in the border
                    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
                        AddToOpt ObjOnBlocked, X, Y
                    'Check if on a blocked tile
                    Else
                        If MapData(X, Y).Blocked Then AddToOpt ObjOnBlocked, X, Y
                    End If
                End If
            Next Y
        Next X
    End If
    
    'Check for NPCs on blocked tiles
    If BlockedNPCChk.Value = 1 Then
        OptBtn.Caption = "Checking NPC tiles..."
        OptBtn.Refresh
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                If MapData(X, Y).NPCIndex > 0 Then
                    'Check if the NPC is in the border
                    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
                        AddToOpt NPCOnBlocked, X, Y
                    'Check if on a blocked tile
                    Else
                        If MapData(X, Y).Blocked Then AddToOpt NPCOnBlocked, X, Y
                    End If
                End If
            Next Y
        Next X
    End If
    
    'Check for empty lights
    If EmptyLightsChk.Value = 1 Then
        OptBtn.Caption = "Checking lights..."
        OptBtn.Refresh
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                For i = 2 To 6  'Loop through layers
                    If MapData(X, Y).Graphic(i).GrhIndex = 0 Then   'Check for empty layer
                        'Check the 4 lights that correspond to the layer
                        If MapData(X, Y).Light((i - 1) * 4 + 1) <> -1 Then
                            AddToOpt EmptyLight, X, Y, i
                        ElseIf MapData(X, Y).Light((i - 1) * 4 + 2) <> -1 Then
                            AddToOpt EmptyLight, X, Y, i
                        ElseIf MapData(X, Y).Light((i - 1) * 4 + 3) <> -1 Then
                            AddToOpt EmptyLight, X, Y, i
                        ElseIf MapData(X, Y).Light((i - 1) * 4 + 4) <> -1 Then
                            AddToOpt EmptyLight, X, Y, i
                        End If
                    End If
                Next i
            Next Y
        Next X
    End If
    
    'Check for duplicate grh layers (same tile with the same grh on two or more layers)
    If EmptyLightsChk.Value = 1 Then
        OptBtn.Caption = "Checking grh layers..."
        OptBtn.Refresh
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                For i = 1 To 6  'Loop through base layers
                    If MapData(X, Y).Graphic(i).GrhIndex > 0 Then   'We dont care if the grh = 0
                        For j = i + 1 To 6  'Loop through comparison layers
                            If MapData(X, Y).Graphic(i).GrhIndex = MapData(X, Y).Graphic(j).GrhIndex Then
                                AddToOpt DuplicateGrhLayers, X, Y, i, j
                            End If
                        Next j
                    End If
                Next i
            Next Y
        Next X
    End If
    
    'Restore the label
    OptBtn.Caption = "Begin Optimization Check"
    OptBtn.Refresh
    
    'Show report
    HideFrmOptimizeStart
    ShowFrmReport

End Sub
