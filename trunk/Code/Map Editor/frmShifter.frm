VERSION 5.00
Begin VB.Form frmShifter 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Shifter Tool"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   1605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ApplyCmd 
      Caption         =   "Apply Shift"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox ShiftYTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "Amount of tiles to shift the map on the Y axis (negative is up, positive is down)"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox ShiftXTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "Amount of tiles to shift the map on the X axis (negative is left, positive is right)"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shift Y:"
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
      TabIndex        =   1
      Top             =   480
      Width           =   645
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shift X:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "frmShifter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ApplyCmd_Click()
Dim OldSaveLight() As LightType
Dim OldMap() As MapBlock
Dim MapWidth As Long
Dim MapHeight As Long
Dim ShiftX As Long
Dim ShiftY As Long
Dim nX As Long
Dim nY As Long
Dim b As Byte
Dim X As Long
Dim Y As Long

    'Confirm
    If MsgBox("Are you sure you wish to apply the shifting?", vbYesNo) = vbNo Then Exit Sub
    
    'Store the map width/height
    MapWidth = MapInfo.Width
    MapHeight = MapInfo.Height
    
    'Check for valid values
    ShiftX = -Val(ShiftXTxt.Text)
    ShiftY = -Val(ShiftYTxt.Text)
    If ShiftX = 0 And ShiftY = 0 Then Exit Sub
    If ShiftX < 0 Then
        If ShiftX <= -MapWidth Then
            SetInfo "Invalid ShiftX value!", 1
            Exit Sub
        End If
        If ShiftX >= MapWidth Then
            SetInfo "Invalid ShiftX value!", 1
            Exit Sub
        End If
    End If
    If ShiftY < 0 Then
        If ShiftY <= -MapHeight Then
            SetInfo "Invalid ShiftY value!", 1
            Exit Sub
        End If
        If ShiftY >= MapHeight Then
            SetInfo "Invalid ShiftY value!", 1
            Exit Sub
        End If
    End If
    
    'Store the old map
    ReDim OldMap(1 To MapWidth, 1 To MapHeight)
    ReDim OldSaveLight(1 To MapWidth, 1 To MapHeight)
    For X = 1 To MapWidth
        For Y = 1 To MapHeight
            OldMap(X, Y) = MapData(X, Y)
            OldSaveLight(X, Y) = SaveLightBuffer(X, Y)
        Next Y
    Next X

    'Loop through all the tiles
    For X = 1 To MapWidth
        For Y = 1 To MapHeight
        
            b = 0
            
            'Check if the shift-to tile is in range
            nX = X + ShiftX
            nY = Y + ShiftY
            If nX >= 1 Then
                If nX <= MapWidth Then
                    If nY >= 1 Then
                        If nY <= MapHeight Then
                        
                            'Set the OldMap tile to the MapData tile
                            MapData(X, Y) = OldMap(nX, nY)
              
                            'Do the same with the light buffer
                            SaveLightBuffer(X, Y) = OldSaveLight(nX, nY)
                            
                            'We don't need to clear the tile
                            b = 1
                            
                        End If
                    End If
                End If
            End If
            
            'Clear the tile if it was not set
            If b = 0 Then
                ZeroMemory MapData(X, Y), Len(MapData(X, Y))
                ZeroMemory SaveLightBuffer(X, Y), Len(SaveLightBuffer(X, Y))
            End If
            
        Next Y
    Next X
    
    'Move the NPCs
    For X = 1 To LastChar
        If CharList(X).Active Then
            CharList(X).Pos.X = CharList(X).Pos.X - ShiftX
            CharList(X).Pos.Y = CharList(X).Pos.Y - ShiftY
        End If
    Next X
    
    'Move the particle effects
    For X = 1 To NumEffects
        If Effect(X).Used Then
            Effect(X).X = Effect(X).X - (ShiftX * 32)
            Effect(X).Y = Effect(X).Y - (ShiftY * 32)
        End If
    Next X
    
    'Refresh the map
    Engine_BuildMiniMap
    UpdateEffectList
    Engine_CreateTileLayers

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If IsUnloading = 0 Then Cancel = 1
    Me.Visible = False

End Sub

Private Sub ShiftYTxt_KeyPress(KeyAscii As Integer)

    If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 0

End Sub

Private Sub ShiftXTxt_KeyPress(KeyAscii As Integer)

    If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 0

End Sub
