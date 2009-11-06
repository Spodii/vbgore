VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "Game Configuration"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6120
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   467
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton DefaultsCmd 
      Caption         =   "Default Controls"
      Height          =   255
      Left            =   1560
      TabIndex        =   62
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton CloseCmd 
      Caption         =   "Close"
      Height          =   255
      Left            =   4440
      TabIndex        =   61
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton SaveCmd 
      Caption         =   "Save"
      Height          =   255
      Left            =   3000
      TabIndex        =   60
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton LoadCmd 
      Caption         =   "Load"
      Height          =   255
      Left            =   120
      TabIndex        =   59
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000005&
      Caption         =   "Quick Bar Hot-Keys"
      Height          =   1935
      Left            =   120
      TabIndex        =   46
      Top             =   2880
      Width           =   5895
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   27
         Left            =   4320
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   26
         Left            =   2400
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   25
         Left            =   480
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   24
         Left            =   4320
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   23
         Left            =   2400
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   22
         Left            =   480
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   21
         Left            =   4320
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   20
         Left            =   2400
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   19
         Left            =   480
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   18
         Left            =   4320
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   17
         Left            =   2400
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   16
         Left            =   480
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "12:"
         Height          =   195
         Left            =   3960
         TabIndex        =   58
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "11:"
         Height          =   195
         Left            =   2040
         TabIndex        =   57
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "10:"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "9:"
         Height          =   195
         Left            =   3960
         TabIndex        =   55
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "8:"
         Height          =   195
         Left            =   2040
         TabIndex        =   54
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "7:"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "6:"
         Height          =   195
         Left            =   3960
         TabIndex        =   52
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "5:"
         Height          =   195
         Left            =   2040
         TabIndex        =   51
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "4:"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "3:"
         Height          =   195
         Left            =   3960
         TabIndex        =   49
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "2:"
         Height          =   195
         Left            =   2040
         TabIndex        =   48
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "1:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Caption         =   "Window Controls"
      Height          =   1575
      Left            =   120
      TabIndex        =   39
      Top             =   4920
      Width           =   5895
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   14
         Left            =   1200
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   13
         Left            =   4200
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   12
         Left            =   1200
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   11
         Left            =   4200
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   10
         Left            =   1200
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Stats:"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Quick bar:"
         Height          =   195
         Left            =   2880
         TabIndex        =   43
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Menu:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Inventory:"
         Height          =   195
         Left            =   2880
         TabIndex        =   41
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Chat:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "General Controls"
      Height          =   2655
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   5895
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   28
         Left            =   4200
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   15
         Left            =   4200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   9
         Left            =   4200
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   8
         Left            =   4200
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   7
         Left            =   4200
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   6
         Left            =   4200
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         BackColor       =   &H8000000E&
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Reset GUI:"
         Height          =   195
         Left            =   2880
         TabIndex        =   64
         Top             =   2160
         Width           =   1230
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Mini-map:"
         Height          =   195
         Left            =   2880
         TabIndex        =   45
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Zoom out:"
         Height          =   195
         Left            =   2880
         TabIndex        =   38
         Top             =   1800
         Width           =   1230
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Zoom in:"
         Height          =   195
         Left            =   2880
         TabIndex        =   37
         Top             =   1440
         Width           =   1230
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Scroll chat down:"
         Height          =   195
         Left            =   2880
         TabIndex        =   36
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Scroll chat up:"
         Height          =   195
         Left            =   2880
         TabIndex        =   35
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Move right:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Move left:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Move down:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Move up:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Pick up:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Attack:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const KeyPress_Shift As Integer = 2 ^ 12
Private Const KeyPress_Control As Integer = 2 ^ 13
Private Const KeyPress_Alt As Integer = 2 ^ 14

Private Type KeyDefinitions
    MiniMap As Integer
    PickUpObj As Integer
    QuickBar(1 To 12) As Integer
    Attack As Integer
    ChatBufferUp As Integer
    ChatBufferDown As Integer
    InventoryWindow As Integer
    QuickBarWindow As Integer
    ChatWindow As Integer
    StatWindow As Integer
    MenuWindow As Integer
    ZoomIn As Integer
    ZoomOut As Integer
    MoveNorth As Integer
    MoveEast As Integer
    MoveSouth As Integer
    MoveWest As Integer
    ResetGUI As Integer
End Type
Private KeyDefinitions As KeyDefinitions

Private HasChanged As Boolean

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Private Function KeyName(ByVal KeyCode As Integer) As String
Dim s As String

    'Check for shift, alt and control
    If KeyCode And KeyPress_Shift Then s = "(SHIFT)": KeyCode = KeyCode Xor KeyPress_Shift
    If KeyCode And KeyPress_Control Then s = "(CTRL)": KeyCode = KeyCode Xor KeyPress_Control
    If KeyCode And KeyPress_Alt Then s = "(ALT)": KeyCode = KeyCode Xor KeyPress_Alt
    
    'Remove the Shift, Control and Alt bits
    KeyCode = KeyCode And 2047

    'Check for known names
    Select Case KeyCode
        
        Case 1          'Left-click
        Case 2          'Right-click
        Case 3          'Cancel
        Case 4          'Middle-click
        
        Case 16, 160
            KeyName = "(SHIFT)"
        Case 17, 162
            KeyName = "(CTRL)"
        Case 18, 164
            KeyName = "(ALT)"
            
        Case 8
            KeyName = "(BACK)"
        Case 9
            KeyName = "(TAB)"
        Case 12
            KeyName = "(CLEAR)"
        Case 13
            KeyName = "(RETURN)"
        Case 19
            KeyName = "(PAUSE)"
        Case 20
            KeyName = "(CAP)"
        Case 27
            KeyName = "(ESC)"
        Case 32
            KeyName = "(SPACE)"
        Case 33
            KeyName = "(PGUP)"
        Case 34
            KeyName = "(PGDOWN)"
        Case 35
            KeyName = "(END)"
        Case 36
            KeyName = "(HOME)"
        Case 37
            KeyName = "(LEFT)"
        Case 38
            KeyName = "(UP)"
        Case 39
            KeyName = "(RIGHT)"
        Case 40
            KeyName = "(DOWN)"
        Case 41
            KeyName = "(SELECT)"
        Case 42
            KeyName = "(PRINT)"
        Case 43
            KeyName = "(EXECUTE)"
        Case 44
            KeyName = "(SNAPSHOT)"
        Case 45
            KeyName = "(INS)"
        Case 46
            KeyName = "(DEL)"
        Case 47
            KeyName = "(HELP)"
        Case 112 To 127
            KeyName = "F" & (KeyCode - 111)
        Case 144
            KeyName = "(NUMLCK)"
        Case 145
            KeyName = "(SCRLLCK)"
        Case Else
            If KeyCode >= 32 Then
                KeyName = UCase$(Chr$(KeyCode))
            Else
                KeyName = "(UNKNOWN)"
            End If
    End Select
    
    If s <> vbNullString Then
        KeyName = s & " + " & KeyName
    End If
    
End Function

Private Function GetKeyValue(ByVal KeyCode As Integer) As Integer

    'Only add on Shift, Control or Alt combos if they aren't pressed
    If KeyCode <> 16 Then
        If KeyCode <> 17 Then
            If KeyCode <> 18 Then
                If GetAsyncKeyState(16) Then GetKeyValue = GetKeyValue Or KeyPress_Shift
                If GetAsyncKeyState(17) Then GetKeyValue = GetKeyValue Or KeyPress_Control
                If GetAsyncKeyState(18) Then GetKeyValue = GetKeyValue Or KeyPress_Alt
            End If
        End If
    End If
    
    'Add on the keycode
    GetKeyValue = GetKeyValue Or KeyCode
    
    'Clear the previous alt/control/shift key presses
    GetAsyncKeyState 16
    GetAsyncKeyState 17
    GetAsyncKeyState 18

End Function

Private Sub CloseCmd_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'Set the file paths
    InitFilePaths

    'Clear the key cache
    Input_Keys_ClearQueue
    
    'Load the key config
    Input_KeyDefinitions_Load

End Sub

Private Sub Input_Keys_ClearQueue()

'*****************************************************************
'Clears the GetAsyncKeyState queue to prevent key presses from a long time
' ago falling into "have been pressed"
'*****************************************************************
Dim i As Long

    For i = 0 To 255
        GetAsyncKeyState i
    Next i

End Sub

Private Sub Input_KeyDefinitions_Load()

'*****************************************************************
'Load the key definitions
'*****************************************************************
Dim i As Long

    KeyDefinitions.Attack = Val(Var_Get(DataPath & "Game.ini", "INPUT", "Attack"))
    KeyDefinitions.ChatBufferDown = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ChatBufferDown"))
    KeyDefinitions.ChatBufferUp = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ChatBufferUp"))
    KeyDefinitions.ChatWindow = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ChatWindow"))
    KeyDefinitions.InventoryWindow = Val(Var_Get(DataPath & "Game.ini", "INPUT", "InventoryWindow"))
    KeyDefinitions.MenuWindow = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MenuWindow"))
    KeyDefinitions.MiniMap = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MiniMap"))
    KeyDefinitions.MoveEast = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MoveEast"))
    KeyDefinitions.MoveNorth = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MoveNorth"))
    KeyDefinitions.MoveSouth = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MoveSouth"))
    KeyDefinitions.MoveWest = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MoveWest"))
    KeyDefinitions.PickUpObj = Val(Var_Get(DataPath & "Game.ini", "INPUT", "PickUpObj"))
    KeyDefinitions.QuickBarWindow = Val(Var_Get(DataPath & "Game.ini", "INPUT", "QuickBarWindow"))
    KeyDefinitions.StatWindow = Val(Var_Get(DataPath & "Game.ini", "INPUT", "StatWindow"))
    KeyDefinitions.ZoomIn = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ZoomIn"))
    KeyDefinitions.ZoomOut = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ZoomOut"))
    KeyDefinitions.ResetGUI = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ResetGUI"))
    For i = 1 To 12
        KeyDefinitions.QuickBar(i) = Val(Var_Get(DataPath & "Game.ini", "INPUT", "QuickBar" & i))
    Next i
    
    'Only used in the config editor
    SetTextBoxes
    
End Sub

Private Sub Input_KeyDefinitions_Save()

'*****************************************************************
'Save the key definitions
'*****************************************************************
Dim i As Long

    Var_Write DataPath & "Game.ini", "INPUT", "Attack", KeyDefinitions.Attack
    Var_Write DataPath & "Game.ini", "INPUT", "ChatBufferDown", KeyDefinitions.ChatBufferDown
    Var_Write DataPath & "Game.ini", "INPUT", "ChatBufferUp", KeyDefinitions.ChatBufferUp
    Var_Write DataPath & "Game.ini", "INPUT", "ChatWindow", KeyDefinitions.ChatWindow
    Var_Write DataPath & "Game.ini", "INPUT", "InventoryWindow", KeyDefinitions.InventoryWindow
    Var_Write DataPath & "Game.ini", "INPUT", "MenuWindow", KeyDefinitions.MenuWindow
    Var_Write DataPath & "Game.ini", "INPUT", "MiniMap", KeyDefinitions.MiniMap
    Var_Write DataPath & "Game.ini", "INPUT", "MoveEast", KeyDefinitions.MoveEast
    Var_Write DataPath & "Game.ini", "INPUT", "MoveNorth", KeyDefinitions.MoveNorth
    Var_Write DataPath & "Game.ini", "INPUT", "MoveSouth", KeyDefinitions.MoveSouth
    Var_Write DataPath & "Game.ini", "INPUT", "MoveWest", KeyDefinitions.MoveWest
    Var_Write DataPath & "Game.ini", "INPUT", "PickUpObj", KeyDefinitions.PickUpObj
    Var_Write DataPath & "Game.ini", "INPUT", "QuickBarWindow", KeyDefinitions.QuickBarWindow
    Var_Write DataPath & "Game.ini", "INPUT", "StatWindow", KeyDefinitions.StatWindow
    Var_Write DataPath & "Game.ini", "INPUT", "ZoomIn", KeyDefinitions.ZoomIn
    Var_Write DataPath & "Game.ini", "INPUT", "ZoomOut", KeyDefinitions.ZoomOut
    Var_Write DataPath & "Game.ini", "INPUT", "ResetGUI", KeyDefinitions.ResetGUI
    For i = 1 To 12
        Var_Write DataPath & "Game.ini", "INPUT", "QuickBar" & i, KeyDefinitions.QuickBar(i)
    Next i

End Sub

Private Function Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
Dim sSpaces As String

    sSpaces = Space$(1000)
    GetPrivateProfileString Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    Var_Get = RTrim$(sSpaces)
    If Len(Var_Get) > 0 Then
        Var_Get = Left$(Var_Get, Len(Var_Get) - 1)
    Else
        Var_Get = vbNullString
    End If
    
End Function

Private Sub Var_Write(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    WritePrivateProfileString Main, Var, Value, File

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If HasChanged Then
        If MsgBox("Are you sure you wish to quit? Any unsaved changes will be lost!", vbYesNo) = vbNo Then
            Cancel = 1
            UnloadMode = 0
            Exit Sub
        End If
    End If
    
End Sub

Private Sub KeyTxt_GotFocus(Index As Integer)

    'Set the high-light
    KeyTxt(Index).BackColor = &H80000013

End Sub

Private Sub SetTextBoxes()
Dim i As Long

    'Set the values in the text boxes
    With KeyDefinitions
        KeyTxt(0).Text = KeyName(.Attack)
        KeyTxt(1).Text = KeyName(.PickUpObj)
        KeyTxt(2).Text = KeyName(.MoveNorth)
        KeyTxt(3).Text = KeyName(.MoveSouth)
        KeyTxt(4).Text = KeyName(.MoveWest)
        KeyTxt(5).Text = KeyName(.MoveEast)
        KeyTxt(6).Text = KeyName(.ChatBufferUp)
        KeyTxt(7).Text = KeyName(.ChatBufferDown)
        KeyTxt(8).Text = KeyName(.ZoomIn)
        KeyTxt(9).Text = KeyName(.ZoomOut)
        KeyTxt(10).Text = KeyName(.ChatWindow)
        KeyTxt(11).Text = KeyName(.InventoryWindow)
        KeyTxt(12).Text = KeyName(.MenuWindow)
        KeyTxt(13).Text = KeyName(.QuickBarWindow)
        KeyTxt(14).Text = KeyName(.StatWindow)
        KeyTxt(15).Text = KeyName(.MiniMap)
        For i = 1 To 12
            KeyTxt(15 + i).Text = KeyName(.QuickBar(i))
        Next i
        KeyTxt(28).Text = KeyName(.ResetGUI)
    End With

End Sub

Private Sub DefaultsCmd_Click()
Dim i As Long

    If MsgBox("Are you sure you wish to restore the default control settings?" & vbNewLine & "Any unsaved changes will be lost!", vbYesNo) = vbNo Then Exit Sub

    'Set to the default settings used
    With KeyDefinitions
        .Attack = 17
        .PickUpObj = 18
        .MoveNorth = 87
        .MoveEast = 68
        .MoveSouth = 83
        .MoveWest = 65
        .ChatBufferUp = 33
        .ChatBufferDown = 34
        .ZoomIn = 104
        .ZoomOut = 98
        .ChatWindow = KeyPress_Control Or 67
        .InventoryWindow = KeyPress_Control Or 69
        .MenuWindow = 27
        .QuickBarWindow = KeyPress_Control Or 81
        .StatWindow = KeyPress_Control Or 83
        .MiniMap = 9
        For i = 1 To 12
            .QuickBar(i) = 111 + i
        Next i
        .ResetGUI = KeyPress_Shift Or 123
    End With
    
    'Display the changes
    SetTextBoxes
    HasChanged = False

End Sub

Private Sub KeyTxt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Long
        
    'Get the value
    i = GetKeyValue(KeyCode)
    
    'Display the value
    KeyTxt(Index).Text = KeyName(i)
    
    'Set the value to the appropriate variable
    With KeyDefinitions
        Select Case Index
            Case 0: .Attack = i
            Case 1: .PickUpObj = i
            Case 2: .MoveNorth = i
            Case 3: .MoveSouth = i
            Case 4: .MoveWest = i
            Case 5: .MoveEast = i
            Case 6: .ChatBufferUp = i
            Case 7: .ChatBufferDown = i
            Case 8: .ZoomIn = i
            Case 9: .ZoomOut = i
            Case 10: .ChatWindow = i
            Case 11: .InventoryWindow = i
            Case 12: .MenuWindow = i
            Case 13: .QuickBarWindow = i
            Case 14: .StatWindow = i
            Case 15: .MiniMap = i
            Case 16 To 27: .QuickBar(Index - 15) = i
            Case 28: .ResetGUI = i
        End Select
    End With
    
    'Clear the key so no text will be entered in the control
    KeyCode = 0
    Shift = 0
    
    'A change has been made
    HasChanged = True
    
End Sub

Private Sub KeyTxt_KeyPress(Index As Integer, KeyAscii As Integer)

    'Clear the key so no text will be entered in the control
    KeyAscii = 0

End Sub

Private Sub KeyTxt_LostFocus(Index As Integer)
    
    'Remove the high-light
    KeyTxt(Index).BackColor = &H80000005

End Sub

Private Sub LoadCmd_Click()

    If MsgBox("Are you sure you wish to load the last saved settings?", vbYesNo) = vbNo Then Exit Sub
    Input_KeyDefinitions_Load
    HasChanged = False

End Sub

Private Sub SaveCmd_Click()

    If MsgBox("Are you sure you wish to save the current settings?", vbYesNo) = vbNo Then Exit Sub
    Input_KeyDefinitions_Save
    HasChanged = False

End Sub
