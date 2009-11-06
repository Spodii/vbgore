VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "vbGORE Login"
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   196
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox PasswordTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "ggg"
      Top             =   720
      Width           =   1275
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Text            =   "ggg"
      Top             =   420
      Width           =   1275
   End
   Begin VB.CheckBox SavePassChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   915
      Width           =   1500
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    'Get the username/password
    NameTxt.Text = Engine_Var_Get(DataPath & "Game.ini", "INIT", "Name")
    PasswordTxt.Text = Engine_Var_Get(DataPath & "Game.ini", "INIT", "Password")
    
    'Get the background
    Me.Picture = LoadPicture(App.Path & "\Grh\Connect.bmp")

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*****************************************************************
'Process clicking events
'*****************************************************************
    
    'New
    If Engine_Collision_Rect(X, Y, 1, 1, 29, 85, 141, 36) Then
        frmNew.Visible = True
        frmNew.Show
        Me.Visible = False
    End If
    
    'Connect
    If Engine_Collision_Rect(X, Y, 1, 1, 29, 129, 141, 36) Then
        UserName = NameTxt.Text
        UserPassword = PasswordTxt.Text
        If Game_CheckUserData Then
            SendNewChar = False
            InitSocket
        End If
    End If
    
    'Exit
    If Engine_Collision_Rect(X, Y, 1, 1, 29, 174, 141, 36) Then
    
        'Save the game ini
        Engine_Var_Write DataPath & "Game.ini", "INIT", "Name", NameTxt.Text
        If SavePassChk.Value = 0 Then
            Engine_Var_Write DataPath & "Game.ini", "INIT", "Password", ""
        Else
            Engine_Var_Write DataPath & "Game.ini", "INIT", "Password", PasswordTxt.Text
        End If
    
        'End program
        IsUnloading = 1
        
    End If
    
End Sub

Private Sub NameTxt_Change()

    'Make sure the string is legal
    If Len(NameTxt.Text) > 0 Then
        If Game_LegalString(NameTxt.Text) = False Then
            NameTxt.Text = Left$(NameTxt.Text, Len(NameTxt.Text) - 1)
            NameTxt.SelStart = Len(NameTxt.Text)
        End If
    End If

End Sub

Private Sub PasswordTxt_Change()

    'Make sure the string is legal
    If Len(PasswordTxt.Text) > 0 Then
        If Game_LegalString(PasswordTxt.Text) = False Then
            PasswordTxt.Text = Left$(PasswordTxt.Text, Len(PasswordTxt.Text) - 1)
            PasswordTxt.SelStart = Len(PasswordTxt.Text)
        End If
    End If

End Sub
