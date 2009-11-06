VERSION 5.00
Begin VB.Form frmConnect 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "vbGORE Login"
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   352
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox PasswordTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   3045
      MultiLine       =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "frmConnect.frx":17D2A
      Top             =   1320
      Width           =   1860
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3045
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmConnect.frx":17D2E
      Top             =   840
      Width           =   1860
   End
   Begin VB.Image SavePassImg 
      Height          =   180
      Left            =   3165
      Top             =   1605
      Width           =   180
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = Asc(vbNewLine) Then ClickConnect

End Sub

Private Sub Form_Load()

    'Set the text boxes to transparent
    SetPictureTextboxes Me.hwnd
    
    'Get the username/password
    NameTxt.Text = Var_Get(DataPath & "Game.ini", "INIT", "Name")
    PasswordTxt.Text = Var_Get(DataPath & "Game.ini", "INIT", "Password")
    SavePass = CBool(Val(Var_Get(DataPath & "Game.ini", "INIT", "SavePass")) * -1)
    
    'Set the SavePass image
    SavePass = Not SavePass 'Since the routine reverses, we reverse to reverse the reverse... trust me, it just works ;)
    SavePassImg_Click
    
    'Get the background
    Me.Picture = LoadPicture(App.Path & "\Grh\Connect.bmp")

End Sub

Private Sub ClickNew()

    'New character
    frmNew.Visible = True
    frmNew.Show
    Me.Visible = False
    
End Sub

Private Sub ClickConnect()

    'Connect
    UserName = NameTxt.Text
    UserPassword = PasswordTxt.Text
    If Game_CheckUserData Then
        SendNewChar = False
        InitSocket
    End If

End Sub

Private Sub ClickExit()

    'Save the game ini
    Var_Write DataPath & "Game.ini", "INIT", "Name", NameTxt.Text
    Var_Write DataPath & "Game.ini", "INIT", "SavePass", -CInt(SavePass)
    If Not SavePass Then
        Var_Write DataPath & "Game.ini", "INIT", "Password", ""
    Else
        Var_Write DataPath & "Game.ini", "INIT", "Password", PasswordTxt.Text
    End If

    'End program
    IsUnloading = 1
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*****************************************************************
'Process clicking events
'*****************************************************************
    
    'New
    If Engine_Collision_Rect(X, Y, 1, 1, 217, 149, 96, 18) Then ClickNew
    
    'Connect
    If Engine_Collision_Rect(X, Y, 1, 1, 217, 127, 96, 18) Then ClickConnect

    'Exit
    If Engine_Collision_Rect(X, Y, 1, 1, 217, 171, 96, 18) Then ClickExit
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Free the form
    FreePictureTextboxes Me.hwnd

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

Private Sub NameTxt_KeyPress(KeyAscii As Integer)

    'Because we have to use multiline to have the image set on the background, cancel new lines
    If KeyAscii = Asc(vbNewLine) Then
        KeyAscii = 0
        ClickConnect
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

Private Sub PasswordTxt_KeyPress(KeyAscii As Integer)

    'Because we have to use multiline to have the image set on the background, cancel new lines
    If KeyAscii = Asc(vbNewLine) Then
        KeyAscii = 0
        ClickConnect
    End If

End Sub

Private Sub SavePassImg_Click()

    'Change the value
    SavePass = Not SavePass
    
    'Display the image or remove it
    If SavePass Then
        SavePassImg.Picture = LoadPicture(GrhPath & "Check.gif")
    Else
        Set SavePassImg.Picture = Nothing
    End If

End Sub
