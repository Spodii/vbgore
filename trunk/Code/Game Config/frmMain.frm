VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Game Configuration"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   477
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   381
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Private Function KeyName(ByVal KeyAscii As Integer) As String

'16 = Shift
'17 = Ctrl
'18 = Alt

    'Returns the key name
    Select Case KeyAscii
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
            KeyName = "F" & (KeyAscii - 111)
        Case 144
            KeyName = "(NUMLCK)"
        Case 145
            KeyName = "(SCRLLCK)"
        Case Else
            If KeyAscii >= 32 Then
                KeyName = UCase$(Chr$(KeyAscii))
            Else
                KeyName = "UNKNOWN"
            End If
    End Select
    
End Function

Private Sub Form_Load()
Dim i As Long

    'Clear the key cache
    For i = 1 To 255
        GetAsyncKeyState i
    Next i

End Sub

Private Sub Timer1_Timer()
Dim i As Long

    For i = 1 To 255
        If GetAsyncKeyState(i) Then
            Me.Caption = KeyName(i)
            Debug.Print i
        End If
    Next i
    
End Sub
