VERSION 5.00
Object = "{CFFD2C69-2D50-4C10-A50C-89488104482B}#1.0#0"; "GOREsockClient.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "vbGORE Client"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin GOREsock.GOREsockClient GOREsock 
      Left            =   1080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer PTDTmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer ShutdownTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements DirectXEvent8

Private NC As Byte
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long

Private Sub DirectXEvent8_DXCallback(ByVal EventID As Long)
Dim DevData(1 To 50) As DIDEVICEOBJECTDATA
Dim NumEvents As Long
Dim LoopC As Long
Dim Moved As Byte
Dim OldMousePos As POINTAPI

    On Error GoTo ErrOut

    'Check if message is for us
    If EventID <> MouseEvent Then Exit Sub
    If GetActiveWindow = 0 Then Exit Sub

    'Retrieve data
    NumEvents = DIDevice.GetDeviceData(DevData, DIGDD_DEFAULT)

    'Loop through data
    For LoopC = 1 To NumEvents
        Select Case DevData(LoopC).lOfs

            'Move on X axis
        Case DIMOFS_X
            If Windowed Then
                OldMousePos = MousePos
                GetCursorPos MousePos
                MousePos.X = MousePos.X - (Me.Left \ Screen.TwipsPerPixelX)
                MousePos.Y = MousePos.Y - (Me.Top \ Screen.TwipsPerPixelY)
                MousePosAdd.X = -(OldMousePos.X - MousePos.X)
                MousePosAdd.Y = -(OldMousePos.Y - MousePos.Y)
            Else
                MousePosAdd.X = (DevData(LoopC).lData * MouseSpeed)
                MousePos.X = MousePos.X + MousePosAdd.X
                If MousePos.X < 0 Then MousePos.X = 0
                If MousePos.X > frmMain.ScaleWidth Then MousePos.X = frmMain.ScaleWidth
            End If
            Moved = 1

            'Move on Y axis
        Case DIMOFS_Y
            If Windowed Then
                OldMousePos = MousePos
                GetCursorPos MousePos
                MousePos.X = MousePos.X - (Me.Left \ Screen.TwipsPerPixelX)
                MousePos.Y = MousePos.Y - (Me.Top \ Screen.TwipsPerPixelY)
                MousePosAdd.X = -(OldMousePos.X - MousePos.X)
                MousePosAdd.Y = -(OldMousePos.Y - MousePos.Y)
            Else
                MousePosAdd.Y = (DevData(LoopC).lData * MouseSpeed)
                MousePos.Y = MousePos.Y + MousePosAdd.Y
                If MousePos.Y < 0 Then MousePos.Y = 0
                If MousePos.Y > frmMain.ScaleHeight Then MousePos.Y = frmMain.ScaleHeight
            End If
            Moved = 1

            'Left button pressed
        Case DIMOFS_BUTTON0
            If DevData(LoopC).lData = 0 Then
                MouseLeftDown = 0
                SelGameWindow = 0
            Else
                If MouseLeftDown = 0 Then   'Clicked down
                    MouseLeftDown = 1
                    Input_Mouse_LeftClick
                End If
            End If

            'Right button pressed
        Case DIMOFS_BUTTON1
            If DevData(LoopC).lData = 0 Then
                MouseRightDown = 0
                Input_Mouse_RightRelease
            Else
                If MouseRightDown = 0 Then  'Clicked down
                    MouseRightDown = 1
                    Input_Mouse_RightClick
                End If
            End If
        
            'Mouse wheel is scrolled
        Case DIMOFS_Z
            If ShowGameWindow(ChatWindow) And Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, GameWindow.ChatWindow.Screen.X, GameWindow.ChatWindow.Screen.Y, GameWindow.ChatWindow.Screen.Width, GameWindow.ChatWindow.Screen.Height) Then
                If DevData(LoopC).lData > 0 Then
                    ChatBufferChunk = ChatBufferChunk + 0.25
                ElseIf DevData(LoopC).lData < 0 Then
                    ChatBufferChunk = ChatBufferChunk - 0.25
                End If
                Engine_UpdateChatArray
            Else
                If DevData(LoopC).lData > 0 Then
                    ZoomLevel = ZoomLevel + (ElapsedTime * 0.001)
                    If ZoomLevel > MaxZoomLevel Then ZoomLevel = MaxZoomLevel
                ElseIf DevData(LoopC).lData < 0 Then
                    ZoomLevel = ZoomLevel - (ElapsedTime * 0.001)
                    If ZoomLevel < 0 Then ZoomLevel = 0
                End If
            End If
        End Select

        'Update movement
        If Moved Then
            Input_Mouse_Move

            'Reset move variables
            Moved = 0
            MousePosAdd.X = 0
            MousePosAdd.Y = 0
        End If

    Next LoopC

Exit Sub

ErrOut:
    NC = 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Input_Keys_Down KeyCode, Shift
    KeyCode = 0
    Shift = 0

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Input_Keys_Press KeyAscii
    KeyAscii = 0

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Input_Keys_Up KeyCode, Shift
    KeyCode = 0
    Shift = 0

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Regain focus to Direct Input mouse
    If NC Then
        DIDevice.Acquire
        NC = 0
        MousePos.X = X
        MousePos.Y = Y
    End If
    
End Sub

Private Sub Form_Resize()

    If Not DIDevice Is Nothing Then
        If Windowed = False Then DIDevice.Acquire
    End If
    
End Sub

Private Sub PTDTmr_Timer()

    sndBuf.Put_Byte DataCode.Server_PTD
    PTDSTime = timeGetTime

End Sub

Private Sub ShutdownTimer_Timer()
Static FailedUnloads As Long

    On Error Resume Next    'Who cares about an error if we are closing down

    'Quit the client - we must user a timer since DoEvents wont work (since we're not multithreaded)
    
    'Close down the socket
    If FailedUnloads > 5 Or frmMain.GOREsock.ShutDown <> soxERROR Then
        frmMain.GOREsock.UnHook

        'Unload the engine
        Engine_Init_UnloadTileEngine
        
        'Unload the forms
        Engine_UnloadAllForms
        
        'Unload everything else
        End

    Else
        
        'If the socket is making an error on the shutdown sequence for more than a second, just unload anyways
        FailedUnloads = FailedUnloads + 1
        
    End If
            

End Sub

Private Sub GOREsock_OnDataArrival(inSox As Long, inData() As Byte)

'*********************************************
'Retrieve the CommandIDs and send to corresponding data handler
'*********************************************

Dim rBuf As DataBuffer
Dim CommandID As Byte
Dim BufUBound As Long
Static X As Long

    'Display packet
    If DEBUG_PrintPacket_In Then
        Engine_AddToChatTextBuffer "DataIn: " & StrConv(inData, vbUnicode), -1
    End If

    'Set up the data buffer
    Set rBuf = New DataBuffer
    rBuf.Set_Buffer inData
    BufUBound = UBound(inData)
    
    'Uncomment this to see packets going into the client
    'Dim i As Long
    'Dim s As String
    'For i = LBound(inData) To UBound(inData)
    '    If inData(i) >= 100 Then
    '        s = s & inData(i) & " "
    '    ElseIf inData(i) >= 10 Then
    '        s = s & "0" & inData(i) & " "
    '    Else
    '        s = s & "00" & inData(i) & " "
    '    End If
    'Next i
    'Debug.Print s

    Do
        'Get the Command ID
        CommandID = rBuf.Get_Byte

        'Make the appropriate call based on the Command ID
        With DataCode
            Select Case CommandID

            Case 0
                If DEBUG_PrintPacketReadErrors Then
                    X = X + 1
                    Debug.Print "Empty Command ID #" & X
                End If

            Case .Comm_Talk: Data_Comm_Talk rBuf

            Case .Map_LoadMap: Data_Map_LoadMap rBuf
            Case .Map_SendName:  Data_Map_SendName rBuf

            Case .Server_ChangeChar: Data_Server_ChangeChar rBuf
            Case .Server_ChangeCharType: Data_Server_ChangeCharType rBuf
            Case .Server_CharHP: Data_Server_CharHP rBuf
            Case .Server_CharMP: Data_Server_CharMP rBuf
            Case .Server_Connect: Data_Server_Connect
            Case .Server_Disconnect: Data_Server_Disconnect
            Case .Server_EraseChar: Data_Server_EraseChar rBuf
            Case .Server_EraseObject: Data_Server_EraseObject rBuf
            Case .Server_IconBlessed: Data_Server_IconBlessed rBuf
            Case .Server_IconCursed: Data_Server_IconCursed rBuf
            Case .Server_IconIronSkin: Data_Server_IconIronSkin rBuf
            Case .Server_IconProtected: Data_Server_IconProtected rBuf
            Case .Server_IconStrengthened: Data_Server_IconStrengthened rBuf
            Case .Server_IconWarCursed:  Data_Server_IconWarCursed rBuf
            Case .Server_IconSpellExhaustion: Data_Server_IconSpellExhaustion rBuf
            Case .Server_MailBox: Data_Server_Mailbox rBuf
            Case .Server_MailItemRemove: Data_Server_MailItemRemove rBuf
            Case .Server_MailMessage: Data_Server_MailMessage rBuf
            Case .Server_MailObjUpdate: Data_Server_MailObjUpdate rBuf
            Case .Server_MakeChar: Data_Server_MakeChar rBuf
            Case .Server_MakeEffect: Data_Server_MakeEffect rBuf
            Case .Server_MakeSlash: Data_Server_MakeSlash rBuf
            Case .Server_MakeObject: Data_Server_MakeObject rBuf
            Case .Server_MakeProjectile: Data_Server_MakeProjectile rBuf
            Case .Server_Message: Data_Server_Message rBuf
            Case .Server_MoveChar: Data_Server_MoveChar rBuf
            Case .Server_PTD: Data_Server_PTD
            Case .Server_PlaySound: Data_Server_PlaySound rBuf
            Case .Server_PlaySound3D: Data_Server_PlaySound3D rBuf
            Case .Server_SendQuestInfo: Data_Server_SendQuestInfo rBuf
            Case .Server_SetCharDamage: Data_Server_SetCharDamage rBuf
            Case .Server_SetCharSpeed: Data_Server_SetCharSpeed rBuf
            Case .Server_SetUserPosition: Data_Server_SetUserPosition rBuf
            Case .Server_UserCharIndex: Data_Server_UserCharIndex rBuf

            Case .User_Attack: Data_User_Attack rBuf
            Case .User_Bank_Open: Data_User_Bank_Open rBuf
            Case .User_Bank_UpdateSlot: Data_User_Bank_UpdateSlot rBuf
            Case .User_BaseStat: Data_User_BaseStat rBuf
            Case .User_Blink: Data_User_Blink rBuf
            Case .User_CastSkill: Data_User_CastSkill rBuf
            Case .User_ChangeServer: Data_User_ChangeServer rBuf
            Case .User_Emote: Data_User_Emote rBuf
            Case .User_KnownSkills: Data_User_KnownSkills rBuf
            Case .User_LookLeft: Data_User_LookLeft rBuf
            Case .User_LookRight: Data_User_LookLeft rBuf
            Case .User_ModStat: Data_User_ModStat rBuf
            Case .User_Rotate: Data_User_Rotate rBuf
            Case .User_SetInventorySlot: Data_User_SetInventorySlot rBuf
            Case .User_SetWeaponRange: Data_User_SetWeaponRange rBuf
            Case .User_Target: Data_User_Target rBuf
            Case .User_Trade_StartNPCTrade: Data_User_Trade_StartNPCTrade rBuf
            Case .User_Trade_Trade: Data_User_Trade_Trade rBuf
            Case .User_Trade_UpdateTrade: Data_User_Trade_UpdateTrade rBuf

            Case .Combo_ProjectileSoundRotateDamage: Data_Combo_ProjectileSoundRotateDamage rBuf
            Case .Combo_SlashSoundRotateDamage: Data_Combo_SlashSoundRotateDamage rBuf
            Case .Combo_SoundRotateDamage: Data_Combo_SoundRotateDamage rBuf

            Case Else
                If DEBUG_PrintPacketReadErrors Then Debug.Print "Command ID " & CommandID & " caused a premature packet handling abortion!"
                Exit Do 'Something went wrong or we hit the end, either way, RUN!!!!

            End Select
        End With

        'Exit when the buffer runs out
        If rBuf.Get_ReadPos > BufUBound Then Exit Do

    Loop

End Sub

Private Sub GOREsock_OnConnecting(inSox As Long)

    If SocketOpen = 0 Then

        PacketInPos = 0
        PacketOutPos = 0
        
        Sleep 50
        DoEvents
        
        'Pre-saved character
        If SendNewChar = False Then
            sndBuf.Put_Byte DataCode.User_Login
            sndBuf.Put_String UserName
            sndBuf.Put_String UserPassword
        Else
            'New character
            sndBuf.Put_Byte DataCode.User_NewLogin
            sndBuf.Put_String UserName
            sndBuf.Put_String UserPassword
            sndBuf.Put_Integer UserHead
            sndBuf.Put_Integer UserBody
            sndBuf.Put_Byte UserClass
        End If
    
        'Save Game.ini
        If frmConnect.SavePassChk.Value = 0 Then UserPassword = vbNullString
        Var_Write DataPath & "Game.ini", "INIT", "Name", UserName
        Var_Write DataPath & "Game.ini", "INIT", "Password", UserPassword
        
        'Send the data
        Data_Send
        DoEvents
    
    End If
    
End Sub
