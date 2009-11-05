VERSION 5.00
Object = "{D1EE5822-4214-490C-81BE-49A1E232B2F0}#1.0#0"; "vbgoresocketbinary.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Packet Sender"
   ClientHeight    =   450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7560
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   30
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   19
      Left            =   6960
      TabIndex        =   19
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   18
      Left            =   6600
      TabIndex        =   18
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   17
      Left            =   6240
      TabIndex        =   17
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   16
      Left            =   5880
      TabIndex        =   16
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   15
      Left            =   5520
      TabIndex        =   15
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   14
      Left            =   5160
      TabIndex        =   14
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   13
      Left            =   4800
      TabIndex        =   13
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   12
      Left            =   4440
      TabIndex        =   12
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   11
      Left            =   4080
      TabIndex        =   11
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   10
      Left            =   3720
      TabIndex        =   10
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   9
      Left            =   3360
      TabIndex        =   9
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   8
      Left            =   3000
      TabIndex        =   8
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   7
      Left            =   2640
      TabIndex        =   7
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   6
      Left            =   2280
      TabIndex        =   6
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   5
      Left            =   1920
      TabIndex        =   5
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox ByteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.Timer DispTmr 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   0
   End
   Begin SoxOCX.Sox Sox 
      Height          =   420
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Defines how we want to flood our packets
'If this byte = 0, then we will wait for a response back from the server for every packet we send to send the next one
'If this byte = 1, then we will send packets as fast as we possibly can until we break something / everything
Private Const HeavyFlooding As Byte = 0

'Our current sending progress
Private NumBytes As Integer
Private ByteVal() As Integer
Private Const MaxBytes As Long = 20 'Even if we reach this value, we'd just be sending broken packets

'Misc variables
Private SoxID As Long
Private SocketOpen As Byte
Private Connected As Byte
Private sndBuf As DataBuffer

Private Sub ByteTxt_Change(Index As Integer)
Dim b As Byte

    'Make sure it is a value byte value (or -1)
    If Val(ByteTxt(Index).Text) > 255 Then
        ByteTxt(Index).Text = 255
        Exit Sub
    End If
    If Val(ByteTxt(Index).Text) < 0 Then
        ByteTxt(Index).Text = 0
        Exit Sub
    End If
    On Error GoTo ErrOut
    b = Val(ByteTxt(Index).Text)
    On Error GoTo 0

    'Change the value
    ByteVal(Index + 1) = b
    
    'Change the NumBytes
    If ByteVal(Index + 1) > 0 Then
        If (Index + 1) > NumBytes Then NumBytes = (Index + 1)
    End If
    
    Exit Sub
    
ErrOut:

    ByteTxt(Index).Text = 0

End Sub

Private Sub DispTmr_Timer()
Dim i As Long

    'Display information - this program is supposed to hardly cause any crashes and run for hours,
    ' even days at a time. There is no point to showing every number change since it slows things
    ' down a LOT!
    For i = 1 To MaxBytes
        If i > NumBytes Then
            ByteTxt(i - 1).Text = ""
        Else
            ByteTxt(i - 1).Text = ByteVal(i)
        End If
    Next i
    
End Sub

Private Sub Form_Load()
Dim j As Byte
Dim FileNum As Byte
    
    'Load our saved state if it exists
    ReDim ByteVal(MaxBytes)
    If FileExist(App.Path & "\Data2\packetflooder.dat", vbNormal) Then
        FileNum = FreeFile
        Open App.Path & "\Data2\packetflooder.dat" For Binary As FileNum
            Get #FileNum, , NumBytes
            For j = 1 To NumBytes
                Get #FileNum, , ByteVal(j)
            Next j
        Close #FileNum
    Else
        'Set up our basic variables
        NumBytes = 1
    End If
    
    'Turn off nagling since bandwidth is not our worry if we are working locally
    Sox.SetOption SoxID, soxSO_TCP_NODELAY, True
    
    'Show our form
    Me.Show

    'Create the buffer
    Set sndBuf = New DataBuffer

    'Connect to the server
    Do While Connected = 0
    
        'Make the connected
        SoxID = Sox.Connect("64.187.111.152", 7234)
        
        'Check if the connected fialed
        If SoxID = -1 Then
            'Connection failed, so just keep trying
            DoEvents
        Else
            'We connected, so move on
            Connected = 1
            Exit Do
        End If
        
    Loop
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim FileNum As Byte
Dim j As Long

    If Sox.ShutDown = soxERROR Then 'Terminate will be True if we have ShutDown properly
        If MsgBox("ShutDown procedure has not completed!" & vbCrLf & "(Hint - Select No and Try again!)" & vbCrLf & "Execute Forced ShutDown?", vbApplicationModal + vbCritical + vbYesNo, "UNABLE TO COMPLY!") = vbNo Then
            Let Cancel = True
            Exit Sub
        Else
            Sox.UnHook  'Unfortunately for now, I can't get around doing this automatically for you :( VB crashes if you don't do this!
        End If
    Else
        Sox.UnHook  'The reason is VB closes my Mod which stores the WindowProc function used for SubClassing and VB doesn't know that! So it closes the Mod before the Control!
    End If
    
    'Save state
    If NumBytes > 1 Then
        If MsgBox("Do you wish to save the current state?", vbYesNo) = vbYes Then
            If FileExist(App.Path & "\Data2\packetflooder.dat", vbNormal) Then Kill App.Path & "\Data2\packetflooder.dat"
            FileNum = FreeFile
            Open App.Path & "\Data2\packetflooder.dat" For Binary As FileNum
                Put #FileNum, , NumBytes
                For j = 1 To NumBytes
                    Put #FileNum, , ByteVal(j)
                Next j
            Close #FileNum
        End If
    End If

    'Force to quit the connection loop
    Connected = 1
    Set sndBuf = Nothing
    Erase ByteVal
    Unload Me
    End

End Sub

Function FileExist(file As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

    FileExist = (Dir$(file, FileType) <> "")

End Function

Private Sub SendNextPacket()
Dim i As Byte

    'Build our next byte value
    i = 1
    Do While ByteVal(i) >= 255
        If ByteVal(NumBytes) = 255 Then NumBytes = NumBytes + 1
        ByteVal(i) = 0
        ByteVal(i + 1) = ByteVal(i + 1) + 1
        i = i + 1
    Loop
    ByteVal(1) = ByteVal(1) + 1

    'Build the buffer
    sndBuf.Clear
    For i = 1 To NumBytes
    
        'Put the byte in the buffer
        sndBuf.Put_Byte ByteVal(i)

    Next i
    Data_Send

End Sub

Private Sub Sox_OnDataArrival(inSox As Long, inData() As Byte)

    'Wait for a packet (first byte = 111) from the server saying it is ready for the next one
    If HeavyFlooding = 0 Then
        If inData(0) = 111 Then SendNextPacket
    End If
    
End Sub

Private Sub Sox_OnSendComplete(inSox As Long)

    'If we are doing heavy flooding, we wait for no man - fire at will!
    SendNextPacket

End Sub

Private Sub Sox_OnState(inSox As Long, inState As SoxOCX.enmSoxState)

    'Exit if we are already connected
    If SocketOpen = 1 Then Exit Sub
    
    'Check if we reached the first idle state
    If Sox.State(SoxID) = soxIdle Then
        SocketOpen = 1
        
        DispTmr.Enabled = True
        
        'Connect to our primary char - default PcktFloodr
        'Never use the ID of an admin account for safety reasons along with so they dont spam dev commands (like map saving)!
        sndBuf.Put_Byte 29  'ID for the new login packet
        sndBuf.Put_String "PcktFloodr"
        sndBuf.Put_String "vbgore"
        
        Data_Send
        
    End If

End Sub

Private Sub Data_Send()

'*********************************************
'Send data buffer to the server
'*********************************************
Dim TempBuffer() As Byte

    'Check if the socket is open
    If SocketOpen = 0 Then Exit Sub
    
    'Make sure we are in a valid state to send data
    If Not Sox.State(SoxID) = soxClosing Or soxERROR Or soxDisconnected Or soxIdle Then
        
        'Check that we have data to send
        If UBound(sndBuf.Get_Buffer) > 0 Then
        
            'Set our temp buffer
            ReDim TempBuffer(UBound(sndBuf.Get_Buffer))
            TempBuffer() = sndBuf.Get_Buffer
            
            'Crop off the last byte, which will always be 0 - bad way to do it, but oh well
            ReDim Preserve TempBuffer(UBound(TempBuffer) - 1)
        
            'Send the data
            Sox.SendData SoxID, TempBuffer()
            
            'Clear the buffer, get it ready for next use
            sndBuf.Clear
            
        End If
        
    End If

End Sub
