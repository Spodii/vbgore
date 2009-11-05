VERSION 5.00
Object = "{9842967E-F54F-4981-93DF-0772B2672E38}#1.0#0"; "vbgoresocketbinary.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Packet Sender"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4560
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
   ForeColor       =   &H8000000E&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   2  'CenterScreen
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
'**       ____        _________   ______   ______  ______   _______           **
'**       \   \      /   /     \ /  ____\ /      \|      \ |   ____|          **
'**        \   \    /   /|      |  /     |        |       ||  |____           **
'***        \   \  /   / |     /| |  ___ |        |      / |   ____|         ***
'****        \   \/   /  |     \| |  \  \|        |   _  \ |  |____         ****
'******       \      /   |      |  \__|  |        |  | \  \|       |      ******
'********      \____/    |_____/ \______/ \______/|__|  \__\_______|    ********
'*******************************************************************************
'*******************************************************************************
'************ vbGORE - Visual Basic 6.0 Graphical Online RPG Engine ************
'************            Official Release: Version 0.1.2            ************
'************                 http://www.vbgore.com                 ************
'*******************************************************************************
'*******************************************************************************
'***** Source Distribution Information: ****************************************
'*******************************************************************************
'** If you wish to distribute this source code, you must distribute as-is     **
'** from the vbGORE website unless permission is given to do otherwise. This  **
'** comment block must remain in-tact in the distribution. If you wish to     **
'** distribute modified versions of vbGORE, please contact Spodi (info below) **
'** before distributing the source code. You may never label the source code  **
'** as the "Official Release" or similar unless the code and content remains  **
'** unmodified from the version downloaded from the official website.         **
'** You may also never sale the source code without permission first. If you  **
'** want to sell the code, please contact Spodi (below). This is to prevent   **
'** people from ripping off other people by selling an insignificantly        **
'** modified version of open-source code just to make a few quick bucks.      **
'*******************************************************************************
'***** Creating Engines With vbGORE: *******************************************
'*******************************************************************************
'** If you plan to create an engine with vbGORE that, please contact Spodi    **
'** before doing so. You may not sell the engine unless told elsewise (the    **
'** engine must has substantial modifications), and you may not claim it as   **
'** all your own work - credit must be given to vbGORE, along with a link to  **
'** the vbGORE homepage. Failure to gain approval from Spodi directly to      **
'** make a new engine with vbGORE will result in first a friendly reminder,   **
'** followed by much more drastic measures.                                   **
'*******************************************************************************
'***** Helping Out vbGORE: *****************************************************
'*******************************************************************************
'** If you want to help out with vbGORE's progress, theres a few things you   **
'** can do:                                                                   **
'**  *Donate - Great way to keep a free project going. :) Info and benifits   **
'**        for donating can be found at:                                      **
'**        http://www.vbgore.com/en/index.php?title=Donate                    **
'**  *Contribute - Check out our forums, contribute ideas, report bugs, or    **
'**        help expend the wiki pages!                                        **
'**  *Link To Us - Creating a link to vbGORE, whether it is on your own web   **
'**        page or a link to vbGORE in a forum you visit, every link helps    **
'**        spread the word of vbGORE's existance! Buttons and banners for     **
'**        linking to vbGORE can be found on the following page:              **
'**        http://www.vbgore.com/en/index.php?title=Buttons_and_Banners       **
'*******************************************************************************
'***** Conact Information: *****************************************************
'*******************************************************************************
'** Please contact the creator of vbGORE (Spodi) directly with any questions: **
'** AIM: Spodii                          Yahoo: Spodii                        **
'** MSN: Spodii@hotmail.com              Email: spodi@vbgore.com              **
'** 2nd Email: spodii@hotmail.com        Website: http://www.vbgore.com       **
'*******************************************************************************
'***** Credits: ****************************************************************
'*******************************************************************************
'** Below are credits to those who have helped with the project or who have   **
'** distributed source code which has help this project's creation. The below **
'** is listed in no particular order of significance:                         **
'**                                                                           **
'** ORE (Aaron Perkins): Used as base engine and for learning experience      **
'**   http://www.baronsoft.com/                                               **
'** SOX (Trevor Herselman): Used for all the networking                       **
'**   http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=35239&lngWId=1      **
'** Compression Methods (Marco v/d Berg): Provided compression algorithms     **
'**   http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=37867&lngWId=1      **
'** All Files In Folder (Jorge Colaccini): Algorithm implimented into engine  **
'**   http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=51435&lngWId=1      **
'** Game Programming Wiki (All community): Help on many different subjects    **
'**   http://wwww.gpwiki.org/                                                 **
'**                                                                           **
'** Also, all the members of the vbGORE community who have submitted          **
'** tutorials, bugs, suggestions, criticism and have just stuck around!!      **
'**                                                                           **
'** If you feel you belong in these credits, please contact Spodi (above).    **
'*******************************************************************************
'*******************************************************************************

Option Explicit

'Our current sending progress
Private NumBytes As Integer
Private ByteVal() As Byte
Private Const MaxBytes As Long = 5000   'Even if we reach this value, we'd just be sending broken packets

'Misc variables
Private s As String
Private SoxID As Long
Private SocketOpen As Byte
Private Connected As Byte
Private sndBuf As DataBuffer

Private Sub DispTmr_Timer()

    'Display information - this program is supposed to hardly cause any crashes and run for hours,
    ' even days at a time. There is no point to showing every number change since it slows things
    ' down a LOT!
    Me.Cls
    Me.Print "Size: " & NumBytes
    Me.Print s
    
End Sub

Private Sub Form_Load()
Dim j As Byte
Dim FileNum As Byte
    
    'Load our saved state if it exists
    If FileExist(App.Path & "\Data2\packetflooder.dat", vbNormal) Then
        FileNum = FreeFile
        Open App.Path & "\Data2\packetflooder.dat" For Binary As FileNum
            Get #FileNum, , NumBytes
            ReDim ByteVal(NumBytes)
            For j = 1 To NumBytes
                Get #FileNum, , ByteVal(j)
            Next j
        Close #FileNum
    Else
        'Set up our basic variables
        NumBytes = 1
        ReDim ByteVal(MaxBytes)
    End If
    
    'Turn off nagling since bandwidth is not our worry if we are working locally
    Sox.SetOption SoxID, soxSO_TCP_NODELAY, True
    
    'Show our form
    Me.Show
    Me.Print "Connecting to server..."

    'Create the buffer
    Set sndBuf = New DataBuffer

    'Connect to the server
    Do While Connected = 0
    
        'Make the connected
        SoxID = Sox.Connect("127.0.0.1", 10200)
        
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
Dim j As String

    'Build our next byte value
    i = 1
    Do While ByteVal(i) >= 255
        If ByteVal(NumBytes) = 255 Then NumBytes = NumBytes + 1
        ByteVal(i) = 0
        ByteVal(i + 1) = ByteVal(i + 1) + 1
        i = i + 1
    Loop
    ByteVal(1) = ByteVal(1) + 1
    
    'Clear the display string
    s = ""

    'Build the buffer
    sndBuf.Clear
    For i = 1 To NumBytes
    
        'Put the byte in the buffer
        sndBuf.Put_Byte ByteVal(i)
        
        'Update the display
        j = ByteVal(i)
        If ByteVal(i) < 100 Then j = "0" & ByteVal(i)
        If ByteVal(i) < 10 Then j = "00" & ByteVal(i)
        s = s & j
        If (i Mod 11) = 0 Then
            s = s & vbCrLf
        Else
            If i <> NumBytes Then s = s & " "
        End If
        
    Next i
    Data_Send

End Sub

Private Sub Sox_OnDataArrival(inSox As Long, inData() As Byte)

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
