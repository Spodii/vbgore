Attribute VB_Name = "General"
Option Explicit

'Local socket ID
Public LocalID As Long

'Server file list information
Public Type ServerFile
    Path As String      'Path of the file from the server
    Size As Long        'Size of the file from the server
    Hash As String * 32 'MD5 hash of the file from the server
    NeedFile As Boolean 'If the file is needed to be updated
End Type
Public ServerFile() As ServerFile

'File we are currently on
Public FileIndex As Long

'The size of all the files that we need combined
Public TotalNeed As Long

'The size of all the files we have current acquired (only counts finished files)
Public TotalGot As Long

'Time when we first started receiving stuff
Public StartTime As Long

'Last time we updated the KBps
Public UpdateTime As Long

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Type typHOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 127) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Public Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Public Declare Sub ReleaseCapture Lib "User32" ()
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationname As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Sub apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function apiGetHostByName Lib "wsock32" Alias "gethostbyname" (ByVal HostName As String) As Long
Private Declare Function WSACleanup Lib "wsock32" () As Long
Private Declare Function WSAStartup Lib "wsock32" (ByVal VersionReq As Long, WSADataReturn As WSADATA) As Long

Private Function IsIP(ByVal IPAddress As String) As Boolean
Dim s() As String
Dim i As Long

    'If there are no periods, I have no idea what we have...
    If InStr(1, IPAddress, ".") = 0 Then Exit Function
    
    'Split up the string by the periods
    s = Split(IPAddress, ".")
    
    'Confirm we have ubound = 3, since xxx.xxx.xxx.xxx has 4 elements and we start at index 0
    If UBound(s) <> 3 Then Exit Function
    
    'Check that the values are numeric and in a valid range
    For i = 0 To 3
        If Val(s(i)) < 0 Then Exit Function
        If Val(s(i)) > 255 Then Exit Function
    Next i
    
    'Looks like we were passed a valid IP!
    IsIP = True
    
End Function

Public Function GetIPFromHost(ByVal HostName As String) As String
Dim udtWSAData As WSADATA
Dim HostAddress As Long
Dim HostInfo As typHOSTENT
Dim IPLong As Long
Dim IPBytes() As Byte
Dim i As Integer

    On Error Resume Next
    
    If WSAStartup(257, udtWSAData) Then
        MsgBox "Error initializing winsock on WSAStartup!"
        GetIPFromHost = HostName
        Exit Function
    End If

    'Make sure a HTTP:// or FTP:// something wasn't added... some people like to do that
    If UCase$(Left$(HostName, 7)) = "HTTP://" Then
        HostName = Right$(HostName, Len(HostName) - 7)
    ElseIf UCase$(Left$(HostName, 6)) = "FTP://" Then
        HostName = Right$(HostName, Len(HostName) - 6)
    End If
    
    'If we were already passed an IP, just abort since we have what we want
    If IsIP(HostName) Then
        GetIPFromHost = HostName
        Exit Function
    End If
    
    'Get the host address
    HostAddress = apiGetHostByName(HostName)
    
    'Failure!
    If HostAddress = 0 Then Exit Function
    
    'Move the memory around to get it in a format we can read
    apiCopyMemory HostInfo, HostAddress, LenB(HostInfo)
    apiCopyMemory IPLong, HostInfo.hAddrList, 4
    
    'Get the number of parts to the IP (will always be 4 as far as I know)
    ReDim IPBytes(1 To HostInfo.hLength)

    'Convert the address, stored in the format of a long, to 4 bytes (just simple long -> byte array conversion)
    apiCopyMemory IPBytes(1), IPLong, HostInfo.hLength
    
    'Add in the periods
    For i = 1 To HostInfo.hLength
        GetIPFromHost = GetIPFromHost & IPBytes(i) & "."
    Next
    
    'Remove the final period
    GetIPFromHost = Left$(GetIPFromHost, Len(GetIPFromHost) - 1)
    
    'Clean up the socket
    WSACleanup

End Function

Private Sub CommandLine(ByVal CommandLineString As String)
Dim Start As STARTUPINFO
Dim Proc As PROCESS_INFORMATION

    Start.dwFlags = &H1
    Start.wShowWindow = 1
    CreateProcessA 0&, CommandLineString, 0&, 0&, 1&, &H20&, 0&, 0&, Start, Proc
    Do While WaitForSingleObject(Proc.hProcess, 0) = 258
        DoEvents
        Sleep 1
    Loop

End Sub

Public Sub GOREsock_Connection(inSox As Long)

'*********************************************
'Empty procedure
'*********************************************

End Sub

Public Sub GOREsock_Close(inSox As Long)

    frmMain.ConnectCmd.Enabled = True

End Sub

Public Sub GOREsock_DataArrival(inSox As Long, inData() As Byte)

'*********************************************
'Handle data received from the server
'*********************************************
Dim Data As String
Dim b() As Byte
Dim TempS() As String
Dim TempS2() As String
Dim i As Long
Dim FileNum As Byte

    If StartTime = 0 Then StartTime = timeGetTime

    Data = StrConv(inData(), vbUnicode)
    
    'File list from the server
    If Left$(Data, 8) = "***FL***" Then
    
        'Get our data and decompress it
        Data = Right$(Data, Len(Data) - 8)
        b() = StrConv(Data, vbFromUnicode)
        Compression_DeCompress_LZMA b()
        Data = StrConv(b(), vbUnicode)
        
        'Split up the files
        FileIndex = 0
        TotalNeed = 0
        TempS = Split(Data, Chr$(254))
        ReDim Preserve ServerFile(0 To UBound(TempS))
        For i = 0 To UBound(TempS)
        
            'Split up the file information
            TempS2 = Split(TempS(i), Chr$(255))

            'Store the data
            ServerFile(i).Path = App.Path & TempS2(0)
            ServerFile(i).Size = Val(TempS2(1))
            ServerFile(i).Hash = TempS2(2)
            ServerFile(i).NeedFile = NeedFile(i)
            
            'Add to the total size
            If ServerFile(i).NeedFile Then TotalNeed = TotalNeed + ServerFile(i).Size
            
        Next i
        
        'Get the first file
        RequestNextFile
        
    End If
    
    'Server sends the file
    If Right$(Data, 9) = "***EOF***" Then
        
        'Get our data, decompress it and save it to the file
        Data = Left$(Data, Len(Data) - 9)
        b() = StrConv(Data, vbFromUnicode)
        DoEvents
        
        'Find which decompression to use (MonkeyAudio for *.wav, LZMA for the rest)
        If Len(Data) > 0 Then
            If LCase$(Right$(ServerFile(FileIndex).Path, 4)) = ".wav" Then
                Compression_DeCompress_MonkeyAudio b()
            Else
                Compression_DeCompress_LZMA b()
            End If
        End If
        
        'Make the directory path if needed
        DoEvents
        MakeSureDirectoryPathExists ServerFile(FileIndex).Path
        
        'Save the file
        If Engine_FileExist(ServerFile(FileIndex).Path, vbNormal) Then Kill ServerFile(FileIndex).Path
        FileNum = FreeFile
        Open ServerFile(FileIndex).Path For Binary Access Write As #FileNum
            Put #FileNum, , b()
        Close #FileNum
        
        'Register the file if it is an OCX or DLL
        If LCase$(Right$(ServerFile(FileIndex).Path, 4)) = ".ocx" Or LCase$(Right$(ServerFile(FileIndex).Path, 4)) = ".dll" Then
            CommandLine "regsvr32 " & Chr$(34) & ServerFile(FileIndex).Path & Chr$(34) & " /s"
        End If
        
        'Confirm the file data
        ServerFile(FileIndex).NeedFile = NeedFile(FileIndex)

        TotalGot = TotalGot + ServerFile(FileIndex).Size
        
        DoEvents
        
        'Request the next file
        RequestNextFile
        
    End If

End Sub

Public Sub FinishUpdate()

    'Close down the connection
    GOREsock_Shut LocalID
    GOREsock_ShutDown
    
    frmMain.StatusLbl.Caption = "Download Successful!"
    frmMain.FileLbl.Caption = vbNullString
    frmMain.PercentLbl.Caption = "100%"
    
    'Load the client
    If MsgBox("The update has been completed! Do you wish to run the client now?", vbYesNo) = vbYes Then
        ShellExecute frmMain.hWnd, vbNullString, App.Path & "\GameClient.exe", "-sdf@041jkdf0)21`~", vbNullString, 1
    End If
    
    'Unload the updater
    GOREsock_Shut LocalID
    DoEvents
    GOREsock_UnHook
    DoEvents
    
    'Initiate the closedown (gives the socket time to unload)
    frmMain.CloseTimer.Enabled = True

End Sub

Public Sub RequestNextFile()
Dim s As String

    'Loop until we find a file we need
    Do While ServerFile(FileIndex).NeedFile = False
        FileIndex = FileIndex + 1   'We add one here so that way we can confirm the file we just got is valid
        If FileIndex > UBound(ServerFile) Then
            FinishUpdate
            Exit Sub
        End If
    Loop
    
    'Delete the file version we have since it is out of date
    If LenB(Dir$(ServerFile(FileIndex).Path)) Then Kill ServerFile(FileIndex).Path
    
    'Request the updated file
    s = "***GET***" & FileIndex
    frmMain.FileLbl.Caption = Right$(ServerFile(FileIndex).Path, Len(ServerFile(FileIndex).Path) - Len(App.Path))
    GOREsock_SendData LocalID, StrConv(s, vbFromUnicode)

End Sub

Public Function NeedFile(ByVal i As Long) As Boolean
Dim FileNum As Byte
Dim fSize As Long
Dim fHash As String * 32

    'Check if we already have the file
    If LenB(Dir$(ServerFile(i).Path)) <> 0 Then
        
        'Check for Game.ini
        If UCase$(Right$(ServerFile(i).Path, 9)) = "\GAME.INI" Then
            NeedFile = False
            Exit Function
        End If
        
        'We have the file, compare the size
        FileNum = FreeFile
        Open ServerFile(i).Path For Binary Access Read As #FileNum
            fSize = LOF(FileNum)
        Close #FileNum
        If fSize = ServerFile(i).Size Then
        
            'File size is the same, compare the MD5 hashes
            fHash = MD5_File(ServerFile(i).Path)
            If fHash = ServerFile(i).Hash Then
                
                'We don't need the file
                NeedFile = False
                Exit Function
                
            End If
        End If
    End If
    
    'One of the tests failed, we need the file
    NeedFile = True
    
End Function

Public Sub GOREsock_RecvProgress(inSox As Long, bytesRecv As Long, bytesRemaining As Long)
Dim TotalRecv As Long
Dim ElapsedSecs As Long

    TotalRecv = bytesRecv + TotalGot

    'Prevent division by 0 errors
    If TotalNeed > 0 Then
        
        'Update recieved percentage
        frmMain.PercentLbl.Caption = Round(TotalRecv / TotalNeed, 2) * 100 & "%"
        
    End If
    
    'Update the download speed
    If StartTime > 0 Then
        If (timeGetTime - StartTime) > 1000 Then
            If UpdateTime + 1000 < timeGetTime Then
            
                ElapsedSecs = (timeGetTime - StartTime) * 0.001
                If ElapsedSecs < 1 Then ElapsedSecs = 1
            
                frmMain.spid.Caption = Int((TotalRecv \ ElapsedSecs) \ 1024)
                UpdateTime = timeGetTime
            End If
        End If
    End If
    
End Sub

Sub GOREsock_Connecting(inSox As Long)

'*********************************************
'Empty procedure
'*********************************************

End Sub

Public Sub GOREsock_SendProgress(inSox As Long, bytesSent As Long, bytesRemaining As Long)

'*********************************************
'Empty procedure
'*********************************************

End Sub

Public Sub GOREsock_SendComplete(inSox As Long)

'*********************************************
'Empty procedure
'*********************************************

End Sub

Public Function Engine_FileExist(File As String, FileType As VbFileAttribute) As Boolean

    'Checks if a file exists
    Engine_FileExist = (Dir$(File, FileType) <> "")

End Function
