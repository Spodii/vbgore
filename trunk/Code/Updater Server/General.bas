Attribute VB_Name = "General"
Option Explicit

'Local socket ID
Public LocalID As Long

'User information
Public Const MaxConnections As Byte = 10
Public Type User
    ConnID As Integer   'ID of the user's socket
    CurrFile As Long    'File the user is currently downloading
    LastBytesSent As Long   'Used to calculate the amount sent from a file
End Type
Public UserList(1 To MaxConnections) As User

'Upload amount
Public Const UploadBufferSize As Byte = 1  'How many seconds to average over
Public UploadBuffer(1 To UploadBufferSize) As Long
Public UploadCount As Long 'Counts the amount uploaded (reset by timer)

'Used to calculate MBs uploaded
Public upBytes As Long
Public upKBytes As Long
Public upMBytes As Long

'Counter to clear the data (clearing after every transfer makes a lot of flashing - very annoying)
Public ClearDataCount As Byte

'How many people are currently connected
Public CurrConnections As Long

'Totals
Public FilesUploaded As Long   'Total amount of files uploaded
Public ConnectionsEst As Long  'Total amount of connections established

'File list information
Public FileList() As String            'List of files by their complete path on the server
Public FileListShortName() As String   'List of files by their shortened path
Public FileSize() As Long              'Size of the file in bytes
Public FileHash() As String * 32       'MD5 hash of the file
Public NumFiles As Long

'Path to the compressed file (CompressPath + FileListShortName)
Public CompressPath As String

Private Type bArray
    b() As Byte
End Type

'Cached packet headers
Private PH_EOF As bArray
Private PH_DONE As bArray
Private PH_FILE() As bArray

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Sub ReleaseCapture Lib "User32" ()
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Sub Initialize()
Dim FileNum As Byte
Dim i As Long
Dim j As Integer

    GOREsock_Initialize frmMain.hwnd

    InitFilePaths
    
    frmMain.StatusLbl.Caption = "Loading file list"
    DoEvents
    
    'Set the compressed path
    CompressPath = App.Path & "\_Compressed"

    'Get the file list
    FileList() = AllFilesInFolders(App.Path & "\UpdateFiles\", True)
    NumFiles = UBound(FileList())
    
    'Quit if we have no files to update
    If NumFiles = 0 Then
        MsgBox "Error: You must include files to update to run the server!" & vbCrLf & _
            "Place them in the following path:" & vbCrLf & vbCrLf & _
            App.Path & "\UpdateFiles\", vbOKOnly
        Unload frmMain
        Exit Sub
    End If

    'Create the short file list
    ReDim FileListShortName(0 To NumFiles)
    j = Len(App.Path & "UpdateFiles\")
    For i = 0 To NumFiles
        FileListShortName(i) = Right$(FileList(i), Len(FileList(i)) - j)
    Next i
    
    frmMain.StatusLbl.Caption = "Loading file sizes"
    DoEvents

    'Create the file size list
    ReDim FileSize(0 To NumFiles)
    FileNum = FreeFile
    For i = 0 To NumFiles
        Open FileList(i) For Append As #FileNum
        FileSize(i) = LOF(FileNum)
        Close #FileNum
    Next i
    
    frmMain.StatusLbl.Caption = "Creating MD5 hashes"
    DoEvents
    
    'Create MD5 hashes
    ReDim FileHash(0 To NumFiles)
    For i = 0 To NumFiles
        FileHash(i) = MD5_File(FileList(i))
    Next i
    
    frmMain.StatusLbl.Caption = "Compressing files"
    DoEvents
    
    'Create compressed files
    For i = 0 To NumFiles
        If Engine_FileExist(App.Path & "\_Compressed" & FileListShortName(i), vbNormal) Then Kill App.Path & "\_Compressed" & FileListShortName(i)
        MakeSureDirectoryPathExists App.Path & "\_Compressed" & FileListShortName(i)
        Compression_Compress FileList(i), App.Path & "\_Compressed" & FileListShortName(i), LZW
        DoEvents
    Next i
    
    frmMain.StatusLbl.Caption = "Creating socket"
    DoEvents
    
    'Create the packet header caches
    PH_EOF.b() = StrConv("****ENDOFFILE****", vbFromUnicode)
    PH_DONE.b() = StrConv("****DONE****", vbFromUnicode)
    ReDim PH_FILE(0 To NumFiles)
    For i = 0 To NumFiles
        PH_FILE(i).b() = StrConv("****sendrequest****|" & FileListShortName(i) & "|" & FileSize(i) & "|" & FileHash(i), vbFromUnicode)
    Next i

    'Start up the socket (change the ip to 0.0.0.0 or your internal IP)
    LocalID = GOREsock_Listen("127.0.0.1", Val(Var_Get(ServerDataPath & "Server.ini", "INIT", "UpdatePort")))
    GOREsock_SetOption LocalID, soxSO_TCP_NODELAY, False
    
    If GOREsock_Address(LocalID) = "-1" Then MsgBox "Error while creating server connection. Please make sure you are connected to the internet and supplied a valid IP" & vbCrLf & "Make sure you use your INTERNAL IP, which can be found by Start -> Run -> 'Cmd' (Enter) -> IPConfig" & vbCrLf & "Finally, make sure you are NOT running another instance of the server, since two applications can not bind to the same port. If problems persist, you can try changing the port.", vbOKOnly

    frmMain.StatusLbl.Caption = "Loaded!"

End Sub

Public Sub GOREsock_Close(inSox As Long)
Dim i As Long
    
    'Remove the user's ConnID if they are on the socket that just closed
    For i = 1 To UBound(UserList)
        If UserList(i).ConnID = inSox Then
            UserList(i).ConnID = 0
            Exit For
        End If
    Next i

End Sub

Public Sub GOREsock_Connection(inSox As Long)

    'Assign the inSox ID as the ConnID
    UserList(inSox).ConnID = inSox

    'Start the transfering with the first file
    UserList(inSox).CurrFile = 0
    UserList(inSox).LastBytesSent = 0
    DoEvents
    SendFileRequest inSox, UserList(inSox).CurrFile

    'Raise connections count
    ConnectionsEst = ConnectionsEst + 1
    frmMain.ConnectionsEstablishedLbl.Caption = ConnectionsEst

End Sub

Public Sub GOREsock_DataArrival(inSox As Long, inData() As Byte)
Dim FileNum As Byte
Dim rData As String
Dim l As Long
Dim b() As Byte

'Turn the data into a string

    rData = StrConv(inData(), vbUnicode)

    'The file was accepted to be transfered
    If Left$(rData, 10) = "****ok****" Then

        'Send the whole file
        FileNum = FreeFile
        Open CompressPath & FileListShortName(UserList(inSox).CurrFile) For Binary Access Read As #FileNum
            l = LOF(FileNum)    'Get the size of the file
            ReDim b(0 To l + UBound(PH_EOF.b))  'Redim enough for the file + the EOF message
            Get #FileNum, , b   'Grab the whole file
        Close #FileNum
        CopyMemory b(l), PH_EOF.b(0), UBound(PH_EOF.b()) + 1
        GOREsock_SendData inSox, b()

        'Raise files uploaded count
        FilesUploaded = FilesUploaded + 1
        frmMain.FilesUploadedLbl.Caption = FilesUploaded

        'Check if the user has finished updating
        If UserList(inSox).CurrFile > NumFiles Then
        
            'Tell the user they have finished
            GOREsock_SendData inSox, PH_DONE.b()
            
        Else
        
            'Send the next file
            SendFileRequest inSox, UserList(inSox).CurrFile
            
        End If
        Exit Sub

    End If

    'The file was not accepted to be transfered - the user must already be up to date
    If rData = "****no****" Then

        'Raise file number
        UserList(inSox).CurrFile = UserList(inSox).CurrFile + 1

        'Check if the user has finished updating
        If UserList(inSox).CurrFile > NumFiles Then
            'Tell the user they have finished
            GOREsock_SendData inSox, PH_DONE.b()
            UserList(inSox).ConnID = 0
            UserList(inSox).CurrFile = 0
            UserList(inSox).LastBytesSent = 0
        Else
            'Send the next file
            SendFileRequest inSox, UserList(inSox).CurrFile
        End If
        Exit Sub

    End If

End Sub

Public Sub GOREsock_SendComplete(inSox As Long)

'Clear information

    UserList(inSox).LastBytesSent = 0

End Sub

Public Sub GOREsock_RecvProgress(inSox As Long, bytesRecv As Long, bytesRemaining As Long)

'*********************************************
'Empty procedure
'*********************************************

End Sub

Sub GOREsock_Connecting(inSox As Long)

'*********************************************
'Empty procedure
'*********************************************

End Sub

Public Sub GOREsock_SendProgress(inSox As Long, bytesSent As Long, bytesRemaining As Long)

    'Update the sending information
    If UserList(inSox).CurrFile <= NumFiles Then frmMain.FileNameLbl.Caption = FileListShortName(UserList(inSox).CurrFile)
    frmMain.PercentLbl.Caption = Round(bytesSent / (bytesSent + bytesRemaining), 2) * 100 & "%"
    frmMain.ClientIPLbl.Caption = GOREsock_Address(inSox)
    frmMain.ClientIDLbl.Caption = inSox

    'Update the bytes remaining info
    UploadCount = UploadCount + (bytesSent - UserList(inSox).LastBytesSent)
    UserList(inSox).LastBytesSent = bytesSent

End Sub

Public Sub SendFileRequest(ByVal UserIndex As Long, ByVal FileIndex As Long)
Dim FileNum As Byte
Dim b() As Byte

    'Check for valid file index
    If FileIndex < 0 Then Exit Sub
    If FileIndex > NumFiles Then Exit Sub

    'Send the request to start transfering the file
    DoEvents
    GOREsock_SendData UserIndex, PH_FILE(FileIndex).b()

End Sub

Public Function Engine_FileExist(File As String, FileType As VbFileAttribute) As Boolean

    'Checks if a file exists
    Engine_FileExist = (Dir$(File, FileType) <> "")

End Function

Public Function Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

'*****************************************************************
'Gets a variable from a text file
'*****************************************************************
Dim sSpaces As String

    sSpaces = Space$(100)
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File

    Var_Get = RTrim$(sSpaces)
    Var_Get = Left$(Var_Get, Len(Var_Get) - 1)

End Function
