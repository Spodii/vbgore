Attribute VB_Name = "General"
Option Explicit

'Local socket ID
Public LocalID As Long

'User information
Public Const MaxConnections As Byte = 100
Public Type User
    ConnID As Integer   'ID of the user's socket
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

'How many people are currently connected
Public CurrConnections As Long

'Totals
Public ConnectionsEst As Long  'Total amount of connections established

'File list information
Public FileList() As String            'List of files by their complete path on the server
Public FileListShortName() As String   'List of files by their shortened path
Public NumFiles As Long

'The list of file information sent to the client at connection
Public ServerFileList() As Byte

'Path to the compressed file (CompressPath + FileListShortName)
Public CompressPath As String

Public EOFTag() As Byte

Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Sub Initialize()
Dim FileSize() As Long              'Size of the file in bytes
Dim FileHash() As String * 32       'MD5 hash of the file
Dim FileNum As Byte
Dim i As Long
Dim j As Integer
Dim s As String
Dim b() As Byte

    GOREsock_Initialize frmMain.hwnd

    InitFilePaths
    
    frmMain.StatusLbl.Caption = "Loading file list"
    DoEvents
    
    EOFTag() = StrConv("***EOF***", vbFromUnicode)
    
    'Set the compressed path
    CompressPath = App.Path & "\_Compressed"

    'Get the file list
    FileList() = AllFilesInFolders(App.Path & "\UpdateFiles\", True)
    On Error GoTo ErrOut
    NumFiles = UBound(FileList())
    On Error GoTo 0

    'Create the short file list
    ReDim FileListShortName(0 To NumFiles)
    j = Len(App.Path & "UpdateFiles\")
    For i = 0 To NumFiles
        FileListShortName(i) = Right$(FileList(i), Len(FileList(i)) - j)
    Next i

    'Create the file size list
    ReDim FileSize(0 To NumFiles)
    FileNum = FreeFile
    For i = 0 To NumFiles
        frmMain.StatusLbl.Caption = "Loading file sizes (" & Int((i / NumFiles) * 100) & "%)"
        DoEvents
        Open FileList(i) For Append As #FileNum
            FileSize(i) = LOF(FileNum)
        Close #FileNum
    Next i
    
    'Create MD5 hashes
    ReDim FileHash(0 To NumFiles)
    For i = 0 To NumFiles
        frmMain.StatusLbl.Caption = "Creating MD5 hashes (" & Int((i / NumFiles) * 100) & "%)"
        DoEvents
        FileHash(i) = MD5_File(FileList(i))
    Next i
    
    frmMain.StatusLbl.Caption = "Creating file list"
    DoEvents
    
    'Create the list of the files on the server and the information on them
    s = vbNullString
    For i = 0 To NumFiles
        s = s & FileListShortName(i) & Chr$(255) & FileSize(i) & Chr$(255) & FileHash(i)
        If i < NumFiles Then s = s & Chr$(254)
    Next i
    b = StrConv(s, vbFromUnicode)
    Compression_Compress_LZMA b()
    ReDim ServerFileList(0 To UBound(b) + 9)
    CopyMemory ServerFileList(8), b(0), UBound(b) + 1
    Erase b
    b = StrConv("***FL***", vbFromUnicode)
    CopyMemory ServerFileList(0), b(0), 8
    Erase b
    
    'Create compressed files
    For i = 0 To NumFiles
    
        'If the MD5 hashes are equal, we don't have to update the file
        j = 0
        If Engine_FileExist(App.Path & "\_Compressed" & FileListShortName(i) & ".md5", vbNormal) Then
            FileNum = FreeFile
            s = Space$(32)
            Open App.Path & "\_Compressed" & FileListShortName(i) & ".md5" For Binary Access Read As #FileNum
                Get #FileNum, , s
            Close #FileNum
            If s = FileHash(i) Then j = 1   'Hash matched!
        End If
        
        'If j = 0, the hash we found was invalid or didn't exist
        If j = 0 Then
            frmMain.StatusLbl.Caption = "Compressing files (" & Int((i / NumFiles) * 100) & "%)"
            DoEvents
            If Engine_FileExist(App.Path & "\_Compressed" & FileListShortName(i), vbNormal) Then Kill App.Path & "\_Compressed" & FileListShortName(i)
            MakeSureDirectoryPathExists App.Path & "\_Compressed" & FileListShortName(i)
            
            'Check whether to use normal compression or WAV-specific compression
            If LCase$(Right$(FileListShortName(i), 4)) = ".wav" Then
                Compression_Compress FileList(i), App.Path & "\_Compressed" & FileListShortName(i), MonkeyAudio
            Else
                Compression_Compress FileList(i), App.Path & "\_Compressed" & FileListShortName(i), LZMA
            End If
            
            'Since the hash didn't match, we have to store the hash we have now
            Open App.Path & "\_Compressed" & FileListShortName(i) & ".md5" For Binary Access Write As #FileNum
                Put #FileNum, , FileHash(i)
            Close #FileNum
            
        End If
        
    Next i
    
    frmMain.StatusLbl.Caption = "Creating socket"
    DoEvents

    'Start up the socket (change the ip to 0.0.0.0 or your internal IP)
    LocalID = GOREsock_Listen("127.0.0.1", Val(Var_Get(ServerDataPath & "Server.ini", "INIT", "UpdatePort")))
    GOREsock_SetOption LocalID, soxSO_TCP_NODELAY, False
    
    If GOREsock_Address(LocalID) = "-1" Then MsgBox "Error while creating server connection. Please make sure you are connected to the internet and supplied a valid IP" & vbCrLf & "Make sure you use your INTERNAL IP, which can be found by Start -> Run -> 'Cmd' (Enter) -> IPConfig" & vbCrLf & "Finally, make sure you are NOT running another instance of the server, since two applications can not bind to the same port. If problems persist, you can try changing the port.", vbOKOnly

    frmMain.StatusLbl.Caption = "Loaded!"
    
    Exit Sub
    
ErrOut:

    'Quit if we have no files to update
    MsgBox "Error: You must include files to update to run the server!" & vbCrLf & _
        "Place them in the following path:" & vbCrLf & vbCrLf & _
        App.Path & "\UpdateFiles\", vbOKOnly
    Unload frmMain
    Exit Sub

End Sub

Public Sub GOREsock_Close(inSox As Long)
Dim NumConnections As Integer
Dim i As Long

    'Remove the user's ConnID if they are on the socket that just closed
    For i = 1 To UBound(UserList)
        If UserList(i).ConnID = inSox Then
            UserList(i).ConnID = 0
        Else
            If UserList(i).ConnID > 0 Then
                NumConnections = NumConnections + 1
            End If
        End If
    Next i
    
    'Refresh the "Connections established"
    frmMain.ConnectionsEstablishedLbl.Caption = NumConnections

End Sub

Public Sub GOREsock_Connection(inSox As Long)

    'Assign the inSox ID as the ConnID
    UserList(inSox).ConnID = inSox

    'Start the transfering with the first file
    UserList(inSox).LastBytesSent = 0

    'Raise connections count
    ConnectionsEst = ConnectionsEst + 1
    frmMain.ConnectionsEstablishedLbl.Caption = ConnectionsEst
    
    'Send the file list
    DoEvents
    GOREsock_SendData inSox, ServerFileList()

End Sub

Public Sub GOREsock_DataArrival(inSox As Long, inData() As Byte)

'*********************************************
'Handle data received from the client
'*********************************************
Dim Data As String
Dim ReqFileNum As Long
Dim FileNum As Byte
Dim l As Long
Dim b() As Byte

    Data = StrConv(inData, vbUnicode)
    
    'Check for a file request
    If Left$(Data, 9) = "***GET***" Then
        ReqFileNum = Val(Right$(Data, Len(Data) - 9))
        
        'Check for a valid requested file
        If ReqFileNum < 0 Then Exit Sub
        If ReqFileNum > NumFiles Then Exit Sub
        
        'Open the file
        FileNum = FreeFile
        Open CompressPath & FileListShortName(ReqFileNum) For Binary Access Read As #FileNum
            l = LOF(FileNum)    'Get the size of the file
            ReDim b(0 To l + 8) '+ 8 is 9 - 1, -1 because we start at index 0
            Get #FileNum, , b   'Grab the whole file
        Close #FileNum
        
        'Update the information
        frmMain.FileNameLbl.Caption = FileListShortName(ReqFileNum)
        frmMain.ClientIPLbl.Caption = GOREsock_Address(inSox)
        frmMain.ClientIDLbl.Caption = inSox
    
        'Add the EOF tag
        CopyMemory b(l), EOFTag(0), 9
        
        'Send the file
        GOREsock_SendData inSox, b()
        
    End If

End Sub

Public Sub GOREsock_SendComplete(inSox As Long)

    'Clear information
    UserList(inSox).LastBytesSent = 0
    frmMain.FileNameLbl.Caption = "None"
    frmMain.PercentLbl.Caption = "0%"
    frmMain.ClientIPLbl.Caption = "0.0.0.0"
    frmMain.ClientIDLbl.Caption = "0"
    
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
    frmMain.PercentLbl.Caption = Round(bytesSent / (bytesSent + bytesRemaining), 2) * 100 & "%"

    'Update the bytes remaining info
    UploadCount = UploadCount + (bytesSent - UserList(inSox).LastBytesSent)
    UserList(inSox).LastBytesSent = bytesSent

End Sub

Public Function Engine_FileExist(File As String, FileType As VbFileAttribute) As Boolean

    'Checks if a file exists
    Engine_FileExist = (LenB(Dir$(File, FileType)) <> 0)

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
