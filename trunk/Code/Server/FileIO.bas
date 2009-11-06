Attribute VB_Name = "FileIO"
Option Explicit

'Want to remove the log lines? Just run ToolLogRemover.exe! Be sure to back up, since you can not undo it!          '//\\LOGLINE//\\
Public Enum LogType                                                                                                 '//\\LOGLINE//\\
    'Uncategorized information - it is best that you don't use this since you want to keep everything               '//\\LOGLINE//\\
    ' categorized for easy searching (no one wants to dig through a million lines to find one thing)                '//\\LOGLINE//\\
    General = 0                                                                                                     '//\\LOGLINE//\\
    'Tracking what code was called last - very useful for finding what caused your application to crash             '//\\LOGLINE//\\
    CodeTracker = 1                                                                                                 '//\\LOGLINE//\\
    'Printing the incoming packets and how they are being handled                                                   '//\\LOGLINE//\\
    PacketIn = 2                                                                                                    '//\\LOGLINE//\\
    'Printing the outgoing packets and how they are put together / who they are sent to                             '//\\LOGLINE//\\
    PacketOut = 3                                                                                                   '//\\LOGLINE//\\
    'Critical errors that should really be looked at - this often contains incorrect / invalid usage                '//\\LOGLINE//\\
    ' of the vbGORE engine                                                                                          '//\\LOGLINE//\\
    CriticalError = 4                                                                                               '//\\LOGLINE//\\
    'Packet data received, but did not work with the routine - common with either incorrect packet offsets          '//\\LOGLINE//\\
    ' (packet handling is wrong) or packet hacking (people sending custom packets)                                  '//\\LOGLINE//\\
    InvalidPacketData = 5                                                                                           '//\\LOGLINE//\\
End Enum                                                                                                            '//\\LOGLINE//\\
#If False Then                                                                                                      '//\\LOGLINE//\\
Private General, CodeTracker, PacketIn, PacketOut, CriticalError, InvalidPacketData                                 '//\\LOGLINE//\\
#End If                                                                                                             '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
Public LogFileNumGeneral As Byte                                                                                    '//\\LOGLINE//\\
Public LogFileNumCodeTracker As Byte                                                                                '//\\LOGLINE//\\
Public LogFileNumPacketIn As Byte                                                                                   '//\\LOGLINE//\\
Public LogFileNumPacketOut As Byte                                                                                  '//\\LOGLINE//\\
Public LogFileNumCriticalError As Byte                                                                              '//\\LOGLINE//\\
Public LogFileNumInvalidPacketData As Byte                                                                          '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
'How much of the file we preserve when cropping (starting from the end and working backwords)                       '//\\LOGLINE//\\
Private Const MinLogFileSize As Long = 5242880   '5 MB                                                              '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
'How large the file must be before we crop it - recommended you keep a decent value between the min and max         '//\\LOGLINE//\\
' values since the cropping routine can be pretty slow                                                              '//\\LOGLINE//\\
Private Const MaxLogFileSize As Long = 10485760  '10 MB                                                             '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
Public Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long             '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
Public Sub Log(ByVal Text As String, ByVal LogType As LogType)                                                      '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
'*****************************************************************                                                  '//\\LOGLINE//\\
'Logs data for finding errors                                                                                       '//\\LOGLINE//\\
'*****************************************************************                                                  '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
'Check if we are using logging                                                                                      '//\\LOGLINE//\\
If Not DEBUG_UseLogging Then Exit Sub                                                                               '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
Dim LogFile As String   'Path to our log file (depends on LogType)                                                  '//\\LOGLINE//\\
Dim b() As Byte         'Used for cropping down the file if it gets too large                                       '//\\LOGLINE//\\
Dim C() As Byte         'The cropped down version of b()                                                            '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
    'We put the line break on here because it'd be soooo tedious and worthless to write it for every log call       '//\\LOGLINE//\\
    Text = Text & vbNewLine                                                                                         '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
    'Define the log file path according to the log type                                                             '//\\LOGLINE//\\
    Select Case LogType                                                                                             '//\\LOGLINE//\\
        Case General                                                                                                '//\\LOGLINE//\\
            If LogFileNumGeneral = 0 Then                                                                           '//\\LOGLINE//\\
                LogFile = LogPath & ServerID & "\General.log"                                                       '//\\LOGLINE//\\
                If LenB(Dir$(LogFile, vbNormal)) Then Kill LogFile                                                  '//\\LOGLINE//\\
                MakeSureDirectoryPathExists LogFile                                                                 '//\\LOGLINE//\\
                LogFileNumGeneral = FreeFile                                                                        '//\\LOGLINE//\\
                Open LogFile For Binary As #LogFileNumGeneral                                                       '//\\LOGLINE//\\
            End If                                                                                                  '//\\LOGLINE//\\
            Put #LogFileNumGeneral, , Text                                                                          '//\\LOGLINE//\\
            If LOF(LogFileNumGeneral) > MaxLogFileSize Then                                                         '//\\LOGLINE//\\
                LogFile = LogPath & ServerID & "\General.log"                                                       '//\\LOGLINE//\\
                Seek #LogFileNumGeneral, 1                                                                          '//\\LOGLINE//\\
                ReDim b(LOF(LogFileNumGeneral))                                                                     '//\\LOGLINE//\\
                ReDim C(MinLogFileSize)                                                                             '//\\LOGLINE//\\
                Get #LogFileNumGeneral, , b                                                                         '//\\LOGLINE//\\
                CopyMemory C(0), b(LOF(LogFileNumGeneral) - MinLogFileSize), MinLogFileSize                         '//\\LOGLINE//\\
                Close #LogFileNumGeneral                                                                            '//\\LOGLINE//\\
                Kill LogFile                                                                                        '//\\LOGLINE//\\
                Open LogFile For Binary As #LogFileNumGeneral                                                       '//\\LOGLINE//\\
                Put #LogFileNumGeneral, , C                                                                         '//\\LOGLINE//\\
            End If                                                                                                  '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
        Case CodeTracker                                                                                            '//\\LOGLINE//\\
            If LogFileNumCodeTracker = 0 Then                                                                       '//\\LOGLINE//\\
                LogFile = LogPath & ServerID & "\CodeTracker.log"                                                   '//\\LOGLINE//\\
                If LenB(Dir$(LogFile, vbNormal)) Then Kill LogFile                                                  '//\\LOGLINE//\\
                MakeSureDirectoryPathExists LogFile                                                                 '//\\LOGLINE//\\
                LogFileNumCodeTracker = FreeFile                                                                    '//\\LOGLINE//\\
                Open LogFile For Binary As #LogFileNumCodeTracker                                                   '//\\LOGLINE//\\
            End If                                                                                                  '//\\LOGLINE//\\
            Put #LogFileNumCodeTracker, , Text                                                                      '//\\LOGLINE//\\
            If LOF(LogFileNumCodeTracker) > MaxLogFileSize Then                                                     '//\\LOGLINE//\\
                LogFile = LogPath & ServerID & "\CodeTracker.log"                                                   '//\\LOGLINE//\\
                Seek #LogFileNumCodeTracker, 1                                                                      '//\\LOGLINE//\\
                ReDim b(LOF(LogFileNumCodeTracker))                                                                 '//\\LOGLINE//\\
                ReDim C(MinLogFileSize)                                                                             '//\\LOGLINE//\\
                Get #LogFileNumCodeTracker, , b                                                                     '//\\LOGLINE//\\
                CopyMemory C(0), b(LOF(LogFileNumCodeTracker) - MinLogFileSize), MinLogFileSize                     '//\\LOGLINE//\\
                Close #LogFileNumCodeTracker                                                                        '//\\LOGLINE//\\
                Kill LogFile                                                                                        '//\\LOGLINE//\\
                Open LogFile For Binary As #LogFileNumCodeTracker                                                   '//\\LOGLINE//\\
                Put #LogFileNumCodeTracker, , C                                                                     '//\\LOGLINE//\\
            End If                                                                                                  '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
        Case PacketIn                                                                                               '//\\LOGLINE//\\
            If LogFileNumPacketIn = 0 Then                                                                          '//\\LOGLINE//\\
                LogFile = LogPath & ServerID & "\PacketIn.log"                                                      '//\\LOGLINE//\\
                If LenB(Dir$(LogFile, vbNormal)) Then Kill LogFile                                                  '//\\LOGLINE//\\
                MakeSureDirectoryPathExists LogFile                                                                 '//\\LOGLINE//\\
                LogFileNumPacketIn = FreeFile                                                                       '//\\LOGLINE//\\
                Open LogFile For Binary As #LogFileNumPacketIn                                                      '//\\LOGLINE//\\
            End If                                                                                                  '//\\LOGLINE//\\
            Put #LogFileNumPacketIn, , Text                                                                         '//\\LOGLINE//\\
            If LOF(LogFileNumPacketIn) > MaxLogFileSize Then                                                        '//\\LOGLINE//\\
                LogFile = LogPath & ServerID & "\PacketIn.log"                                                      '//\\LOGLINE//\\
                Seek #LogFileNumPacketIn, 1                                                                         '//\\LOGLINE//\\
                ReDim b(LOF(LogFileNumPacketIn))                                                                    '//\\LOGLINE//\\
                ReDim C(MinLogFileSize)                                                                             '//\\LOGLINE//\\
                Get #LogFileNumPacketIn, , b                                                                        '//\\LOGLINE//\\
                CopyMemory C(0), b(LOF(LogFileNumPacketIn) - MinLogFileSize), MinLogFileSize                        '//\\LOGLINE//\\
                Close #LogFileNumPacketIn                                                                           '//\\LOGLINE//\\
                Kill LogFile                                                                                        '//\\LOGLINE//\\
                Open LogFile For Binary As #LogFileNumPacketIn                                                      '//\\LOGLINE//\\
                Put #LogFileNumPacketIn, , C                                                                        '//\\LOGLINE//\\
            End If                                                                                                  '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
        Case PacketOut                                                                                              '//\\LOGLINE//\\
            If LogFileNumPacketOut = 0 Then                                                                         '//\\LOGLINE//\\
                LogFile = LogPath & ServerID & "\PacketOut.log"                                                     '//\\LOGLINE//\\
                If LenB(Dir$(LogFile, vbNormal)) Then Kill LogFile                                                  '//\\LOGLINE//\\
                MakeSureDirectoryPathExists LogFile                                                                 '//\\LOGLINE//\\
                LogFileNumPacketOut = FreeFile                                                                      '//\\LOGLINE//\\
                Open LogFile For Binary As #LogFileNumPacketOut                                                     '//\\LOGLINE//\\
            End If                                                                                                  '//\\LOGLINE//\\
            Put #LogFileNumPacketOut, , Text                                                                        '//\\LOGLINE//\\
            If LOF(LogFileNumPacketOut) > MaxLogFileSize Then                                                       '//\\LOGLINE//\\
                LogFile = LogPath & ServerID & "\PacketOut.log"                                                     '//\\LOGLINE//\\
                Seek #LogFileNumPacketOut, 1                                                                        '//\\LOGLINE//\\
                ReDim b(LOF(LogFileNumPacketOut))                                                                   '//\\LOGLINE//\\
                ReDim C(MinLogFileSize)                                                                             '//\\LOGLINE//\\
                Get #LogFileNumPacketOut, , b                                                                       '//\\LOGLINE//\\
                CopyMemory C(0), b(LOF(LogFileNumPacketOut) - MinLogFileSize), MinLogFileSize                       '//\\LOGLINE//\\
                Close #LogFileNumPacketOut                                                                          '//\\LOGLINE//\\
                Kill LogFile                                                                                        '//\\LOGLINE//\\
                Open LogFile For Binary As #LogFileNumPacketOut                                                     '//\\LOGLINE//\\
                Put #LogFileNumPacketOut, , C                                                                       '//\\LOGLINE//\\
            End If                                                                                                  '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
        Case CriticalError                                                                                          '//\\LOGLINE//\\
            If LogFileNumCriticalError = 0 Then                                                                     '//\\LOGLINE//\\
                LogFile = LogPath & ServerID & "\CriticalError.log"                                                 '//\\LOGLINE//\\
                If LenB(Dir$(LogFile, vbNormal)) Then Kill LogFile                                                  '//\\LOGLINE//\\
                MakeSureDirectoryPathExists LogFile                                                                 '//\\LOGLINE//\\
                LogFileNumCriticalError = FreeFile                                                                  '//\\LOGLINE//\\
                Open LogFile For Binary As #LogFileNumCriticalError                                                 '//\\LOGLINE//\\
            End If                                                                                                  '//\\LOGLINE//\\
            Put #LogFileNumCriticalError, , Text                                                                    '//\\LOGLINE//\\
            If LOF(LogFileNumCriticalError) > MaxLogFileSize Then                                                   '//\\LOGLINE//\\
                LogFile = LogPath & ServerID & "\CriticalError.log"                                                 '//\\LOGLINE//\\
                Seek #LogFileNumCriticalError, 1                                                                    '//\\LOGLINE//\\
                ReDim b(LOF(LogFileNumCriticalError))                                                               '//\\LOGLINE//\\
                ReDim C(MinLogFileSize)                                                                             '//\\LOGLINE//\\
                Get #LogFileNumCriticalError, , b                                                                   '//\\LOGLINE//\\
                CopyMemory C(0), b(LOF(LogFileNumCriticalError) - MinLogFileSize), MinLogFileSize                   '//\\LOGLINE//\\
                Close #LogFileNumCriticalError                                                                      '//\\LOGLINE//\\
                Kill LogFile                                                                                        '//\\LOGLINE//\\
                Open LogFile For Binary As #LogFileNumCriticalError                                                 '//\\LOGLINE//\\
                Put #LogFileNumCriticalError, , C                                                                   '//\\LOGLINE//\\
            End If                                                                                                  '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
        Case InvalidPacketData                                                                                      '//\\LOGLINE//\\
            If LogFileNumInvalidPacketData = 0 Then                                                                 '//\\LOGLINE//\\
                LogFile = LogPath & ServerID & "\InvalidPacketData.log"                                             '//\\LOGLINE//\\
                If LenB(Dir$(LogFile, vbNormal)) Then Kill LogFile                                                  '//\\LOGLINE//\\
                MakeSureDirectoryPathExists LogFile                                                                 '//\\LOGLINE//\\
                LogFileNumInvalidPacketData = FreeFile                                                              '//\\LOGLINE//\\
                Open LogFile For Binary As #LogFileNumInvalidPacketData                                             '//\\LOGLINE//\\
            End If                                                                                                  '//\\LOGLINE//\\
            Put #LogFileNumInvalidPacketData, , Text                                                                '//\\LOGLINE//\\
            If LOF(LogFileNumInvalidPacketData) > MaxLogFileSize Then                                               '//\\LOGLINE//\\
                LogFile = LogPath & ServerID & "\InvalidPacketData.log"                                             '//\\LOGLINE//\\
                Seek #LogFileNumInvalidPacketData, 1                                                                '//\\LOGLINE//\\
                ReDim b(LOF(LogFileNumInvalidPacketData))                                                           '//\\LOGLINE//\\
                ReDim C(MinLogFileSize)                                                                             '//\\LOGLINE//\\
                Get #LogFileNumInvalidPacketData, , b                                                               '//\\LOGLINE//\\
                CopyMemory C(0), b(LOF(LogFileNumInvalidPacketData) - MinLogFileSize), MinLogFileSize               '//\\LOGLINE//\\
                Close #LogFileNumInvalidPacketData                                                                  '//\\LOGLINE//\\
                Kill LogFile                                                                                        '//\\LOGLINE//\\
                Open LogFile For Binary As #LogFileNumInvalidPacketData                                             '//\\LOGLINE//\\
                Put #LogFileNumInvalidPacketData, , C                                                               '//\\LOGLINE//\\
            End If                                                                                                  '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
    End Select                                                                                                      '//\\LOGLINE//\\
                                                                                                                    '//\\LOGLINE//\\
End Sub                                                                                                             '//\\LOGLINE//\\

Public Function Load_Mail(ByVal MailIndex As Long) As MailData
Dim DataSplit() As String
Dim ObjSplit() As String
Dim ObjStr As String
Dim i As Long

    Log "Call Load_Mail(" & MailIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Open the database
    DB_RS.Open "SELECT * FROM mail WHERE id=" & MailIndex, DB_Conn, adOpenStatic, adLockOptimistic
    
    'Make sure we have a valid mail index
    If Not DB_RS.EOF Then
        
        'Apply the values
        Load_Mail.Subject = Trim$(DB_RS!Sub)
        Load_Mail.WriterName = Trim$(DB_RS!By)
        Load_Mail.RecieveDate = DB_RS!Date
        Load_Mail.Message = Trim$(DB_RS!Msg)
        Load_Mail.New = Val(DB_RS!New)
        ObjStr = Trim$(DB_RS!Objs)
    
        'Check for a valid object string
        If LenB(ObjStr) Then
        
            Log "Load_Mail: Splitting ObjStr (" & ObjStr & ")", CodeTracker '//\\LOGLINE//\\
        
            'Split the objects up from the object string
            ObjSplit = Split(ObjStr, vbNewLine)

            'Loop through the objects
            For i = 0 To UBound(ObjSplit)
            
                Log "Load_Mail: Splitting object data (" & ObjSplit(i) & ")", CodeTracker '//\\LOGLINE//\\
            
                'Split up the index and amount
                DataSplit = Split(ObjSplit(i), " ")
                
                'Set the data
                Load_Mail.Obj(i + 1).ObjIndex = Val(DataSplit(0))
                Load_Mail.Obj(i + 1).Amount = Val(DataSplit(1))
                
            Next i
        
        End If
        
    End If
   
    'Close the database
    DB_RS.Close

End Function

Public Sub Load_Maps_Temp(ByVal MapNum As Integer)
'*****************************************************************
'Take the bulk temp map dump and load it instead of using the compressed
' (and in result, slower) load map system from Load_Maps_Data
'*****************************************************************
Dim NPCInfo() As NPCLoadData
Dim CharIndex As Integer
Dim NPCIndex As Integer
Dim intNumNPCs As Integer
Dim FileNum As Byte
Dim i As Long

    'Don't load a loaded map
    If MapInfo(MapNum).DataLoaded = 1 Then Exit Sub
    
    'Set the data as loaded
    MapInfo(MapNum).DataLoaded = 1

    'Create the data arrays
    ReDim MapInfo(MapNum).Data(1 To CLng(MapInfo(MapNum).Width), 1 To CLng(MapInfo(MapNum).Height))
    ReDim MapInfo(MapNum).ObjTile(1 To CLng(MapInfo(MapNum).Width), 1 To CLng(MapInfo(MapNum).Height))

    'Open the file
    FileNum = FreeFile
    Open ServerTempPath & "m" & MapNum & ".temp" For Binary Access Read As #FileNum
    
        'Get the NPC information
        Get #FileNum, , intNumNPCs
        If intNumNPCs > 0 Then
            ReDim NPCInfo(1 To intNumNPCs) As NPCLoadData
            Get #FileNum, , NPCInfo()
        End If

        'Get the tile information
        Get #FileNum, , MapInfo(MapNum).Data()
        
    'Close up
    Close #FileNum
    
    'Load the NPCs
    If intNumNPCs > 0 Then
        For i = 1 To intNumNPCs
        
            NPCIndex = Load_NPC(NPCInfo(i).NPCNum)
            
            With NPCList(NPCIndex)
            
                'Create the NPC
                .Pos.Map = MapNum
                .Pos.X = NPCInfo(i).X
                .Pos.Y = NPCInfo(i).Y
                .StartPos = .Pos
    
                'Give the NPC a char index
                CharIndex = Server_NextOpenCharIndex
                .Char.CharIndex = CharIndex
                CharList(CharIndex).Index = NPCIndex
                CharList(CharIndex).CharType = CharType_NPC
                
                'Set the NPC as used
                .Flags.NPCActive = 1
            
            End With
        Next i
    End If

End Sub

Private Sub Save_Maps_Temp(ByVal MapNum As Integer)

'*****************************************************************
'Take the data from a loaded map and saves it in a bulky yet fast access storage
'This is used to make on-the-fly map loading much faster
'*****************************************************************
Dim NPCInfo() As NPCLoadData
Dim intNumNPCs As Integer
Dim FileNum As Byte
Dim X As Long
Dim Y As Long

    'Delete any existing temp file
    If Server_FileExist(ServerTempPath & "m" & MapNum & ".temp", vbNormal) Then Kill ServerTempPath & "m" & MapNum & ".temp"

    'Count and store the NPCs (then clear them off)
    For X = 1 To MapInfo(MapNum).Width
        For Y = 1 To MapInfo(MapNum).Height
            If MapInfo(MapNum).Data(X, Y).NPCIndex > 0 Then
                
                'Raise the NPC count and store the information
                intNumNPCs = intNumNPCs + 1
                If intNumNPCs = 1 Then ReDim NPCInfo(1 To intNumNPCs) Else ReDim Preserve NPCInfo(1 To intNumNPCs)
                With NPCInfo(intNumNPCs)
                    .NPCNum = MapInfo(MapNum).Data(X, Y).NPCIndex
                    .X = X
                    .Y = Y
                End With

            End If
            
            'Clear off the variables that need to be removed
            MapInfo(MapNum).Data(X, Y).NPCIndex = 0
            MapInfo(MapNum).Data(X, Y).UserIndex = 0
            
        Next Y
    Next X
    
    'Open the file
    FileNum = FreeFile
    Open ServerTempPath & "m" & MapNum & ".temp" For Binary Access Write As #FileNum
        
        'Store the NPCs
        Put #FileNum, , intNumNPCs
        If intNumNPCs > 0 Then Put #FileNum, , NPCInfo
    
        'Store the tile information
        Put #FileNum, , MapInfo(MapNum).Data()
        
    'Close up
    Close #FileNum

End Sub

Private Sub Load_Maps_Data(ByVal MapNum As Integer)

'*****************************************************************
'Loads the Data() of a map (holds the tile data)
'*****************************************************************
Dim FileNumMap As Byte
Dim FileNumInf As Byte
Dim TempLng As Long
Dim TempInt As Integer
Dim ByFlags As Long
Dim BxFlags As Byte
Dim X As Long
Dim Y As Long
Dim i As Long

    Log "Call Load_Maps_Data(" & MapNum & ")", CodeTracker '//\\LOGLINE//\\
    
    'Load the files
    'Map
    FileNumMap = FreeFile
    Open MapPath & MapNum & ".map" For Binary Access Read As #FileNumMap
    Seek #FileNumMap, 1

    'Inf
    FileNumInf = FreeFile
    Open MapEXPath & MapNum & ".inf" For Binary Access Read As #FileNumInf
    Seek #FileNumInf, 1

    'Map header
    Get #FileNumMap, , MapInfo(MapNum).MapVersion
    Get #FileNumMap, , MapInfo(MapNum).Width
    Get #FileNumMap, , MapInfo(MapNum).Height
    
    'Create the array
    ReDim MapInfo(MapNum).Data(1 To CLng(MapInfo(MapNum).Width), 1 To CLng(MapInfo(MapNum).Height))

    'Load arrays
    For Y = 1 To MapInfo(MapNum).Height
        For X = 1 To MapInfo(MapNum).Width

            'Get tile's flags
            Get #FileNumMap, , ByFlags

            'Blocked
            MapInfo(MapNum).Data(X, Y).Blocked = 0
            If ByFlags And 1 Then Get #FileNumMap, , MapInfo(MapNum).Data(X, Y).Blocked

            'Graphic layers (values dont need to be stored)
            If ByFlags And 2 Then Get #FileNumMap, , TempLng
            If ByFlags And 4 Then Get #FileNumMap, , TempLng
            If ByFlags And 8 Then Get #FileNumMap, , TempLng
            If ByFlags And 16 Then Get #FileNumMap, , TempLng
            If ByFlags And 32 Then Get #FileNumMap, , TempLng
            If ByFlags And 64 Then Get #FileNumMap, , TempLng

            'Get lighting values (values dont need to be stored)
            If ByFlags And 128 Then
                For i = 1 To 4
                    Get #FileNumMap, , TempLng
                Next i
            End If
            If ByFlags And 256 Then
                For i = 5 To 8
                    Get #FileNumMap, , TempLng
                Next i
            End If
            If ByFlags And 512 Then
                For i = 9 To 12
                    Get #FileNumMap, , TempLng
                Next i
            End If
            If ByFlags And 1024 Then
                For i = 13 To 16
                    Get #FileNumMap, , TempLng
                Next i
            End If
            If ByFlags And 2048 Then
                For i = 17 To 20
                    Get #FileNumMap, , TempLng
                Next i
            End If
            If ByFlags And 4096 Then
                For i = 21 To 24
                    Get #FileNumMap, , TempLng
                Next i
            End If

            'Mailbox
            If ByFlags And 8192 Then MapInfo(MapNum).Data(X, Y).Mailbox = 1 Else MapInfo(MapNum).Data(X, Y).Mailbox = 0
            
            'Sfx (value doesn't need to be stored)
            If ByFlags And 1048576 Then Get #FileNumMap, , TempInt
            
            'Blocked attack (value stuck into the Blocked flag to save RAM)
            If ByFlags And 2097152 Then MapInfo(MapNum).Data(X, Y).Blocked = MapInfo(MapNum).Data(X, Y).Blocked Or 128
            
            'Sign (value doesn't need to be stored)
            If ByFlags And 4194304 Then Get #FileNumMap, , TempInt
            
            '.inf file

            'Get flag's byte
            Get #FileNumInf, , BxFlags

            'Load Tile Exit
            If BxFlags And 1 Then
                With MapInfo(MapNum).Data(X, Y)
                    Get #FileNumInf, , .TileExitMap
                    Get #FileNumInf, , .TileExitX
                    Get #FileNumInf, , .TileExitY
                End With
            End If

            'Load NPC
            If BxFlags And 2 Then
                Get #FileNumInf, , TempInt
                MapInfo(MapNum).Data(X, Y).NPCIndex = TempInt
            End If
            
        Next X
    Next Y

    'Close files
    Close #FileNumMap
    Close #FileNumInf

End Sub

Sub Unload_Map(ByVal MapNum As Integer)

'*****************************************************************
'Unloads the map data from memory, and any NPCs and objects on it
'*****************************************************************
Dim i As Long

    'Don't unload an unloaded map
    If MapInfo(MapNum).DataLoaded = 1 Then
    
        'Check the map life time
        If MapInfo(MapNum).UnloadTimer = 0 Then
            MapInfo(MapNum).UnloadTimer = EmptyMapLife + timeGetTime
            Exit Sub
        End If
        
        'Check if to remove the map
        If MapInfo(MapNum).UnloadTimer + EmptyMapLife < timeGetTime Then

            'Set the map as unloaded
            MapInfo(MapNum).DataLoaded = 0
            MapInfo(MapNum).UnloadTimer = 0
        
            'Unload all the NPCs on the map
            For i = 1 To LastNPC
                With NPCList(i)
                    If .Pos.Map = MapNum Then
                        CharList(.Char.CharIndex).Index = 0
                        CharList(.Char.CharIndex).CharType = 0
                        .Flags.NPCActive = 0
                        NPC_Close i, 0
                    End If
                End With
            Next i
            
            'Clean the NPC array
            NPC_CleanArray
        
            'Completely unload the map data
            Erase MapInfo(MapNum).Data()
            
        End If
            
    End If

End Sub

Public Sub Load_Maps()

'*****************************************************************
'Loads the MapX.X files
'*****************************************************************
Dim LoopC As Long
Dim Map As Long

    Log "Call Load_Maps", CodeTracker '//\\LOGLINE//\\

    NumMaps = Val(Var_Get(DataPath & "Map.dat", "INIT", "NumMaps"))
    ReDim MapInfo(1 To NumMaps)

    'Create MapUsers
    ReDim MapUsers(1 To NumMaps)
    For LoopC = 1 To NumMaps
        ReDim MapUsers(LoopC).Index(0)
    Next LoopC
    
    'Load the server settings (this has to be done right here)
    Load_ServerIni

    For Map = 1 To NumMaps
    
        Log "Load_Maps: Loading map (" & Map & ")", CodeTracker '//\\LOGLINE//\\

        If Server_FileExist(MapPath & Map & ".map", vbNormal) Then
            If Server_FileExist(MapEXPath & Map & ".dat", vbNormal) Then
                If Server_FileExist(MapEXPath & Map & ".inf", vbNormal) Then
                
                    'Other Room Data
                    With MapInfo(Map)
                        .Name = Var_Get(MapEXPath & Map & ".dat", "1", "Name")
                        .Weather = Val(Var_Get(MapEXPath & Map & ".dat", "1", "Weather"))
                        .Music = Val(Var_Get(MapEXPath & Map & ".dat", "1", "Music"))
                    End With
                    
                    'Create the temp maps
                    Load_Maps_Data Map
                    Save_Maps_Temp Map
                    
                End If
            End If
        End If
        
    Next Map
    
End Sub

Public Sub Save_NPCs_Temp()

'*****************************************************************
'Creates and saves the .temp NPCs
'*****************************************************************
Dim ObjNums As NPCBytes
Dim FileNum As Byte
Dim ShopStr As String
Dim DropStr As String
Dim ItemSplit() As String
Dim TempSplit() As String
Dim i As Long
Dim j As Byte

    Log "Call Save_NPCs_Temp", CodeTracker '//\\LOGLINE//\\
    
    'Resize the NPC array to fit the one NPC we are using
    ReDim NPCList(1 To 1) As NPC

    'Grab all the NPCs from the database
    DB_RS.Open "SELECT * FROM npcs", DB_Conn, adOpenStatic, adLockOptimistic
    
    'Loop through them, and put the data into the NPCList(1)
    Do While Not DB_RS.EOF  'Loop until we reach the end of the recordset
        
        'Delete any existing temp file
        If Server_FileExist(ServerTempPath & "n" & DB_RS!id & ".temp", vbNormal) Then Kill ServerTempPath & "n" & DB_RS!id & ".temp"
    
        'Clear the variables so nothing gets transferred over to the next NPC
        Erase NPCList(1).VendItems
        Erase NPCList(1).DropItems
        Erase NPCList(1).DropRate
        ZeroMemory NPCList(1), Len(NPCList(1))
        i = 0
        ObjNums.Drop = 0
        ObjNums.Vend = 0
    
        With NPCList(1)
            Log "Save_NPCs_Temp: Filling in values for NPC " & DB_RS!id, CodeTracker '//\\LOGLINE//\\
            .Name = Trim$(DB_RS!Name)
            .Desc = Trim$(DB_RS!Descr)
            .AttackGrh = Val(DB_RS!AttackGrh)
            .AttackRange = Val(DB_RS!AttackRange)
            .AttackSfx = Val(DB_RS!AttackSfx)
            .AI = Val(DB_RS!AI)
            .ChatID = Val(DB_RS!Chat)
            .RespawnWait = Val(DB_RS!RespawnWait)
            .ProjectileRotateSpeed = Val(DB_RS!ProjectileRotateSpeed)
            .Attackable = Val(DB_RS!Attackable)
            .Hostile = Val(DB_RS!Hostile)
            .Quest = Val(DB_RS!Quest)
            .GiveEXP = Val(DB_RS!give_exp)
            .GiveGLD = Val(DB_RS!give_gold)
            .Char.Hair = Val(DB_RS!char_hair)
            .Char.Head = Val(DB_RS!char_head)
            .Char.Body = Val(DB_RS!char_body)
            .Char.Weapon = Val(DB_RS!char_weapon)
            .Char.Wings = Val(DB_RS!char_wings)
            .Char.Heading = Val(DB_RS!char_heading)
            .Char.HeadHeading = Val(DB_RS!char_headheading)
            .BaseStat(SID.Agi) = Val(DB_RS!stat_hitrate)
            .BaseStat(SID.Speed) = Val(DB_RS!stat_speed)
            .BaseStat(SID.Mag) = Val(DB_RS!stat_mag)
            .BaseStat(SID.DEF) = Val(DB_RS!stat_def)
            .BaseStat(SID.MinHIT) = Val(DB_RS!stat_hit_min)
            .BaseStat(SID.MaxHIT) = Val(DB_RS!stat_hit_max)
            .BaseStat(SID.MaxHP) = Val(DB_RS!stat_hp)
            .BaseStat(SID.MaxMAN) = Val(DB_RS!stat_mp)
            .BaseStat(SID.MaxSTA) = Val(DB_RS!stat_sp)
            .BaseStat(SID.MinHP) = .BaseStat(SID.MaxHP)
            .BaseStat(SID.MinMAN) = .BaseStat(SID.MaxMAN)
            .BaseStat(SID.MinSTA) = .BaseStat(SID.MaxSTA)
            .NPCNumber = DB_RS!id
            .Flags.NPCActive = 1
            ShopStr = Trim$(DB_RS!objs_shop)
            DropStr = Trim$(DB_RS!drops)

            'Create the shop list
            If LenB(ShopStr) Then
                Log "Load_NPC: Splitting ShopStr (" & ShopStr & ")", CodeTracker '//\\LOGLINE//\\
                TempSplit = Split(ShopStr, vbNewLine)
                j = UBound(TempSplit)   'Cache the ubound - it is much faster to cache it then call UBound twice or more!
                ReDim .VendItems(1 To j + 1)
                .NumVendItems = j + 1
                For i = 0 To j
                    Log "Save_NPCs_Temp: Splitting item information (" & TempSplit(i) & ")", CodeTracker '//\\LOGLINE//\\
                    ItemSplit = Split(Trim$(TempSplit(i)), " ")
                    If UBound(ItemSplit) = 1 Then   'If ubound <> 1, we have an invalid item entry
                        .VendItems(i + 1).ObjIndex = Val(ItemSplit(0))
                        .VendItems(i + 1).Amount = Val(ItemSplit(1))
                    Else
                        Log "Save_NPCs_Temp: Invalid shop/vending item entry found in the database. NPC: " & DB_RS!id & " Slot: " & i, CriticalError '//\\LOGLINE//\\
                    End If
                Next i
            End If
            
            'Create the drop list
            If LenB(DropStr) Then
                Log "Load_NPC: Splitting DropStr (" & DropStr & ")", CodeTracker '//\\LOGLINE//\\
                TempSplit = Split(Trim$(DropStr), vbNewLine)
                j = UBound(TempSplit)
                .NumDropItems = j + 1
                ReDim .DropItems(1 To .NumDropItems)
                ReDim .DropRate(1 To .NumDropItems)
                For i = 0 To j
                    Log "Save_NPCs_Temp: Splitting item information (" & TempSplit(i) & ")", CodeTracker '//\\LOGLINE//\\
                    ItemSplit = Split(Trim$(TempSplit(i)), " ")
                    If UBound(ItemSplit) = 2 Then   'If ubound <> 2, we have an invalid item entry
                        .DropItems(i + 1).ObjIndex = Val(ItemSplit(0))
                        .DropItems(i + 1).Amount = Val(ItemSplit(1))
                        .DropRate(i + 1) = Val(ItemSplit(2))
                    Else
                        Log "Save_NPCs_Temp: Invalid drop item entry found in the database. NPC: " & DB_RS!id & " Slot: " & i, CriticalError '//\\LOGLINE//\\
                    End If
                Next i
            End If
            
            'Put the values into the ObjNums
            ObjNums.Drop = NPCList(1).NumDropItems
            ObjNums.Vend = NPCList(1).NumVendItems
            
            'Finally, update the NPC's mod stats
            NPC_UpdateModStats 1

            'Save the NPCs to the file
            FileNum = FreeFile
            Open ServerTempPath & "n" & DB_RS!id & ".temp" For Binary Access Write As #FileNum
                
                'Array sizes
                Put #FileNum, , ObjNums

                'The NPC itself
                Put #FileNum, , NPCList(1)
                
            Close #FileNum
        
        End With
        DB_RS.MoveNext
    Loop
    
    'Close the record set
    DB_RS.Close
    
    'Clear the NPC list again
    Erase NPCList
            
End Sub

Public Sub Load_NPC_Names()

'*****************************************************************
'Loads the names of NPCs (only if they are used in a quest)
'*****************************************************************
Dim i As Long

    'Resize the NPC name array by the highest index used
    DB_RS.Open "SELECT finish_req_killnpc FROM quests ORDER BY id DESC", DB_Conn, adOpenStatic, adLockOptimistic
    If Val(DB_RS(0)) = 0 Then
        
        'No NPCs used for quests, abort
        DB_RS.Close
        Exit Sub
        
    End If
    
    ReDim NPCName(1 To DB_RS(0))
    DB_RS.Close

    'Grab all the NPC numbers used in quests
    DB_RS.Open "SELECT finish_req_killnpc FROM quests", DB_Conn, adOpenStatic, adLockOptimistic
    
    'Loop through all the IDs
    Do While Not DB_RS.EOF  'Loop until we reach the end of the recordset
        
        'If the ID is used, mark it with ".", so we can get the real name later
        If Val(DB_RS(0)) > 0 Then NPCName(Val(DB_RS(0))) = "."
        
        'Move to the next record
        DB_RS.MoveNext
        
    Loop
    
    DB_RS.Close
        
    'Fill in the values
    For i = 1 To UBound(NPCName)
        
        'A "." states we need to get the name
        If NPCName(i) = "." Then
            
            'Get the name
            DB_RS.Open "SELECT name FROM npcs WHERE id=" & i, DB_Conn, adOpenStatic, adLockOptimistic
            NPCName(i) = Trim$(DB_RS(0))
            DB_RS.Close
        
        End If
        
    Next i
        
End Sub

Public Function Load_NPC(ByVal NPCNumber As Integer, Optional ByVal Thralled As Byte = 0, Optional ByVal ThralledTime As Long = -1) As Integer

'*****************************************************************
'Loads a NPC and returns its index
'*****************************************************************
Dim ObjNums As NPCBytes
Dim NPCIndex As Integer
Dim FileNum As Byte

    Log "Call Load_NPC(" & NPCNumber & "," & Thralled & ")", CodeTracker '//\\LOGLINE//\\

    'Check for valid NPCNumber
    If NPCNumber <= 0 Then
        Log "Rtrn Load_NPC = " & Load_NPC, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    Log "Load_NPC: Acquiring NPC index", CodeTracker '//\\LOGLINE//\\

    'Find next open NPCindex
    NPCIndex = NPC_NextOpen
    Log "Load_NPC: NPC index acquired (" & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Update NPC counters
    If NPCIndex > LastNPC Then
        LastNPC = NPCIndex
        If LastNPC <> 0 Then
            Log "Load_NPC: ReDimming NPCList array with LastNPC value (" & LastNPC & ")", CodeTracker '//\\LOGLINE//\\
            ReDim Preserve NPCList(1 To LastNPC)
        End If
    End If
    NumNPCs = NumNPCs + 1

    'Make sure the NPC exists
    If Not Server_FileExist(ServerTempPath & "n" & NPCNumber & ".temp", vbNormal) Then
        If Thralled = 0 Then    'Don't give the error from an invalid thrall
            Log "Load_NPC: Error loading NPC " & NPCIndex & " with NPCNumber " & NPCNumber & " - no NPC by the number found!", CriticalError '//\\LOGLINE//\\
        End If
        Log "Rtrn Load_NPC = " & Load_NPC, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If

    With NPCList(NPCIndex)
        
        'Get the NPC information from the file
        FileNum = FreeFile
        Open ServerTempPath & "n" & NPCNumber & ".temp" For Binary Access Read As #FileNum
            
            'Get the array sizes
            Get #FileNum, , ObjNums
            
            'Set the arrays if needed
            If ObjNums.Drop > 0 Then
                ReDim .DropItems(1 To ObjNums.Drop) As Obj
                ReDim .DropRate(1 To ObjNums.Drop) As Single
            Else
                Erase .DropItems
                Erase .DropRate
            End If
            If ObjNums.Vend > 0 Then
                ReDim vendarray(1 To ObjNums.Vend) As Obj
            Else
                Erase .VendItems
            End If
            
            'Get the NPC
            Get #FileNum, , NPCList(NPCIndex)
        
            'Set the NPC's thralled value
            .Flags.Thralled = Thralled
            If ThralledTime <> -1 Then
                .Counters.RespawnCounter = timeGetTime + ThralledTime
            Else
                .Counters.RespawnCounter = -1
            End If
            
        Close #FileNum

    End With

    'Return new NPCIndex
    Load_NPC = NPCIndex
    
    Log "Rtrn Load_NPC = " & Load_NPC, CodeTracker '//\\LOGLINE//\\

End Function

Public Sub Load_OBJs()
Dim TempObjData As udtObjData
Dim FileNum As Byte

    Log "Call Load_OBJs", CodeTracker '//\\LOGLINE//\\
    
    'Get the number of objects (Sort by id, descending, only get 1 entry, only return id)
    DB_RS.Open "SELECT id FROM objects ORDER BY id DESC LIMIT 1", DB_Conn, adOpenStatic, adLockOptimistic
    If DB_RS.EOF Then MsgBox "Oh crap, you don't have any objects! This isn't going to be pretty...", vbOKOnly
    NumObjDatas = Val(DB_RS(0))
    DB_RS.Close

    'Resize the objects array
    ObjData.SetDataUBound NumObjDatas
    
    'Retrieve the objects from the database
    DB_RS.Open "SELECT * FROM objects", DB_Conn, adOpenStatic, adLockOptimistic

    'Fill the object list
    Do While Not DB_RS.EOF   'Loop until we reach the end of the recordset
        With TempObjData
            Log "Load_OBJs: Filling ObjData for object ID " & DB_RS!id, CodeTracker '//\\LOGLINE//\\
            .Name = Trim$(DB_RS!Name)
            .Value = Val(DB_RS!Price)
            .ClassReq = Val(DB_RS!ClassReq)
            .ObjType = Val(DB_RS!ObjType)
            .WeaponType = Val(DB_RS!WeaponType)
            .WeaponRange = Val(DB_RS!WeaponRange)
            .ProjectileRotateSpeed = Val(DB_RS!ProjectileRotateSpeed)
            .UseGrh = Val(DB_RS!UseGrh)
            .UseSfx = Val(DB_RS!UseSfx)
            .GrhIndex = Val(DB_RS!GrhIndex)
            .SpriteBody = Val(DB_RS!sprite_body)
            .SpriteWeapon = Val(DB_RS!sprite_weapon)
            .SpriteHair = Val(DB_RS!sprite_hair)
            .SpriteHead = Val(DB_RS!sprite_head)
            .SpriteWings = Val(DB_RS!sprite_wings)
            .RepHP = Val(DB_RS!replenish_hp)
            .RepMP = Val(DB_RS!replenish_mp)
            .RepSP = Val(DB_RS!replenish_sp)
            .RepHPP = Val(DB_RS!replenish_hp_percent)
            .RepMPP = Val(DB_RS!replenish_mp_percent)
            .RepSPP = Val(DB_RS!replenish_sp_percent)
            .AddStat(SID.Speed) = Val(DB_RS!stat_speed)
            .AddStat(SID.Str) = Val(DB_RS!stat_str)
            .AddStat(SID.Agi) = Val(DB_RS!stat_agi)
            .AddStat(SID.Mag) = Val(DB_RS!stat_mag)
            .AddStat(SID.DEF) = Val(DB_RS!stat_def)
            .AddStat(SID.MinHIT) = Val(DB_RS!stat_hit_min)
            .AddStat(SID.MaxHIT) = Val(DB_RS!stat_hit_max)
            .AddStat(SID.MaxHP) = Val(DB_RS!stat_hp)
            .AddStat(SID.MaxMAN) = Val(DB_RS!stat_mp)
            .AddStat(SID.MaxSTA) = Val(DB_RS!stat_sp)
            .ReqAgi = Val(DB_RS!req_agi)
            .ReqStr = Val(DB_RS!req_str)
            .ReqLvl = Val(DB_RS!req_lvl)
            .ReqMag = Val(DB_RS!req_mag)
            .Stacking = Val(DB_RS!Stacking)
            If .Stacking < 1 Then .Stacking = MaxObjAmount
        End With
        
        'Save the file
        FileNum = FreeFile
        Open ServerTempPath & "o" & DB_RS!id & ".temp" For Binary Access Write As #FileNum
            Put #FileNum, , TempObjData
        Close #FileNum
        
        'Move to the next object
        DB_RS.MoveNext
        
    Loop

    'Close the recordset
    DB_RS.Close
    
End Sub

Public Sub Load_Quests()

    Log "Call Load_Quests", CodeTracker '//\\LOGLINE//\\
    
    'Get the number of quests (Sort by id, descending, only get 1 entry, only return id)
    DB_RS.Open "SELECT id FROM quests ORDER BY id DESC LIMIT 1", DB_Conn, adOpenStatic, adLockOptimistic
    If DB_RS.EOF Then MsgBox "Oh crap, you don't have any quests! This isn't going to be pretty...", vbOKOnly
    NumQuests = DB_RS(0)
    DB_RS.Close
    
    Log "Load_Quests: Resizing QuestData array by NumQuests (" & NumQuests & ")", CodeTracker '//\\LOGLINE//\\
    
    'Resize the quests array
    ReDim QuestData(1 To NumQuests)

    'Retrieve the data from the database
    DB_RS.Open "SELECT * FROM quests", DB_Conn, adOpenStatic, adLockOptimistic
    
    'Fill in the information
    Do While Not DB_RS.EOF  'Loop until we reach the end of the recordset
        With QuestData(DB_RS!id)
            Log "Load_Quests: Filling QuestData for quest " & DB_RS!id, CodeTracker '//\\LOGLINE//\\
            .Name = Trim$(DB_RS!Name)
            .Redoable = Val(DB_RS!Redoable)
            .StartTxt = Trim$(DB_RS!text_start)
            .AcceptTxt = Trim$(DB_RS!text_accept)
            .IncompleteTxt = Trim$(DB_RS!text_incomplete)
            .FinishTxt = Trim$(DB_RS!text_finish)
            .AcceptReqLvl = Val(DB_RS!accept_req_level)
            .AcceptReqObj = Val(DB_RS!accept_req_obj)
            .AcceptReqObjAmount = Val(DB_RS!accept_req_objamount)
            .AcceptReqFinishQuest = Val(DB_RS!accept_req_finishquest)
            .AcceptRewExp = Val(DB_RS!accept_reward_exp)
            .AcceptRewGold = Val(DB_RS!accept_reward_gold)
            .AcceptRewObj = Val(DB_RS!accept_reward_obj)
            .AcceptRewObjAmount = Val(DB_RS!accept_reward_objamount)
            .AcceptLearnSkill = Val(DB_RS!accept_reward_learnskill)
            .FinishReqObj = Val(DB_RS!finish_req_obj)
            .FinishReqObjAmount = Val(DB_RS!finish_req_objamount)
            .FinishReqNPC = Val(DB_RS!finish_req_killnpc)
            .FinishReqNPCAmount = Val(DB_RS!finish_req_killnpcamount)
            .FinishRewExp = Val(DB_RS!finish_reward_exp)
            .FinishRewGold = Val(DB_RS!finish_reward_gold)
            .FinishRewObj = Val(DB_RS!finish_reward_obj)
            .FinishRewObjAmount = Val(DB_RS!finish_reward_objamount)
            .FinishLearnSkill = Val(DB_RS!finish_reward_learnskill)
        End With
        DB_RS.MoveNext
    Loop
    
    'Close the recordset
    DB_RS.Close
    
End Sub

Public Sub Load_ServerIni()

'*****************************************************************
'Loads the Server.ini
'*****************************************************************
Dim TempSplit() As String
Dim ts() As String
Dim i As Byte
Dim s As String
Dim j As Long

    Log "Call Load_ServerIni", CodeTracker '//\\LOGLINE//\\

    'Misc
    IdleLimit = Val(Var_Get(ServerDataPath & "Server.ini", "INIT", "IdleLimit"))
    LastPacket = Val(Var_Get(ServerDataPath & "Server.ini", "INIT", "LastPacket"))

    'Start pos
    TempSplit() = Split(Var_Get(ServerDataPath & "Server.ini", "INIT", "StartPos"), "-")
    StartPos.Map = Val(TempSplit(0))
    StartPos.X = Val(TempSplit(1))
    StartPos.Y = Val(TempSplit(2))

    'Res pos
    TempSplit() = Split(Var_Get(ServerDataPath & "Server.ini", "INIT", "ResPos"), "-")
    ResPos.Map = Val(TempSplit(0))
    ResPos.X = Val(TempSplit(1))
    ResPos.Y = Val(TempSplit(2))
    
    'Get the total number of servers
    NumServers = Val(Var_Get(ServerDataPath & "Server.ini", "INIT", "Servers"))
    ReDim ServerInfo(1 To NumServers)
    
    'Make sure the server ID is not over the number of servers specified
    If ServerID > NumServers Then MsgBox "The ID for this server is " & ServerID & " but the max servers defined (NumServers) is " & NumServers & "!", vbOKOnly Or vbCritical
    
    'Get the list of servers
    ReDim ServerMap(1 To NumMaps)
    For i = 1 To NumServers
        With ServerInfo(i)
             
            'Get the general information
            .IIP = Trim$(Var_Get(ServerDataPath & "Server.ini", "SERVER" & i, "IP"))
            .EIP = Trim$(Var_Get(ServerDataPath & "Server.ini", "SERVER" & i, "ExIP"))
            .Port = Val(Var_Get(ServerDataPath & "Server.ini", "SERVER" & i, "Port"))
            .ServerPort = Val(Var_Get(ServerDataPath & "Server.ini", "SERVER" & i, "ServerPort"))
            If i = ServerID Then MaxUsers = Val(Var_Get(ServerDataPath & "Server.ini", "SERVER" & i, "MaxUsers"))
            
            'Get the map information
            s = Trim$(Var_Get(ServerDataPath & "Server.ini", "SERVER" & i, "Map"))
            TempSplit() = Split(s, ",")
            For j = 0 To UBound(TempSplit)
               If Trim$(TempSplit(j)) = "*" Then
                   FillMemory ServerMap(1), NumMaps + 1, i
               Else
                   If InStr(1, TempSplit(j), "-") Then
                       ts() = Split(Trim$(TempSplit(j)), "-")
                       FillMemory ServerMap(Val(ts(0))), Val(ts(1)) - Val(ts(0)) + 1, i
                   Else
                       If Val(TempSplit(j)) > 0 Then ServerMap(Val(TempSplit(j))) = i
                   End If
               End If
            Next j
            
        End With
        
    Next i
    
    'Make sure all maps have a server, since a map without a server = very bad
    ' - if a user goes to that map, they won't ever be able to get off, along with it can crash all the servers
    j = 10
    s = vbNullString
    For i = 1 To NumMaps
        If ServerMap(i) = 0 Or ServerMap(i) > NumServers Then
            j = j + 1
            If j > 10 Then  'Display only 10 maps per line
                s = s & vbNewLine & i & "(" & ServerMap(i) & ")"
            Else
                s = s & "," & i & "(" & ServerMap(i) & ")"
            End If
        End If
    Next i
    If LenB(s) Then MsgBox "The following map files contain invalid servers:" & vbNewLine & "(Server shown in parenthesis)" & vbNewLine & s, vbCritical Or vbOKOnly

End Sub

Public Sub Load_User(ByVal UserIndex As Integer, ByVal UserName As String)
Dim TempStr() As String
Dim TempStr2() As String
Dim InvStr As String
Dim MailStr As String
Dim BankStr As String
Dim KSStr As String
Dim CurQStr As String
Dim CompQStr As String
Dim tMinHP As Long
Dim tMinSP As Long
Dim tMinMP As Long
Dim i As Long

    Log "Call Load_User(N/A," & UserName & ")", CodeTracker '//\\LOGLINE//\\

    'Retrieve the user from the database
    DB_RS.Open "SELECT * FROM users WHERE `name`='" & UserName & "'", DB_Conn, adOpenStatic, adLockOptimistic

    'Make sure the character exists
    If DB_RS.EOF Then
        DB_RS.Close
        Exit Sub
    End If
    
    'Loop through every field - match up the names then set the data accordingly
    With DB_RS
        UserList(UserIndex).Desc = Trim$(!Descr)
        UserList(UserIndex).Flags.GMLevel = !gm
        InvStr = !inventory
        MailStr = !mail
        BankStr = !Bank
        KSStr = !KnownSkills
        CompQStr = Trim$(!CompletedQuests)
        CurQStr = Trim$(!currentquest)
        UserList(UserIndex).BankGold = Val(!BankGold)
        UserList(UserIndex).Class = Val(!Class)
        UserList(UserIndex).Pos.X = Val(!pos_x)
        UserList(UserIndex).Pos.Y = Val(!pos_y)
        UserList(UserIndex).Pos.Map = Val(!pos_map)
        UserList(UserIndex).Char.Hair = Val(!char_hair)
        UserList(UserIndex).Char.Head = Val(!char_head)
        UserList(UserIndex).Char.Body = Val(!char_body)
        UserList(UserIndex).Char.Weapon = Val(!char_weapon)
        UserList(UserIndex).Char.Wings = Val(!char_wings)
        UserList(UserIndex).Char.Heading = Val(!char_heading)
        UserList(UserIndex).Char.HeadHeading = Val(!char_headheading)
        UserList(UserIndex).WeaponEqpSlot = Val(!eq_weapon)
        UserList(UserIndex).ArmorEqpSlot = Val(!eq_armor)
        UserList(UserIndex).WingsEqpSlot = Val(!eq_wings)
        UserList(UserIndex).Stats.BaseStat(SID.Speed) = Val(!stat_speed)
        UserList(UserIndex).Stats.BaseStat(SID.Str) = Val(!stat_str)
        UserList(UserIndex).Stats.BaseStat(SID.Agi) = Val(!stat_agi)
        UserList(UserIndex).Stats.BaseStat(SID.Mag) = Val(!stat_mag)
        UserList(UserIndex).Stats.BaseStat(SID.DEF) = Val(!stat_def)
        UserList(UserIndex).Stats.BaseStat(SID.Gold) = Val(!stat_gold)
        UserList(UserIndex).Stats.BaseStat(SID.EXP) = Val(!stat_exp)
        UserList(UserIndex).Stats.BaseStat(SID.ELV) = Val(!stat_elv)
        UserList(UserIndex).Stats.BaseStat(SID.ELU) = Val(!stat_elu)
        UserList(UserIndex).Stats.BaseStat(SID.Points) = Val(!stat_points)
        UserList(UserIndex).Stats.BaseStat(SID.MinHIT) = Val(!stat_hit_min)
        UserList(UserIndex).Stats.BaseStat(SID.MaxHIT) = Val(!stat_hit_max)
        UserList(UserIndex).Stats.BaseStat(SID.MaxHP) = Val(!stat_hp_max) 'Max HP/SP/MP MUST be loaded before the mins!
        UserList(UserIndex).Stats.BaseStat(SID.MaxMAN) = Val(!stat_mp_max)
        UserList(UserIndex).Stats.BaseStat(SID.MaxSTA) = Val(!stat_sp_max)
        UserList(UserIndex).Stats.ModStat(SID.MaxHP) = UserList(UserIndex).Stats.BaseStat(SID.MaxHP)
        UserList(UserIndex).Stats.ModStat(SID.MaxMAN) = UserList(UserIndex).Stats.BaseStat(SID.MaxMAN)
        UserList(UserIndex).Stats.ModStat(SID.MaxSTA) = UserList(UserIndex).Stats.BaseStat(SID.MaxSTA)
        
        'We have to wait until we know the modified max stats, we can set the minimum, or else the user will never be able to log
        ' in with their stats full if they have +HP, +SP or +MP stat modifiers equipped
        tMinHP = Val(!stat_hp_min)
        tMinMP = Val(!stat_mp_min)
        tMinSP = Val(!stat_sp_min)
        
        'Update the server the user is on
        !server = ServerID
        .Update
    
        'Close the recordset
        .Close
        
    End With

    'Inventory string
    If LenB(InvStr) Then
        Log "Load_User: Splitting inventory string (" & InvStr & ")", CodeTracker '//\\LOGLINE//\\
        TempStr = Split(InvStr, vbNewLine)  'Split up the inventory slots
        For i = 0 To UBound(TempStr)        'Loop through the slots
            Log "Load_User: Splitting item data (" & TempStr(i) & ")", CodeTracker '//\\LOGLINE//\\
            TempStr2 = Split(TempStr(i), " ")   'Split up the slot, objindex, amount and equipted (in that order)
            If Val(TempStr2(0)) <= MAX_INVENTORY_SLOTS Then
                With UserList(UserIndex).Object(Val(TempStr2(0)))
                    .ObjIndex = Val(TempStr2(1))
                    .Amount = Val(TempStr2(2))
                    .Equipped = Val(TempStr2(3))
                End With
            Else '//\\LOGLINE//\\
                Log "Load_User: User has too many inventory slots - tried applying slot " & Val(TempStr2(0)), CriticalError '//\\LOGLINE//\\
            End If
        Next i
    End If
    
    'Bank string
    If LenB(BankStr) Then
        Log "Load_User: Splitting bank string (" & InvStr & ")", CodeTracker '//\\LOGLINE//\\
        TempStr = Split(BankStr, vbNewLine) 'Split the bank slots
        For i = 0 To UBound(TempStr)        'Loop through the slots
            TempStr2 = Split(TempStr(i), " ")   'Split up the slot, objindex and amount (in that order)
            If Val(TempStr2(0)) <= MAX_INVENTORY_SLOTS Then
                With UserList(UserIndex).Bank(Val(TempStr2(0)))
                    .ObjIndex = Val(TempStr2(1))
                    .Amount = Val(TempStr2(2))
                End With
            Else '//\\LOGLINE//\\
                Log "Load_User: User has too many bank slots - tried applying slot " & Val(TempStr2(0)), CriticalError '//\\LOGLINE//\\
            End If
        Next i
    End If
                    
    'Mail string
    If LenB(MailStr) Then
        Log "Load_User: Splititng mail string (" & MailStr & ")", CodeTracker '//\\LOGLINE//\\
        TempStr = Split(MailStr, vbNewLine) 'Split up the mail indexes
        For i = 0 To UBound(TempStr)
            If i <= MaxMailPerUser Then
                UserList(UserIndex).MailID(i + 1) = Val(TempStr(i))
            Else '//\\LOGLINE//\\
                Log "Load_User: User has too many mails - tried applying slot " & i, CriticalError '//\\LOGLINE//\\
            End If
        Next i
    End If
    
    'Known skills string (if the index is stored, then that skill is known - if not stored, then unknown)
    If LenB(KSStr) Then
        TempStr = Split(KSStr, vbNewLine)   'Split up the known skill indexes
        For i = 0 To UBound(TempStr)
            If Val(TempStr(i)) <= NumSkills Then
                UserList(UserIndex).KnownSkills(Val(TempStr(i))) = 1
            Else '//\\LOGLINE//\\
                Log "Load_User: User has too many skills - tried applying slot " & i, CriticalError '//\\LOGLINE//\\
            End If
        Next i
    End If
    
    'Completed quests string
    If LenB(CompQStr) Then
        TempStr = Split(CompQStr, ",")
        UserList(UserIndex).NumCompletedQuests = UBound(TempStr) + 1
        ReDim UserList(UserIndex).CompletedQuests(1 To UserList(UserIndex).NumCompletedQuests)
        For i = 0 To UserList(UserIndex).NumCompletedQuests - 1
            UserList(UserIndex).CompletedQuests(i + 1) = Int(TempStr(i))
        Next i
    End If
    
    'Current quest string
    If LenB(CurQStr) Then
        TempStr = Split(CurQStr, vbNewLine)    'Split up the quests
        For i = 0 To UBound(TempStr)
            If i + 1 < MaxQuests Then 'Make sure we are within limit
                TempStr2 = Split(TempStr(i), " ")   'Split up the QuestID and NPCKills (in that order)
                UserList(UserIndex).Quest(i + 1) = Val(TempStr2(0))
                UserList(UserIndex).QuestStatus(i + 1).NPCKills = Val(TempStr2(1))
            Else '//\\LOGLINE//\\
                Log "Load_User: User has too many quests - tried applying quest " & i + 1, CriticalError '//\\LOGLINE//\\
            End If
        Next i
    End If
    
    'Equipt items
    If UserList(UserIndex).WeaponEqpSlot > 0 Then UserList(UserIndex).WeaponEqpObjIndex = UserList(UserIndex).Object(UserList(UserIndex).WeaponEqpSlot).ObjIndex
    If UserList(UserIndex).ArmorEqpSlot > 0 Then UserList(UserIndex).ArmorEqpObjIndex = UserList(UserIndex).Object(UserList(UserIndex).ArmorEqpSlot).ObjIndex
    If UserList(UserIndex).WingsEqpSlot > 0 Then UserList(UserIndex).WingsEqpObjIndex = UserList(UserIndex).Object(UserList(UserIndex).WingsEqpSlot).ObjIndex

    'Update the user's mod stats first before setting the min (current) values
    User_UpdateModStats UserIndex

    'Force stat updates to the client
    UserList(UserIndex).Stats.ForceFullUpdate

    'We can finally set the min stats
    UserList(UserIndex).Stats.BaseStat(SID.MinHP) = tMinHP
    UserList(UserIndex).Stats.BaseStat(SID.MinSTA) = tMinSP
    UserList(UserIndex).Stats.BaseStat(SID.MinMAN) = tMinMP
    
    'Misc values
    UserList(UserIndex).Name = UserName
    
End Sub

Public Sub Save_Mail(ByVal MailIndex As Long, ByRef MailData As MailData)
Dim s As String
Dim i As Long

    Log "Call Save_Mail(" & MailIndex & "," & "N/A)", CodeTracker '//\\LOGLINE//\\

    'Build the object string
    For i = 1 To MaxMailObjs
        If MailData.Obj(i).ObjIndex > 0 Then
            If MailData.Obj(i).Amount > 0 Then
                If LenB(s) Then s = s & vbNewLine   'Split the line, but make sure we dont add a split on first entry
                s = s & MailData.Obj(i).ObjIndex & " " & MailData.Obj(i).Amount
            End If
        End If
    Next i
    Log "Save_Mail: Built object string (" & s & ")", CodeTracker '//\\LOGLINE//\\
    
    With DB_RS
        
        'If we are updating the mail, then the record must be deleted, so make sure it isn't there (or else we get a duplicate key entry error)
        DB_Conn.Execute "DELETE FROM mail WHERE id=" & MailIndex
    
        'Open the database with an empty table
        .Open "SELECT * FROM mail WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
        .AddNew
        
        'Put the data in the recordset
        !id = Str$(MailIndex)
        !Sub = MailData.Subject
        !By = MailData.WriterName
        !Date = MailData.RecieveDate
        !Msg = MailData.Message
        !New = Str$(MailData.New)
        !Objs = s
        
        'Update the database with the new piece of mail
        .Update
       
        'Close the database
        .Close
        
    End With

End Sub

Public Sub Save_User(UserChar As User, ByVal UserIndex As Integer, Optional ByVal Password As String, Optional ByVal NewUser As Byte)

'*****************************************************************
'Saves a user's data to a .chr file
'*****************************************************************
Dim BankStr As String
Dim InvStr As String
Dim MailStr As String
Dim KSStr As String
Dim CurQStr As String
Dim CompQStr As String
Dim i As Long

    Log "Call Save_User(N/A," & UserIndex & "," & Password & "," & NewUser & ")", CodeTracker '//\\LOGLINE//\\

    'On special occasions, we want to the typical saving routine and save before the user is even disconnected
    If UserChar.Flags.DoNotSave = 1 Then Exit Sub

    With UserChar
    
        'Make sure we are trying to save a valid user by testing a few variables first
        If Len(.Name) < 3 Then
            Log "Save_User: Specified name was invalid (" & .Name & ")", CriticalError '//\\LOGLINE//\\
            Exit Sub
        End If
        If Len(.Name) > 10 Then
            Log "Save_User: Specified name was invalid (" & .Name & ")", CriticalError '//\\LOGLINE//\\
            Exit Sub
        End If

        'Build the inventory string
        For i = 1 To MAX_INVENTORY_SLOTS
            If .Object(i).ObjIndex > 0 Then
                If LenB(InvStr) Then InvStr = InvStr & vbNewLine   'Add the line break, but dont add it to first entry
                InvStr = InvStr & i & " " & .Object(i).ObjIndex & " " & .Object(i).Amount & " " & .Object(i).Equipped
            End If
        Next i
        Log "Save_User: Built inventory string (" & InvStr & ")", CodeTracker '//\\LOGLINE//\\
        
        'Build mail string
        For i = 1 To MaxMailPerUser
            If .MailID(i) > 0 Then
                If LenB(MailStr) Then
                    MailStr = MailStr & vbNewLine & .MailID(i)
                Else
                    MailStr = MailStr & .MailID(i)
                End If
            End If
        Next i
        Log "Save_User: Built mail string (" & MailStr & ")", CodeTracker '//\\LOGLINE//\\
        
        'Build known skills string
        For i = 1 To NumSkills
            If .KnownSkills(i) > 0 Then
                If LenB(KSStr) Then
                    KSStr = KSStr & vbNewLine & i
                Else
                    KSStr = KSStr & i
                End If
            End If
        Next i
        Log "Save_User: Built known skills string (" & KSStr & ")", CodeTracker '//\\LOGLINE//\\
        
        'Build completed quest string
        For i = 1 To .NumCompletedQuests
            If i < .NumCompletedQuests Then
                CompQStr = CompQStr & .CompletedQuests(i) & ","
            Else
                CompQStr = CompQStr & .CompletedQuests(i)
            End If
        Next i
        Log "Save_User: Built completed quests string (" & CompQStr & ")", CodeTracker '//\\LOGLINE//\\
        
        'Build current quest string
        For i = 1 To MaxQuests
            If .Quest(i) > 0 Then
                If LenB(CurQStr) Then
                    CurQStr = CurQStr & vbNewLine & .Quest(i) & " " & .QuestStatus(i).NPCKills
                Else
                    CurQStr = CurQStr & .Quest(i) & " " & .QuestStatus(i).NPCKills
                End If
            End If
        Next i
        Log "Save_User: Built current quest string (" & CurQStr & ")", CodeTracker '//\\LOGLINE//\\
        
        'Build the bank string
        For i = 1 To MAX_INVENTORY_SLOTS
            If .Bank(i).ObjIndex > 0 Then
                If LenB(BankStr) Then
                    BankStr = BankStr & vbNewLine & i & " " & .Bank(i).ObjIndex & " " & .Bank(i).Amount
                Else
                    BankStr = BankStr & i & " " & .Bank(i).ObjIndex & " " & .Bank(i).Amount
                End If
            End If
        Next i
        Log "Save_User: Built bank string (" & BankStr & ")", CodeTracker '//\\LOGLINE//\\
        
        'Check whether we have to make a new entry or can update an old one
        If NewUser Then
        
            'Open the database with an empty record and create the new user
            DB_RS.Open "SELECT * FROM users WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
            DB_RS.AddNew
            
        Else
        
            'Open the old record and update it
            DB_RS.Open "SELECT * FROM users WHERE `name`='" & .Name & "'", DB_Conn, adOpenStatic, adLockOptimistic
            
        End If
        
        'Put the data in the recordset
        If LenB(Password) Then DB_RS!Password = Password    'If no password is passed, we don't need to update it
        If NewUser Then
            DB_RS!date_create = Date$
            DB_RS!Name = .Name
            DB_RS!IP = frmMain.GOREsock.Address(UserIndex)
        End If
        DB_RS!date_lastlogin = Date$
        DB_RS!gm = .Flags.GMLevel
        DB_RS!Class = .Class
        DB_RS!Descr = .Desc
        DB_RS!BankGold = .BankGold
        DB_RS!inventory = InvStr
        DB_RS!mail = MailStr
        DB_RS!KnownSkills = KSStr
        DB_RS!Bank = BankStr
        DB_RS!CompletedQuests = CompQStr
        DB_RS!currentquest = CurQStr
        DB_RS!pos_x = .Pos.X
        DB_RS!pos_y = .Pos.Y
        DB_RS!pos_map = .Pos.Map
        DB_RS!char_hair = .Char.Hair
        DB_RS!char_head = .Char.Head
        DB_RS!char_body = .Char.Body
        DB_RS!char_weapon = .Char.Weapon
        DB_RS!char_wings = .Char.Wings
        DB_RS!char_heading = .Char.Heading
        DB_RS!char_headheading = .Char.HeadHeading
        DB_RS!eq_weapon = .WeaponEqpSlot
        DB_RS!eq_armor = .ArmorEqpSlot
        DB_RS!eq_wings = .WingsEqpSlot
        DB_RS!stat_speed = .Stats.BaseStat(SID.Speed)
        DB_RS!stat_str = .Stats.BaseStat(SID.Str)
        DB_RS!stat_agi = .Stats.BaseStat(SID.Agi)
        DB_RS!stat_mag = .Stats.BaseStat(SID.Mag)
        DB_RS!stat_def = .Stats.BaseStat(SID.DEF)
        DB_RS!stat_gold = .Stats.BaseStat(SID.Gold)
        DB_RS!stat_exp = .Stats.BaseStat(SID.EXP)
        DB_RS!stat_elv = .Stats.BaseStat(SID.ELV)
        DB_RS!stat_elu = .Stats.BaseStat(SID.ELU)
        DB_RS!stat_points = .Stats.BaseStat(SID.Points)
        DB_RS!stat_hit_min = .Stats.BaseStat(SID.MinHIT)
        DB_RS!stat_hit_max = .Stats.BaseStat(SID.MaxHIT)
        DB_RS!stat_hp_min = .Stats.BaseStat(SID.MinHP)
        DB_RS!stat_hp_max = .Stats.BaseStat(SID.MaxHP)
        DB_RS!stat_mp_min = .Stats.BaseStat(SID.MinMAN)
        DB_RS!stat_mp_max = .Stats.BaseStat(SID.MaxMAN)
        DB_RS!stat_sp_min = .Stats.BaseStat(SID.MinSTA)
        DB_RS!stat_sp_max = .Stats.BaseStat(SID.MaxSTA)
        DB_RS!server = 0
            
    End With
    
    'Update the database
    DB_RS.Update
    
    'Close the recordset
    DB_RS.Close

End Sub

Public Sub Save_PacketsIn()

'*****************************************************************
'Save the outbound packet records
'*****************************************************************
Dim LoopC As Long
Dim FileNum As Byte
Dim s As String
    
    'Build the string
    For LoopC = 0 To 254
        s = s & LoopC & ": " & DebugPacketsIn(LoopC) & vbNewLine
    Next LoopC
    s = s & "255: " & DebugPacketsIn(LoopC)    'Easy way to take off the last vbNewLine
    
    'Make sure the directory exists
    MakeSureDirectoryPathExists LogPath
    
    'Delete the old file if it exists
    If Server_FileExist(LogPath & ServerID & "\packetsin.txt", vbNo) Then Kill LogPath & ServerID & "\packetsin.txt"
    
    'Write to the file
    FileNum = FreeFile
    Open LogPath & ServerID & "\packetsin.txt" For Binary Access Write As #FileNum
        Put #FileNum, , s
    Close #FileNum

End Sub

Public Sub Save_PacketsOut()

'*****************************************************************
'Save the outbound packet records
'*****************************************************************
Dim LoopC As Long
Dim FileNum As Byte
Dim s As String
    
    'Build the string
    For LoopC = 0 To 254
        s = s & LoopC & ": " & DebugPacketsOut(LoopC) & vbNewLine
    Next LoopC
    s = s & "255: " & DebugPacketsOut(LoopC)    'Easy way to take off the last vbNewLine
    
    'Make sure the directory exists
    MakeSureDirectoryPathExists LogPath
    
    'Delete the old file if it exists
    If Server_FileExist(LogPath & ServerID & "\packetsout.txt", vbNo) Then Kill LogPath & ServerID & "\packetsout.txt"
    
    'Write to the file
    FileNum = FreeFile
    Open LogPath & ServerID & "\packetsout.txt" For Binary Access Write As #FileNum
        Put #FileNum, , s
    Close #FileNum

End Sub

Public Sub Save_FPS()

'*****************************************************************
'Save the FPS records
'*****************************************************************
Dim FileNum As Byte

    'Make sure the directory exists
    MakeSureDirectoryPathExists LogPath
    
    'Delete the old file if it exists
    If Server_FileExist(LogPath & ServerID & "\serverfps.txt", vbNo) Then Kill LogPath & ServerID & "\serverfps.txt"
    
    'Write to the file
    FileNum = FreeFile
    Open LogPath & ServerID & "\serverfps.txt" For Binary Access Write As #FileNum
        Put #FileNum, , FPSIndex
        Put #FileNum, , ServerFPS()
    Close #FileNum
    
End Sub

Public Function Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String, Optional ByVal DontLog As Byte = 0) As String

'*****************************************************************
'Gets a variable from a text file
'*****************************************************************

    If DontLog = 0 Then Log "Call Var_Get(" & File & "," & Main & "," & Var & ")", CodeTracker '//\\LOGLINE//\\

    Var_Get = Space$(1000)
    GetPrivateProfileString Main, Var, vbNullString, Var_Get, 1000, File
    Var_Get = RTrim$(Var_Get)
    If LenB(Var_Get) <> 0 Then Var_Get = Left$(Var_Get, Len(Var_Get) - 1)
    
    If DontLog = 0 Then Log "Rtrn Var_Get = " & Var_Get, CodeTracker '//\\LOGLINE//\\

End Function

Public Sub Var_Write(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    Log "Call Var_Write(" & File & "," & Main & "," & Var & "," & Value & ")", CodeTracker '//\\LOGLINE//\\

    WritePrivateProfileString Main, Var, Value, File

End Sub
