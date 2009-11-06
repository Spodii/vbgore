VERSION 5.00
Begin VB.UserControl GOREsockServer 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   PaletteMode     =   4  'None
   Picture         =   "GOREsock.ctx":0000
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ToolboxBitmap   =   "GOREsock.ctx":0C42
End
Attribute VB_Name = "GOREsockServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This is the maximum size of data we will handle at once
'If there is more data then this, we won't get it, which can be very, very bad
'This is only used for receiving, not sending!
Private Const BufferSize As Long = 3072

'This is the maximum number of connections allowed per IP
'Change to 0 to remove the limit
'This is intended for two things:
' - Prevent DOS attack by socket connection flooding from a single computer
' - Prevent multi-logging (not a recommended reason)
'Note that multiple people may play that are on the same network, so try to keep this
' value at a decent height for those people
Private Const MaxConnectionsPerIP As Long = 5

Private WindowhWnd As Long

Private Type typPortal 'Class specific variables
    WndProc As Long 'Pointer to the origional WindowProc of our window (We need to give control of ALL messages back to it before we destroy it)
    Sockets As Long 'How many Sockets are comming through the Portal, Actually hold the Socket array count. NB - MUST change with Redim of Sockets
End Type

Public Enum enmSoxState
    soxDisconnected = 0&
    soxListening = 1&
    soxConnecting = 2&
    soxIdle = 3&
    soxSend = 4&
    soxRecv = 5&
    soxClosing = 6& ' Necessary so we don't try to Send on this Socket (Because apiShutDown says so!)
    soxBound = 10& ' The socket has been bound to its current Port and Address
    soxERROR = -1&
End Enum
#If False Then
Private soxDisconnected, soxListening, soxConnecting, soxIdle, soxSend, soxRecv, soxClosing, soxBound, soxERROR
#End If

Public Enum enmSoxOptions
    soxSO_BROADCAST = &H20& 'BOOL Allow transmission of broadcast messages on the
    soxSO_DEBUG = &H1& 'BOOL Record debugging information.
    soxSO_DONTROUTE = &H10& 'BOOL Do not route: send directly to interface.
    soxSO_KEEPALIVE = &H8& 'BOOL Send keepalives
    soxSO_LINGER = &H80& 'struct LINGER  Linger on close if unsent data is present.
    soxSO_OOBINLINE = &H100& 'BOOL Receive out-of-band data in the normal data stream. (See section DECnet Out-Of-band data for a discussion of this topic.)
    soxSO_RCVBUF = &H1002& 'int Specify the total per-socket buffer space reserved for receives. This is unrelated to SO_MAX_MSG_SIZE or the size of a TCP window.
    soxSO_REUSEADDR = &H4& 'BOOL Allow the socket to be bound to an address that is already in use. (See bind.)
    soxSO_SNDBUF = &H1001& 'int Specify the total per-socket buffer space reserved for sends. This is unrelated to SO_MAX_MSG_SIZE or the size of a TCP window.
    soxSO_TCP_NODELAY = Not &H1& 'BOOL Disables the Nagle algorithm for send coalescing.
    soxSO_USELOOPBACK = &H40& 'bypass hardware when possible
    soxSO_ACCEPTCONN = &H2& 'BOOL Socket is listening.
    soxSO_ERROR = &H1007& 'int Retrieve error status and clear.
    soxSO_TYPE = &H1008& 'Get Socket Type (From FTP - Experimental) (Seems to always returns 1 for a valid TCP socket, -1 for a closed socket)
End Enum
#If False Then
Private soxSO_BROADCAST, soxSO_DEBUG, soxSO_DONTROUTE, soxSO_KEEPALIVE, soxSO_LINGER, soxSO_OOBINLINE, soxSO_RCVBUF, soxSO_REUSEADDR, soxSO_SNDBUF, _
        soxSO_TCP_NODELAY, soxSO_USELOOPBACK, soxSO_ACCEPTCONN, soxSO_ERROR, soxSO_TYPE
#End If

Public Enum enmSoxTypes ' Basically, soxSERVER means the Sox number was 'accepted' by a listening connection, and soxCLIENT means we used connect to connect to a Server (on the other side, our connection will be soxSERVER)
    soxSERVER = 4026& ' This indicates that the Socket is either a Listening Socket, or was created from a Listening Socket, either way, our machine is acting as a Sox Server
    soxCLIENT = 4027& ' This indicates that the Socket is a connection we established to another computer/server, therefore our machine is acting as a Sox Client on this Socket
End Enum
#If False Then
Private soxSERVER, soxCLIENT
#End If

'API Defined
Private Const SOCKADDR_SIZE As Long = 16
Private Type typSocketAddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(7) As Byte
End Type

'Class module Defined
Private Type typSocket
    PacketInPos As Byte     'Used the find the key to encrypt/decrypt with
    PacketOutPos As Byte
    Socket As Long ' The actual WinSock API socket number
    SocketAddr As typSocketAddr ' Info about the connection
    State As enmSoxState
    uMsg As Long ' Server (-1) / Client (0) Socket (Server = A Socket that has a connection to the Server / Client = A Socket that was created in Accept that connected to us)
End Type

Private Type typBuffer ' The advantage of using this is if we sent exactly 8K on the other side, when we receive 8K, FD_READ will not be sent again so we won't get an error like when we use a loop
    Size As Long ' Array Size (To check if there is incomming data, we can check the size of this variable, if -1 then we are not receiving anything)
    Buffer() As Byte
End Type

Private Type typWSAData
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 127) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Const INVALID_SOCKET As Long = -1& ' Indication of an Invalid Socket
Private Const ERROR As Long = -1&
Private Const INADDR_NONE As Long = &HFFFFFFFF 'Was FFFF (Confirmed) ... Returned address is an error

Private Const AF_INET As Integer = 2
Private Const SOCK_STREAM As Long = 1    'stream socket
Private Const SOL_SOCKET As Long = &HFFFF& 'Officially the only option for socket level
Private Const FD_READ As Long = &H1
Private Const FD_WRITE As Long = &H2
Private Const FD_ACCEPT As Long = &H8
Private Const FD_CONNECT As Long = &H10
Private Const FD_CLOSE As Long = &H20
Private Const SD_SEND As Long = &H1
Private Const IPPROTO_TCP As Long = 6 'tcp
Private Const GWL_WndProc As Long = (-4)

'Combinations of flags (process it as a const instead of real-time - slightly (oh so slightly) faster)
Private Const FD_CLOSEREADWRITE As Long = FD_CLOSE Or FD_READ Or FD_WRITE
Private Const FD_CONNECTLISTEN As Long = FD_ACCEPT Or FD_CLOSE Or FD_CONNECT Or FD_READ Or FD_WRITE

Private Declare Function apiWSAStartup Lib "WS2_32" Alias "WSAStartup" (ByVal wVersionRequired As Long, lpWSADATA As typWSAData) As Long
Private Declare Function apiWSACleanup Lib "WS2_32" Alias "WSACleanup" () As Long
Private Declare Function apiSocket Lib "WS2_32" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Private Declare Function apiCloseSocket Lib "WS2_32" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function apiBind Lib "WS2_32" Alias "bind" (ByVal s As Long, addr As typSocketAddr, ByVal namelen As Long) As Long
Private Declare Function apiListen Lib "WS2_32" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function apiConnect Lib "WS2_32" Alias "connect" (ByVal s As Long, Name As typSocketAddr, ByVal namelen As Long) As Long
Private Declare Function apiAccept Lib "WS2_32" Alias "accept" (ByVal s As Long, addr As typSocketAddr, addrlen As Long) As Long
Private Declare Function apiWSAAsyncSelect Lib "WS2_32" Alias "WSAAsyncSelect" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function apiRecv Lib "WS2_32" Alias "recv" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function apiSend Lib "WS2_32" Alias "send" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function apiGetSockOpt Lib "WS2_32" Alias "getsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function apiSetSockOpt Lib "WS2_32" Alias "setsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function apiHToNS Lib "WS2_32" Alias "htons" (ByVal hostshort As Long) As Integer 'Host To Network Short
Private Declare Function apiNToHS Lib "WS2_32" Alias "ntohs" (ByVal netshort As Long) As Integer 'Network To Host Short
Private Declare Function apiIPToNL Lib "WS2_32" Alias "inet_addr" (ByVal cp As String) As Long
Private Declare Function apiNLToIP Lib "WS2_32" Alias "inet_ntoa" (ByVal inn As Long) As Long
Private Declare Function apiShutDown Lib "WS2_32" Alias "shutdown" (ByVal s As Long, ByVal how As Long) As Long
Private Declare Function apiCallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function apiSetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function apiLStrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function apiLstrCpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Sub apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'Events
Public Event OnConnection(inSox As Long)
Public Event OnClose(inSox As Long)
Public Event OnDataArrival(inSox As Long, inData() As Byte)
Public Event OnConnecting(inSox As Long)

'The buffer used to grab data (to prevent having to constantly resize the buffer)
Private GetBuffer(0 To BufferSize - 1) As Byte

Private WSAData As typWSAData   'Stores WinSock data on initialization of WinSock 2
Private Portal As typPortal     'Sorta used for general variables
Private Sockets() As typSocket  'Information on each of the sockets
Private Send() As typBuffer     'Send Buffer

Private PacketEncTypeServerIn As Byte
Private PacketEncTypeServerOut As Byte
Private PacketKeys() As String
Private Const PacketEncTypeNone As Byte = 0  'Use no encryption
Private Const PacketEncTypeRC4 As Byte = 1   'Use RC4 encryption
Private Const PacketEncTypeXOR As Byte = 2   'Use XOR encryption

Public Sub SetEncryption(ByVal gPacketEncTypeServerIn As Byte, ByVal gPacketEncTypeServerOut As Byte, ByRef gPacketKeys() As String)
    
    'Set the in/out encryption types
    PacketEncTypeServerIn = gPacketEncTypeServerIn
    PacketEncTypeServerOut = gPacketEncTypeServerOut
    
    'Set the packet keys (or erase them if no encryption is used)
    If PacketEncTypeServerIn <> PacketEncTypeNone Or PacketEncTypeServerOut <> PacketEncTypeNone Then
        PacketKeys = gPacketKeys
    Else
        Erase PacketKeys
    End If
    
End Sub

Private Function Accept(ByVal inSocket As Long) As Long 'Returns: New Sox Number -- inSocket is the listening WinSocket, not Sox number ...
Dim tmpSocket As Long
Dim tmpSocketAddr As typSocketAddr 'This stores the details of our new socket/client, including the client IP address

    Let tmpSocket = apiAccept(inSocket, tmpSocketAddr, SOCKADDR_SIZE) 'Accept API returns a valid, random, unused socket for us to use for the new client
    If tmpSocket = INVALID_SOCKET Then 'Accept API may not give us a valid socket eg. when all sockets are full, you may have to add additional error trapping if you believe you will use over 32,767 sockets
        'Since a socket was not commited for the new Connection ... we don't have to close it (Since the socket was never even created)
        Let Accept = INVALID_SOCKET
    Else ' Success, A new connection ... Accept now contains the new Socket number
        For Accept = 0 To Portal.Sockets ' First search to see if the socket already exists
            If Sockets(Accept).Socket = tmpSocket Then Exit For
        Next Accept
        If Accept = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
            For Accept = 0 To Portal.Sockets ' First search to see if the socket already exists
                If Sockets(Accept).Socket = soxDisconnected Then Exit For ' Found an open Socket
            Next Accept
            If Accept = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                ReDim Preserve Sockets(Accept) As typSocket
                ReDim Preserve Send(Accept) As typBuffer
                Let Portal.Sockets = Accept
            End If
        End If
        Let Sockets(Accept).Socket = tmpSocket
        Let Sockets(Accept).SocketAddr = tmpSocketAddr 'Set the details of the new socket/client
        Let Sockets(Accept).uMsg = soxSERVER  'This is a Client Socket - It has connected to US
        Let Sockets(Accept).PacketInPos = 0
        Let Sockets(Accept).PacketOutPos = 0
        Let Send(Accept).Size = -1
        Erase Send(Accept).Buffer
        Call RaiseState(Accept, soxConnecting) ' Could possibly leave this on soxDisconnected, and on Select Case State, thurn it on and set it ready to send data (Or set it to connecting)
        RaiseEvent OnConnection(Accept)
    End If

End Function

Public Function Address(ByVal inSox As Long) As String ' Returns the address used by a Socket (Either Local or Remote)

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let Address = soxERROR
    Else
        Let Address = StringFromPointer(apiNLToIP(Sockets(inSox).SocketAddr.sin_addr))
    End If

End Function

Public Function Bind(LocalIP As String, LocalPort As Integer) As Long
Dim tmpSocket As Long
Dim tmpSocketAddr As typSocketAddr

    If LocalPort = 0 Or LocalIP = vbNullString Then
        Let Bind = soxERROR
    Else
        Let tmpSocketAddr.sin_family = AF_INET
        Let tmpSocketAddr.sin_port = apiHToNS(LocalPort)
        If tmpSocketAddr.sin_port = INVALID_SOCKET Then
            Let Bind = INVALID_SOCKET
        Else
            Let tmpSocketAddr.sin_addr = apiIPToNL(LocalIP) 'If this is Zero, it will assign 0.0.0.0 !!!
            If tmpSocketAddr.sin_addr = INADDR_NONE Then 'If 255.255.255.255 is returned ... we have a problem ... I think :)
                Let Bind = INVALID_SOCKET
            Else
                Let tmpSocket = apiSocket(AF_INET, SOCK_STREAM, IPPROTO_TCP) 'This is where you specify what type of protocol to use and what type of Streaming to use, returns a new socket number 4 us (NB - From here, if any further steps fail after this one succeeds, we must close the socket)
                If tmpSocket = INVALID_SOCKET Then
                    Let Bind = INVALID_SOCKET
                Else
                    If apiBind(tmpSocket, tmpSocketAddr, SOCKADDR_SIZE) = ERROR Then 'Socket Number, Socket Address space / Name, Name Length ...
                        apiCloseSocket tmpSocket
                        Let Bind = ERROR
                    Else
                        For Bind = 0 To Portal.Sockets ' First search to see if the socket already exists
                            If Sockets(Bind).Socket = tmpSocket Then Exit For
                        Next Bind
                        If Bind = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                            For Bind = 0 To Portal.Sockets ' First search to see if the socket already exists
                                If Sockets(Bind).Socket = soxDisconnected Then Exit For ' Found an open Socket
                            Next Bind
                            If Bind = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                                ReDim Preserve Sockets(Bind) As typSocket
                                ReDim Preserve Send(Bind) As typBuffer
                                Let Portal.Sockets = Bind
                            End If
                        End If
                        Let Sockets(Bind).Socket = tmpSocket
                        Let Sockets(Bind).SocketAddr = tmpSocketAddr 'Set the details of the new socket/client
                        Call RaiseState(Bind, soxBound)
                    End If
                End If
            End If
        End If
    End If

End Function

Private Sub Closed(inSox As Long) ' This Socket has successfully closed ... free resources (No need to check if it exists, cause we call this internally)

    If Not (inSox < 0 Or inSox > Portal.Sockets) Then ' Detect out of Range of our Array ...
        If apiWSAAsyncSelect(Sockets(inSox).Socket, WindowhWnd, 0&, 0&) = ERROR Then       'FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
            Call RaiseState(inSox, soxDisconnected) ' Force disconnected status, dunno what the implications are!
        Else
            If apiCloseSocket(Sockets(inSox).Socket) = ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                Call RaiseState(inSox, soxDisconnected) ' Force disconnected status, dunno what the implications are!
            Else
                Call RaiseState(inSox, soxDisconnected)
                RaiseEvent OnClose(inSox)
            End If
        End If
    End If

End Sub

Public Function Connect(RemoteHost As String, RemotePort As Integer) As Long 'Returns the new Sox Number / ERROR On Error
Dim tmpSocket As Long
Dim tmpSocketAddr As typSocketAddr

    Let tmpSocketAddr.sin_family = AF_INET
    Let tmpSocketAddr.sin_port = apiHToNS(RemotePort) ' apiHToNS(RemotePort)
    If tmpSocketAddr.sin_port = INVALID_SOCKET Then
        Let Connect = INVALID_SOCKET
    Else
        Let tmpSocketAddr.sin_addr = apiIPToNL(RemoteHost) 'If this is Zero, it will assign 0.0.0.0 !!!
        If tmpSocketAddr.sin_addr = INADDR_NONE Then 'If 255.255.255.255 is returned ... we have a problem ... I think :)
            Let Connect = INVALID_SOCKET
        Else
            Let tmpSocket = apiSocket(AF_INET, SOCK_STREAM, IPPROTO_TCP) 'This is where you specify what type of protocol to use and what type of Streaming to use, returns a new socket number 4 us (NB - From here, if any further steps fail after this one succeeds, we must close the socket)
            If tmpSocket = INVALID_SOCKET Then
                Let Connect = INVALID_SOCKET
            Else
                If apiConnect(tmpSocket, tmpSocketAddr, SOCKADDR_SIZE) = ERROR Then
                    apiCloseSocket tmpSocket
                    Let Connect = ERROR
                Else
                    If apiWSAAsyncSelect(tmpSocket, WindowhWnd, ByVal soxCLIENT, ByVal FD_CONNECTLISTEN) = ERROR Then ' Reassign this Socket to Send and Receive on the DATA channel
                        apiCloseSocket tmpSocket
                        Let Connect = ERROR
                    Else
                        For Connect = 0 To Portal.Sockets ' First search to see if the socket already exists
                            If Sockets(Connect).Socket = tmpSocket Then Exit For
                        Next Connect
                        If Connect = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                            For Connect = 0 To Portal.Sockets ' First search to see if the socket already exists
                                If Sockets(Connect).Socket = soxDisconnected Then Exit For ' Found an open Socket
                            Next Connect
                            If Connect = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                                ReDim Preserve Sockets(Connect) As typSocket
                                ReDim Preserve Send(Connect) As typBuffer
                                Let Portal.Sockets = Connect
                            End If
                        End If
                        Let Sockets(Connect).Socket = tmpSocket
                        Let Sockets(Connect).SocketAddr = tmpSocketAddr 'Set the details of the new socket/client
                        Let Sockets(Connect).uMsg = soxCLIENT ' This is a Server connection - We have connected to it (Could even be another Client computer but the fact is we connected to it)
                        Let Sockets(Connect).PacketInPos = 0
                        Let Sockets(Connect).PacketOutPos = 0
                        Let Send(Connect).Size = -1
                        Erase Send(Connect).Buffer
                        Call RaiseState(Connect, soxConnecting)
                    End If
                End If
            End If
        End If
    End If

End Function

Public Function Connections() As Long ' Returns the number of clients connected to Sox

Dim tmpLoop As Long

    For tmpLoop = 0 To Portal.Sockets
        If Not Sockets(tmpLoop).State = soxDisconnected Then
            If Sockets(tmpLoop).uMsg = soxSERVER Then
                Select Case Sockets(tmpLoop).State ' These are the valid states for 'connected' sockets
                Case soxConnecting: Let Connections = Connections + 1
                Case soxIdle: Let Connections = Connections + 1
                Case soxClosing: Let Connections = Connections + 1
                Case soxRecv: Let Connections = Connections + 1
                Case soxSend: Let Connections = Connections + 1
                End Select
            End If
        End If
    Next tmpLoop

End Function

Private Sub GetData(inSox As Long) ' Extracts data from the WinSock Recv buffers and places it in our local buffer (data() array)
Dim tmpRecvSize As Integer
Dim tmpBuffer() As Byte
Dim Size As Integer
Dim Pos As Long

    'Check for valid parameters
    If Not (inSox < 0 Or inSox > Portal.Sockets) Then
        If Sockets(inSox).State = soxIdle Then
            Call RaiseState(inSox, soxRecv)
            
            'Select the socket
            If apiWSAAsyncSelect(Sockets(inSox).Socket, WindowhWnd, 0&, 0&) <> ERROR Then
                
                'Get the data from the socket
                Let tmpRecvSize = apiRecv(Sockets(inSox).Socket, GetBuffer(0), BufferSize, 0)

                Select Case tmpRecvSize
                
                    'There was an error getting the data, do nothing
                    Case ERROR
                    
                    'Socket was gracefully closed
                    Case 0
                        Call RaiseState(inSox, soxDisconnected)
                        RaiseEvent OnClose(inSox)
                        
                    'Get our data
                    Case Else
                    
                        If PacketEncTypeServerIn <> PacketEncTypeNone Then
                        
                            '*** Encrypted Receive ***
                            Do
                                
                                'Get the size of the packet
                                apiCopyMemory Size, GetBuffer(Pos), 2
                                Pos = Pos + 2

                                'Resize the buffer to fit the size of the packet
                                ReDim tmpBuffer(0 To Size - 1)
                                
                                'Copy the data into the buffer
                                apiCopyMemory tmpBuffer(0), GetBuffer(Pos), Size
                                
                                'Update the read position
                                Pos = Pos + Size + 1
                                
                                'Decrypt the packet
                                With Sockets(inSox)
                                    Select Case PacketEncTypeServerIn
                                        Case PacketEncTypeXOR
                                            Encryption_XOR_DecryptByte tmpBuffer(), PacketKeys(.PacketInPos)
                                        Case PacketEncTypeRC4
                                            Encryption_RC4_DecryptByte tmpBuffer(), PacketKeys(.PacketInPos)
                                    End Select
                                    .PacketInPos = .PacketInPos + 1
                                    If .PacketInPos > PacketEncKeys - 1 Then .PacketInPos = 0
                                End With
                                
                                'Call the data handling routine
                                RaiseEvent OnDataArrival(inSox, tmpBuffer())
                                DoEvents
                                
                            Loop While Pos < tmpRecvSize - 3
                        
                        Else

                            '*** Non-encrypted Receive ***
                            'Copy the actual data over to a new buffer, then send it to the program
                            ReDim tmpBuffer(0 To tmpRecvSize - 1)
                            apiCopyMemory tmpBuffer(0), GetBuffer(0), tmpRecvSize
                            RaiseEvent OnDataArrival(inSox, tmpBuffer())
                            
                        End If
                            
                End Select
                
                'Change the socket state (unless it is closing)
                If Sockets(inSox).State = soxRecv Then Call RaiseState(inSox, soxIdle)
                apiWSAAsyncSelect Sockets(inSox).Socket, WindowhWnd, ByVal Sockets(inSox).uMsg, ByVal FD_CLOSEREADWRITE
                
            End If
        End If
    End If

End Sub

Public Function GetOption(inSox As Long, inOption As enmSoxOptions) As Long

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let GetOption = soxERROR
    Else
        Select Case inOption
        Case soxSO_TCP_NODELAY
            If apiGetSockOpt(Sockets(inSox).Socket, IPPROTO_TCP, Not inOption, GetOption, 4) = ERROR Then Let GetOption = ERROR
        Case Else
            If apiGetSockOpt(Sockets(inSox).Socket, SOL_SOCKET, inOption, GetOption, 4) = ERROR Then Let GetOption = ERROR
        End Select
    End If

End Function

Public Sub Hook() ' WinSock is told to send it's messages to the Sox Control, but we need to intercept these messages!

    If Portal.WndProc = 0 Then ' If it's already hooked to our WindowProc function, we could have problems, this will make sure we've UnHooked before
        Let Portal.WndProc = apiSetWindowLong(WindowhWnd, GWL_WndProc, AddressOf WindowProc)
    End If

End Sub

'Creates a socket and sets it in listen mode. This method works only for TCP connections

Public Function Listen(inAddress As String, inPort As Integer) As Long   'Returns Sox number / ERROR On Error
Dim tmpSocket As Long
Dim tmpSocketAddr As typSocketAddr

    Let tmpSocketAddr.sin_family = AF_INET
    Let tmpSocketAddr.sin_port = apiHToNS(inPort)
    If tmpSocketAddr.sin_port = INVALID_SOCKET Then
        Let Listen = INVALID_SOCKET
    Else
        Let tmpSocketAddr.sin_addr = apiIPToNL(inAddress) 'If this is Zero, it will assign 0.0.0.0 !!!
        If tmpSocketAddr.sin_addr = INADDR_NONE Then 'If 255.255.255.255 is returned ... we have a problem ... I think :)
            Let Listen = INVALID_SOCKET
        Else
            Let tmpSocket = apiSocket(AF_INET, SOCK_STREAM, IPPROTO_TCP) 'This is where you specify what type of protocol to use and what type of Streaming to use, returns a new socket number 4 us (NB - From here, if any further steps fail after this one succeeds, we must close the socket)
            If tmpSocket = INVALID_SOCKET Then
                Let Listen = INVALID_SOCKET
            Else
                If apiBind(tmpSocket, tmpSocketAddr, SOCKADDR_SIZE) = ERROR Then 'Socket Number, Socket Address space / Name, Name Length ...
                    apiCloseSocket tmpSocket
                    Let Listen = ERROR
                Else
                    If apiListen(ByVal tmpSocket, ByVal 5) = ERROR Then ' 5 = Maximum connections
                        apiCloseSocket tmpSocket
                        Let Listen = ERROR
                    Else
                        If apiWSAAsyncSelect(tmpSocket, WindowhWnd, ByVal soxSERVER, ByVal FD_CONNECTLISTEN) = ERROR Then ' Reassign this Socket to Send and Receive on the DATA channel
                            apiCloseSocket tmpSocket
                            Let Listen = ERROR
                        Else
                            For Listen = 0 To Portal.Sockets ' First search to see if the socket already exists
                                If Sockets(Listen).Socket = tmpSocket Then Exit For
                            Next Listen
                            If Listen = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                                For Listen = 0 To Portal.Sockets ' First search to see if the socket already exists
                                    If Sockets(Listen).Socket = soxDisconnected Then Exit For ' Found an open Socket
                                Next Listen
                                If Listen = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                                    ReDim Preserve Sockets(Listen) As typSocket
                                    ReDim Preserve Send(Listen) As typBuffer
                                    Let Portal.Sockets = Listen
                                End If
                            End If
                            Let Sockets(Listen).Socket = tmpSocket
                            Let Sockets(Listen).SocketAddr = tmpSocketAddr 'Set the details of the new socket/client
                            Let Sockets(Listen).uMsg = soxSERVER
                            Let Sockets(Listen).PacketInPos = 0
                            Let Sockets(Listen).PacketOutPos = 0
                            Let Send(Listen).Size = -1
                            Erase Send(Listen).Buffer
                            Call RaiseState(Listen, soxListening)
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Public Function Port(inSox As Long) As Long

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let Port = soxERROR
    Else
        Let Port = apiNToHS(Sockets(inSox).SocketAddr.sin_port)
    End If

End Function

Private Sub RaiseState(inSox As Long, inState As enmSoxState)

    Let Sockets(inSox).State = inState
    If inState = soxConnecting Then RaiseEvent OnConnecting(inSox)

End Sub

Private Sub SendBuffer(inSox As Long)
Dim i As Integer
Dim b() As Byte

    'Check for invalid parameter
    If Not (inSox < 0 Or inSox > Portal.Sockets) Then
        If Not Send(inSox).Size = soxERROR Then 'If there is data in the buffer
            If Sockets(inSox).State = soxIdle Then
            
                Call RaiseState(inSox, soxSend)
                apiWSAAsyncSelect Sockets(inSox).Socket, WindowhWnd, 0&, 0&
                
                'If we have data, send 'er over, cap'n!
                If Send(inSox).Size + 1 > 0 Then
                    
                    With Sockets(inSox)
                        
                        '*** Encrypted send ***
                        If PacketEncTypeServerOut <> PacketEncTypeNone Then
                            
                            'Encrypt the packet
                            Select Case PacketEncTypeServerOut
                                Case PacketEncTypeXOR
                                    Encryption_XOR_EncryptByte Send(inSox).Buffer(), PacketKeys(.PacketOutPos)
                                Case PacketEncTypeRC4
                                    Encryption_RC4_EncryptByte Send(inSox).Buffer(), PacketKeys(.PacketOutPos)
                            End Select
                            
                            'Add the length (this is so we can handle packets that get combined in the network)
                            i = UBound(Send(inSox).Buffer) + 1
                            ReDim b(0 To i + 2)
                            apiCopyMemory b(2), Send(inSox).Buffer(0), i
                            apiCopyMemory b(0), i, 2
                            ReDim Send(inSox).Buffer(0 To i + 2)
                            apiCopyMemory Send(inSox).Buffer(0), b(0), UBound(Send(inSox).Buffer()) + 1
                        
                        End If
                        
                        'Send the data
                        If apiSend(Sockets(inSox).Socket, Send(inSox).Buffer(0), UBound(Send(inSox).Buffer) + 1, 0) <> ERROR Then
                            
                            'All data send, clear the buffer
                            Let Send(inSox).Size = -1
                            Erase Send(inSox).Buffer
                            
                            '*** Encrypted send successfully ***
                            If PacketEncTypeServerOut <> PacketEncTypeNone Then
                                'Raise the encryption count
                                .PacketOutPos = .PacketOutPos + 1
                                If .PacketOutPos > PacketEncKeys - 1 Then .PacketOutPos = 0
                            End If
     
                        Else
                            
                            '*** Encrypted send fail ***
                            'Send failed, unencrypt the packet so we can encrypt it again later
                            If PacketEncTypeServerOut <> PacketEncTypeNone Then
                                Select Case PacketEncTypeServerOut
                                    Case PacketEncTypeXOR
                                        Encryption_XOR_EncryptByte Send(inSox).Buffer(), PacketKeys(.PacketOutPos)
                                    Case PacketEncTypeRC4
                                        Encryption_RC4_EncryptByte Send(inSox).Buffer(), PacketKeys(.PacketOutPos)
                                End Select
                            End If
                            
                        End If
                        
                    End With
                    
                End If
                
                'Change the state of the socket (unless it is closing)
                If Sockets(inSox).State = soxSend Then Call RaiseState(inSox, soxIdle)
                apiWSAAsyncSelect Sockets(inSox).Socket, WindowhWnd, ByVal Sockets(inSox).uMsg, ByVal FD_CLOSEREADWRITE
                
            End If
        End If
    End If

End Sub

Public Sub SendData(ByVal inSox As Long, inData() As Byte)
Dim inDataUBound As Long
    
    'If inData() is empty, it will create an error - in this case, abort
    On Error GoTo ErrOut
    inDataUBound = UBound(inData)
    On Error GoTo 0

    If inDataUBound > -1 Then
        If inSox > -1 Then
            If inSox <= Portal.Sockets Then ' Detect out of Range of our Array ...
                If Sockets(inSox).State = soxIdle Or Sockets(inSox).State = soxSend Or Sockets(inSox).State = soxRecv Then ' If we have initiated a ShutDown, the state would change to Closing
                    ReDim Preserve Send(inSox).Buffer(Send(inSox).Size + inDataUBound + 1) As Byte    'UBound + 1 = DataLength
                    Call apiCopyMemory(Send(inSox).Buffer(Send(inSox).Size + 1), inData(0), inDataUBound + 1) 'Copy the data
                    Let Send(inSox).Size = Send(inSox).Size + inDataUBound + 1    'Increase according to the data
                    apiWSAAsyncSelect Sockets(inSox).Socket, WindowhWnd, ByVal Sockets(inSox).uMsg, ByVal FD_CLOSEREADWRITE
                End If
            End If
        End If
    End If
    
ErrOut:
    
End Sub

Public Function SetOption(ByVal inSox As Long, inOption As enmSoxOptions, inValue As Long) As Long

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let SetOption = soxERROR
    Else
        If inOption = soxSO_TCP_NODELAY Then
            If apiSetSockOpt(Sockets(inSox).Socket, IPPROTO_TCP, Not inOption, inValue, 4) = ERROR Then Let SetOption = ERROR
        Else
            If apiSetSockOpt(Sockets(inSox).Socket, SOL_SOCKET, inOption, inValue, 4) = ERROR Then Let SetOption = ERROR
        End If
    End If

End Function

Public Function Shut(ByVal inSox As Long) As Long ' Initiates ShutDown procedure for a Socket

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let Shut = INVALID_SOCKET
    Else
        Select Case Sockets(inSox).State
            Case soxDisconnected
            Case soxClosing
            Case soxBound
                If apiWSAAsyncSelect(Sockets(inSox).Socket, WindowhWnd, 0&, 0&) = ERROR Then
                    Let Shut = ERROR
                Else
                    If apiCloseSocket(Sockets(inSox).Socket) = ERROR Then
                        Let Shut = ERROR
                    Else
                        Call RaiseState(inSox, soxDisconnected)
                        RaiseEvent OnClose(inSox)
                    End If
                End If
            Case soxListening
                If apiWSAAsyncSelect(Sockets(inSox).Socket, WindowhWnd, 0&, 0&) = ERROR Then
                    Let Shut = ERROR
                Else
                    If apiCloseSocket(Sockets(inSox).Socket) = ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                        Let Shut = ERROR
                    Else
                        Call RaiseState(inSox, soxDisconnected)
                        RaiseEvent OnClose(inSox)
                    End If
                End If
            Case Else
                If apiGetSockOpt(Sockets(inSox).Socket, SOL_SOCKET, soxSO_ERROR, Shut, 4) = ERROR Then
                    Let Shut = ERROR
                Else
                    If apiShutDown(Sockets(inSox).Socket, SD_SEND) = ERROR Then
                        Let Shut = ERROR
                    Else
                        Call RaiseState(inSox, soxClosing)
                    End If
                End If
        End Select
    End If

End Function

' Returns the ShutDown 'Go Ahead' ...

Public Function ShutDown() As Long ' Closes Listening and Bound Sockets immediately, sends apiShutDown to the rest
Dim tmpSox As Long

    For tmpSox = 0 To Portal.Sockets
        Select Case Sockets(tmpSox).State
        Case soxDisconnected
        Case soxClosing ' No need to close a closing Socket
        Case soxBound ' Same as soxListening
            If apiWSAAsyncSelect(Sockets(tmpSox).Socket, WindowhWnd, 0&, 0&) = ERROR Then
                Call RaiseState(tmpSox, soxDisconnected)
                RaiseEvent OnClose(tmpSox)
                Let ShutDown = ERROR
            Else
                If apiCloseSocket(Sockets(tmpSox).Socket) = ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                    Call RaiseState(tmpSox, soxDisconnected)
                    RaiseEvent OnClose(tmpSox)
                    Let ShutDown = ERROR
                Else
                    Call RaiseState(tmpSox, soxDisconnected)
                    RaiseEvent OnClose(tmpSox)
                End If
            End If
        Case soxListening ' Same as soxBound
            If apiWSAAsyncSelect(Sockets(tmpSox).Socket, WindowhWnd, 0&, 0&) = ERROR Then
                Call RaiseState(tmpSox, soxDisconnected)
                RaiseEvent OnClose(tmpSox)
                Let ShutDown = ERROR
            Else
                If apiCloseSocket(Sockets(tmpSox).Socket) = ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                    Call RaiseState(tmpSox, soxDisconnected)
                    RaiseEvent OnClose(tmpSox)
                    Let ShutDown = ERROR
                Else
                    Call RaiseState(tmpSox, soxDisconnected)
                    RaiseEvent OnClose(tmpSox)
                End If
            End If
        Case Else
            If apiShutDown(Sockets(tmpSox).Socket, SD_SEND) = ERROR Then
                Let ShutDown = soxERROR
            Else
                Call RaiseState(tmpSox, soxClosing)
            End If
        End Select
    Next tmpSox
    DoEvents ' There could be an incomming FD_CLOSE
    For tmpSox = 0 To Portal.Sockets
        If Not Sockets(tmpSox).State = soxDisconnected Then Exit For
    Next tmpSox
    If Not tmpSox = Portal.Sockets + 1 Then
        Let ShutDown = soxERROR ' This could also be set by all the soxClosing sockets
    Else
        DoEvents
        Let Portal.Sockets = -1
        Erase Sockets
        Erase Send
    End If

End Function

Private Function Socket2Sox(inSocket As Long) As Long ' Returns the Sockets() Array address of a WinSock Socket

    For Socket2Sox = 0 To Portal.Sockets
        If Sockets(Socket2Sox).Socket = inSocket Then Exit For
    Next Socket2Sox
    If Socket2Sox = Portal.Sockets + 1 Then Let Socket2Sox = INVALID_SOCKET

End Function

Public Function SocketHandle(ByVal inSox As Long) As Long

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let SocketHandle = soxERROR
    Else
        Let SocketHandle = Sockets(inSox).Socket
    End If

End Function

Public Function State(ByVal inSox As Long) As enmSoxState

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let State = soxERROR
    Else
        Let State = Sockets(inSox).State
    End If

End Function

Private Function StringFromPointer(ByVal lPointer As Long) As String

    Let StringFromPointer = Space$(apiLStrLen(ByVal lPointer))
    Call apiLstrCpy(ByVal StringFromPointer, ByVal lPointer)

End Function

Public Function uMsg(inSox As Long) As Long  ' This is the closest I can probably get to defining the type of Sox

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let uMsg = soxERROR
    Else
        Let uMsg = Sockets(inSox).uMsg
    End If

End Function

Public Sub UnHook() ' Once the Control is UnHooked, we will not be able to intercept messages from WinSock API and process them according to our needs!
    
    Call apiSetWindowLong(WindowhWnd, GWL_WndProc, Portal.WndProc)
    Let Portal.WndProc = 0

End Sub

Private Function WinSockError(ByVal lParam As Long) As Integer 'WSAGETSELECTERROR

    Let WinSockError = (lParam And &HFFFF0000) \ &H10000

End Function

Private Function WinSockEvent(ByVal lParam As Long) As Integer 'WSAGETSELECTEVENT

    If (lParam And &HFFFF&) > &H7FFF Then
        Let WinSockEvent = (lParam And &HFFFF&) - &H10000
    Else
        Let WinSockEvent = lParam And &HFFFF&
    End If

End Function

Private Sub ValidateAccept(ByVal inSox As Long)
Dim Count As Long
Dim i As Long

    For i = 0 To Portal.Sockets
        If i <> inSox Then
            If Sockets(i).State <> soxDisconnected Then
                If Sockets(i).State <> soxERROR Then
                    If Sockets(i).State <> soxClosing Then
                        If Sockets(inSox).SocketAddr.sin_addr = Sockets(i).SocketAddr.sin_addr Then
                            Count = Count + 1
                            If Count > MaxConnectionsPerIP Then
                                Shut inSox
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
            
End Sub

Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case uMsg
    Case soxSERVER
        Select Case WinSockEvent(lParam)
        Case FD_ACCEPT
            If WinSockError(lParam) = 0 Then
                ValidateAccept Accept(wParam)
            End If
        Case FD_CLOSE
            Select Case WinSockError(lParam)
            Case 0
                Select Case Sockets(Socket2Sox(wParam)).State
                Case soxClosing: Call Closed(Socket2Sox(wParam))
                Case Else
                    Call Shut(Socket2Sox(wParam))
                    Call Closed(Socket2Sox(wParam))
                End Select
            Case Else
                Call Closed(Socket2Sox(wParam))
            End Select
        Case FD_READ
            Call GetData(Socket2Sox(wParam))
        Case FD_WRITE
            If WinSockError(lParam) = 0 Then
                Select Case Sockets(Socket2Sox(wParam)).State
                Case soxConnecting
                    Call RaiseState(Socket2Sox(wParam), soxIdle)
                Case soxIdle
                    Call SendBuffer(Socket2Sox(wParam))
                Case soxClosing
                    Call Closed(Socket2Sox(wParam))
                End Select
            End If
        End Select
    Case soxCLIENT
        Select Case WinSockEvent(lParam)
        Case FD_CLOSE
            Select Case WinSockError(lParam)
            Case 0
                Select Case Sockets(Socket2Sox(wParam)).State
                Case soxClosing: Call Closed(Socket2Sox(wParam))
                Case Else
                    Call Shut(Socket2Sox(wParam))
                    Call Closed(Socket2Sox(wParam))
                End Select
            Case Else
                Call Closed(Socket2Sox(wParam))
            End Select
        Case FD_READ
            Call GetData(Socket2Sox(wParam))
        Case FD_WRITE
            If WinSockError(lParam) = 0 Then
                Select Case Sockets(Socket2Sox(wParam)).State
                Case soxConnecting
                    Call RaiseState(Socket2Sox(wParam), soxIdle)
                Case soxIdle
                    Call SendBuffer(Socket2Sox(wParam))
                Case soxClosing
                    Call Closed(Socket2Sox(wParam))
                End Select
            End If
        End Select
    Case Else
        Let WndProc = apiCallWindowProc(Portal.WndProc, hwnd, uMsg, wParam, lParam)
    End Select

End Function

Public Sub ClearPicture()

    Set UserControl.Picture = Nothing

End Sub

Private Sub UserControl_Initialize()

    WindowhWnd = hwnd
    If Not InIDE Then
        Set GOREsockServer = Me
        If apiWSAStartup(&H101, WSAData) = -1 Then
            Call MsgBox("WinSock failed to initialize properly - Error#: " & Err.LastDllError, vbApplicationModal + vbCritical, "Critical Error")  'Creates an 'application instance' and memory space in the WinSock DLL (MUST be cleaned up later)
        Else
            Let Portal.WndProc = apiSetWindowLong(UserControl.hwnd, GWL_WndProc, AddressOf WindowProc)
            Let Portal.Sockets = -1
        End If
    Else
        Let Portal.Sockets = -1
    End If

End Sub

Private Function InIDE() As Boolean

    On Local Error GoTo ErrHandler
    
    Debug.Print 1 / 0
    
    Exit Function
    
ErrHandler:

    Let InIDE = True
    
End Function

Private Sub UserControl_Resize()

    UserControl.Width = 480
    UserControl.Height = 480

End Sub

Private Sub UserControl_Terminate()

    'Correctly replaces/reattaches the origional WindowProc procedure to our 'hidden' handle (Basically what the UnHook command does!)
    If Not InIDE Then
        Call apiSetWindowLong(WindowhWnd, GWL_WndProc, Portal.WndProc)
        apiWSACleanup
    End If

End Sub
