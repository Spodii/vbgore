VERSION 5.00
Begin VB.UserControl SoxBinary 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   Picture         =   "Sox.ctx":0000
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   28
   ToolboxBitmap   =   "Sox.ctx":0972
End
Attribute VB_Name = "SoxBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const BufferSize = 8192

Public Enum enmSoxState
    ' Used
    soxDisconnected = 0&
    soxListening = 1&
    soxConnecting = 2&
    soxIdle = 3& ' Change to soxIdle
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
    soxSO_BROADCAST = &H20& 'BOOL Allow transmission of broadcast messages on the socket.
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

Private Type typPortal 'Class specific variables
    WndProc As Long 'Pointer to the origional WindowProc of our window (We need to give control of ALL messages back to it before we destroy it)
    Sockets As Long 'How many Sockets are comming through the Portal, Actually hold the Socket array count. NB - MUST change with Redim of Sockets
End Type

'API Defined
Private Const SOCKADDR_SIZE = 16&
Private Type typSocketAddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(7) As Byte
End Type

'Class module Defined
Private Type typSocket
    Socket As Long ' The actual WinSock API socket number
    SocketAddr As typSocketAddr ' Info about the connection
    State As enmSoxState
    uMsg As Long ' Server (-1) / Client (0) Socket (Server = A Socket that has a connection to the Server / Client = A Socket that was created in Accept that connected to us)
End Type

Private Type typBuffer ' The advantage of using this is if we sent exactly 8K on the other side, when we receive 8K, FD_READ will not be sent again so we won't get an error like when we use a loop
    Size As Long ' Array Size (To check if there is incomming data, we can check the size of this variable, if -1 then we are not receiving anything)
    Pos As Long
    Buffer() As Byte
End Type

Private Const WSADESCRIPTION_LEN = 255&
Private Const WSASYS_STATUS_LEN = 127&
Private Type typWSAData
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADESCRIPTION_LEN) As Byte
    szSystemStatus(0 To WSASYS_STATUS_LEN) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

'VB WinSock OCX Defined Error codes
Public Enum enmError
    sckOutOfMemory = 7                'Out of memory
    sckInvalidPropertyValue = 380     'The property value is invalid
    sckGetNotSupported = 394          'The property cannot be read
    sckSetNotSupported = 383          'The property is read-only
    sckBadState = 40006               'Wrong protocol or connection state for the requested transaction or request
    sckInvalidArg = 40014             'The argument passed to a function was not in the correct format or in the specified range
    sckSuccess = 40017                'Successful
    sckUnsupported = 40018            'Unsupported variant type
    sckInvalidOp = 40020              'Invalid operation at current state
    sckOutOfRange = 40021             'Argument is out of range
    sckWrongProtocol = 40026          'Wrong protocol for the requested transaction or request
    sckOpCanceled = 1004              'The operation was canceled
    sckInvalidArgument = 10014        'The requested address is a broadcast address, but flag is not set
    sckWouldBlock = 10035             'Socket is non-blocking and the specified operation will block
    sckInProgress = 10036             'A blocking Winsock operation in progress
    sckAlreadyComplete = 10037        'The operation is completed. No blocking operation in progress
    sckNotSocket = 10038              'The descriptor is not a socket
    sckMsgTooBig = 10040              'The datagram is too large to fit into the buffer and is truncated
    sckPortNotSupported = 10043       'The specified port is not supported
    sckAddressInUse = 10048           'Address in use
    sckAddressNotAvailable = 10049    'Address not available from the local machine
    sckNetworkSubsystemFailed = 10050 'Network subsystem failed
    sckNetworkUnreachable = 10051     'The network cannot be reached from this host at this time
    sckNetReset = 10052               'Connection has timed out when SO_KEEPALIVE is set
    sckConnectAborted = 10053         'Connection is aborted due to timeout or other failure
    sckConnectionReset = 10054        'The connection is reset by remote side
    sckNoBufferSpace = 10055          'No buffer space is available
    sckAlreadyConnected = 10056       'Socket is already connected
    sckNotConnected = 10057           'Socket is not connected
    sckSocketShutdown = 10058         'Socket has been shut down
    sckTimedout = 10060               'Socket has been shut down
    sckConnectionRefused = 10061      'Connection is forcefully rejected
    sckNotInitialized = 10093         'WinsockInit should be called first
    sckHostNotFound = 11001           'Authoritative answer: Host not found
    sckHostNotFoundTryAgain = 11002   'Non-Authoritative answer: Host not found
    sckNonRecoverableError = 11003    'Non-recoverable errors
    sckNoData = 11004                 'Valid name, no data record of requested type
End Enum
#If False Then
Private sckOutOfMemory, sckInvalidPropertyValue, sckGetNotSupported, sckSetNotSupported, sckBadState, sckInvalidArg, sckSuccess, sckUnsupported, _
        sckInvalidOp, sckOutOfRange, sckWrongProtocol, sckOpCanceled, sckInvalidArgument, sckWouldBlock, sckInProgress, sckAlreadyComplete, _
        sckNotSocket, sckMsgTooBig, sckPortNotSupported, sckAddressInUse, sckAddressNotAvailable, sckNetworkSubsystemFailed, _
        sckNetworkUnreachable, sckNetReset, sckConnectAborted, sckConnectionReset, sckNoBufferSpace, sckAlreadyConnected, _
        sckNotConnected, sckSocketShutdown, sckTimedout, sckConnectionRefused, sckNotInitialized, sckHostNotFound, sckHostNotFoundTryAgain, _
        sckNonRecoverableError, sckNoData
#End If

Private Const INVALID_SOCKET = -1& ' Indication of an Invalid Socket
Private Const SOCKET_ERROR = -1&
Private Const INADDR_NONE = &HFFFFFFFF 'Was FFFF (Confirmed) ... Returned address is an error

Private Const AF_INET = 2
Private Const SOCK_STREAM = 1    'stream socket
Private Const SOL_SOCKET = &HFFFF& 'Officially the only option for socket level
Private Const FD_READ = &H1
Private Const FD_WRITE = &H2
Private Const FD_ACCEPT = &H8
Private Const FD_CONNECT = &H10
Private Const FD_CLOSE = &H20
Private Const SD_SEND = &H1
Private Const IPPROTO_TCP = 6 'tcp
Private Const GWL_WNDPROC = (-4)

Private Declare Function apiWSAStartup Lib "WS2_32" Alias "WSAStartup" (ByVal wVersionRequired As Long, lpWSADATA As typWSAData) As Long
Private Declare Function apiWSACleanup Lib "WS2_32" Alias "WSACleanup" () As Long
Private Declare Function apiSocket Lib "WS2_32" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Private Declare Function apiCloseSocket Lib "WS2_32" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function apiBind Lib "WS2_32" Alias "bind" (ByVal s As Long, addr As typSocketAddr, ByVal namelen As Long) As Long
Private Declare Function apiListen Lib "WS2_32" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function apiConnect Lib "WS2_32" Alias "connect" (ByVal s As Long, name As typSocketAddr, ByVal namelen As Long) As Long
Private Declare Function apiAccept Lib "WS2_32" Alias "accept" (ByVal s As Long, addr As typSocketAddr, addrlen As Long) As Long
Private Declare Function apiWSAAsyncSelect Lib "WS2_32" Alias "WSAAsyncSelect" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function apiRecv Lib "WS2_32" Alias "recv" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function apiSend Lib "WS2_32" Alias "send" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function apiGetSockOpt Lib "WS2_32" Alias "getsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function apiSetSockOpt Lib "WS2_32" Alias "setsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function apiHToNS Lib "WS2_32" Alias "htons" (ByVal hostshort As Long) As Integer 'Host To Network Short
Private Declare Function apiNToHS Lib "WS2_32" Alias "ntohs" (ByVal netshort As Long) As Integer 'Network To Host Short
Private Declare Function apiIPToNL Lib "WS2_32" Alias "inet_addr" (ByVal cp As String) As Long
Private Declare Function apiNLToIP Lib "WS2_32" Alias "inet_ntoa" (ByVal inn As Long) As Long
Private Declare Function apiGetHostName Lib "WS2_32" Alias "gethostname" (ByVal name As String, ByVal namelen As Long) As Long
Private Declare Function apiShutDown Lib "WS2_32" Alias "shutdown" (ByVal s As Long, ByVal how As Long) As Long
Private Declare Function apiCallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function apiSetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function apiLStrLen Lib "Kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function apiLstrCpy Lib "Kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Sub apiCopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Event OnClose(inSox As Long)
Public Event OnConnection(inSox As Long) 'Notification of a new connection (From a Listening Port)
Public Event OnDataArrival(inSox As Long, inData() As Byte) ' Unlike the origional WinSock OCX, a byte Array is used as the Data type instead of a Variant as this is a faster way of getting data to you directly
Public Event OnConnecting(inSox As Long) 'Connecting

Private WSAData As typWSAData 'Stores WinSock data on initialization of WinSock 2
Private Portal As typPortal ' Sorta used for general variables (The word Portal came from my use in previous versions of a STATIC window, do you know what that is? :)
Private Sockets() As typSocket
Private Recv() As typBuffer ' Receive Buffer
Private Send() As typBuffer ' Send Buffer

Private Function Accept(inSocket As Long) As Long 'Returns: New Sox Number -- inSocket is the listening WinSocket, not Sox number ...
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
                ReDim Preserve Recv(Accept) As typBuffer
                ReDim Preserve Send(Accept) As typBuffer
                Let Portal.Sockets = Accept
            End If
        End If
        Let Sockets(Accept).Socket = tmpSocket
        Let Sockets(Accept).SocketAddr = tmpSocketAddr 'Set the details of the new socket/client
        Let Sockets(Accept).uMsg = soxSERVER  'This is a Client Socket - It has connected to US
        Let Recv(Accept).Size = -1
        Let Recv(Accept).Pos = -1
        Erase Recv(Accept).Buffer
        Let Send(Accept).Size = -1
        Let Send(Accept).Pos = -1
        Erase Send(Accept).Buffer
        Call RaiseState(Accept, soxConnecting) ' Could possibly leave this on soxDisconnected, and on Select Case State, thurn it on and set it ready to send data (Or set it to connecting)
        RaiseEvent OnConnection(Accept)
    End If

End Function

Public Function Address(inSox As Long) As String ' Returns the address used by a Socket (Either Local or Remote)

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
                    If apiBind(tmpSocket, tmpSocketAddr, SOCKADDR_SIZE) = SOCKET_ERROR Then 'Socket Number, Socket Address space / Name, Name Length ...
                        apiCloseSocket tmpSocket
                        Let Bind = SOCKET_ERROR
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
                                ReDim Preserve Recv(Bind) As typBuffer
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
        If apiWSAAsyncSelect(Sockets(inSox).Socket, UserControl.hWnd, 0&, 0&) = SOCKET_ERROR Then       'FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
            Call RaiseState(inSox, soxDisconnected) ' Force disconnected status, dunno what the implications are!
        Else
            If apiCloseSocket(Sockets(inSox).Socket) = SOCKET_ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                Call RaiseState(inSox, soxDisconnected) ' Force disconnected status, dunno what the implications are!
            Else
                Call RaiseState(inSox, soxDisconnected)
                RaiseEvent OnClose(inSox)
            End If
        End If
    End If

End Sub

Public Function Connect(RemoteHost As String, RemotePort As Integer) As Long 'Returns the new Sox Number / SOCKET_ERROR On Error
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
                If apiConnect(tmpSocket, tmpSocketAddr, SOCKADDR_SIZE) = SOCKET_ERROR Then
                    apiCloseSocket tmpSocket
                    Let Connect = SOCKET_ERROR
                Else
                    If apiWSAAsyncSelect(tmpSocket, UserControl.hWnd, ByVal soxCLIENT, ByVal FD_ACCEPT Or FD_CLOSE Or FD_CONNECT Or FD_READ Or FD_WRITE) = SOCKET_ERROR Then ' Reassign this Socket to Send and Receive on the DATA channel
                        apiCloseSocket tmpSocket
                        Let Connect = SOCKET_ERROR
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
                                ReDim Preserve Recv(Connect) As typBuffer
                                ReDim Preserve Send(Connect) As typBuffer
                                Let Portal.Sockets = Connect
                            End If
                        End If
                        Let Sockets(Connect).Socket = tmpSocket
                        Let Sockets(Connect).SocketAddr = tmpSocketAddr 'Set the details of the new socket/client
                        Let Sockets(Connect).uMsg = soxCLIENT ' This is a Server connection - We have connected to it (Could even be another Client computer but the fact is we connected to it)
                        Let Recv(Connect).Size = -1
                        Let Recv(Connect).Pos = -1
                        Erase Recv(Connect).Buffer
                        Let Send(Connect).Size = -1
                        Let Send(Connect).Pos = -1
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

Private Function ExtractRecv(inSox As Long) As Byte() ' Extracts complete received data from the buffer

Dim tmpBuffer() As Byte

    ReDim tmpBuffer(RecvSize(inSox))
    Call apiCopyMemory(tmpBuffer(0), Recv(inSox).Buffer(4), RecvSize(inSox) + 1)
    If Recv(inSox).Size - RecvSize(inSox) - 5 = -1 Then
        Let Recv(inSox).Size = -1
        Erase Recv(inSox).Buffer
    Else
        Let Recv(inSox).Size = Recv(inSox).Size - RecvSize(inSox) - 5
        Call apiCopyMemory(Recv(inSox).Buffer(0), Recv(inSox).Buffer(RecvSize(inSox) + 5), UBound(Recv(inSox).Buffer) - (RecvSize(inSox) + 4))
        ReDim Preserve Recv(inSox).Buffer(Recv(inSox).Size)
    End If
    Let ExtractRecv = tmpBuffer

End Function

Private Sub ExtractSend(inSox As Long) ' Just Extracts the Data from the array, no need to send it to the client like ExtractRecv as the client knows what it sent

    If Send(inSox).Size = Send(inSox).Pos Then
        Let Send(inSox).Size = -1
        Let Send(inSox).Pos = -1
        Erase Send(inSox).Buffer
    Else
        Let Send(inSox).Pos = Send(inSox).Pos - SendSize(inSox) - 5
        Let Send(inSox).Size = Send(inSox).Size - SendSize(inSox) - 5
        Call apiCopyMemory(Send(inSox).Buffer(0), Send(inSox).Buffer(SendSize(inSox) + 5), UBound(Send(inSox).Buffer) - (SendSize(inSox) + 4))
        ReDim Preserve Send(inSox).Buffer(Send(inSox).Size)
    End If

End Sub

Private Sub GetData(inSox As Long) ' Extracts data from the WinSock Recv buffers and places it in our local buffer (data() array)
Dim tmpRecvSize As Long
Dim tmpBuffer(0 To (BufferSize - 1)) As Byte 'This buffer could be optimized for small data, eg. A chat program, if you set it's size, to say 256 (0 TO 255), it could retrieve data faster

    If Not (inSox < 0 Or inSox > Portal.Sockets) Then ' Detect out of Range of our Array ...
        If Sockets(inSox).State = soxIdle Then
            Call RaiseState(inSox, soxRecv)
            If apiWSAAsyncSelect(Sockets(inSox).Socket, UserControl.hWnd, 0&, 0&) <> SOCKET_ERROR Then
                Let tmpRecvSize = apiRecv(Sockets(inSox).Socket, tmpBuffer(0), BufferSize, 0)
                DoEvents
                Select Case tmpRecvSize
                Case SOCKET_ERROR
                Case 0 ' The Socket was Gracefully closed (Never seen this happen!!! Maybe it happens in some older/newer version of WinSock API???)
                    Call RaiseState(inSox, soxDisconnected)
                    RaiseEvent OnClose(inSox)
                Case Else
                    ReDim Preserve Recv(inSox).Buffer(Recv(inSox).Size + tmpRecvSize)
                    Call apiCopyMemory(Recv(inSox).Buffer(Recv(inSox).Size + 1), tmpBuffer(0), tmpRecvSize)
                    Let Recv(inSox).Size = Recv(inSox).Size + tmpRecvSize
                    Do While Recv(inSox).Size > 2 ' If for example we received many small 'packets' of data, this will loop until we have returned/extracted all of them!
                        DoEvents
                        If Recv(inSox).Size - 3 > RecvSize(inSox) Then
                            RaiseEvent OnDataArrival(inSox, ExtractRecv(inSox))
                        Else
                            Exit Do
                        End If
                    Loop
                End Select
                If Sockets(inSox).State = soxRecv Then Call RaiseState(inSox, soxIdle) ' If this socket is closing ... we could cause HAVOK too
                apiWSAAsyncSelect Sockets(inSox).Socket, UserControl.hWnd, ByVal Sockets(inSox).uMsg, ByVal FD_CLOSE Or FD_READ Or FD_WRITE
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
            If apiGetSockOpt(Sockets(inSox).Socket, IPPROTO_TCP, Not inOption, GetOption, 4) = SOCKET_ERROR Then Let GetOption = SOCKET_ERROR
        Case Else
            If apiGetSockOpt(Sockets(inSox).Socket, SOL_SOCKET, inOption, GetOption, 4) = SOCKET_ERROR Then Let GetOption = SOCKET_ERROR
        End Select
    End If

End Function

Public Sub Hook() ' WinSock is told to send it's messages to the Sox Control, but we need to intercept these messages!

    If Portal.WndProc = 0 Then ' If it's already hooked to our WindowProc function, we could have problems, this will make sure we've UnHooked before
        Let Portal.WndProc = apiSetWindowLong(UserControl.hWnd, GWL_WNDPROC, AddressOf WindowProc)
    End If

End Sub

Public Function InIDE() As Boolean ' Nifty piece of code I found from AllAPI.Net to detect the presence of the IDE! This code and the references to it can be taken out on final compile!

    On Local Error GoTo ErrHandler
    Debug.Print 1 / 0

Exit Function

ErrHandler:
    Let InIDE = True ' Debug.Print generated an Error, that means we're in the IDE :) Cause all Debug.Print statements are removed when compiling the EXE!

End Function

'Creates a socket and sets it in listen mode. This method works only for TCP connections

Public Function Listen(inAddress As String, inPort As Integer) As Long   'Returns Sox number / SOCKET_ERROR On Error
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
                If apiBind(tmpSocket, tmpSocketAddr, SOCKADDR_SIZE) = SOCKET_ERROR Then 'Socket Number, Socket Address space / Name, Name Length ...
                    apiCloseSocket tmpSocket
                    Let Listen = SOCKET_ERROR
                Else
                    If apiListen(ByVal tmpSocket, ByVal 5) = SOCKET_ERROR Then ' 5 = Maximum connections
                        apiCloseSocket tmpSocket
                        Let Listen = SOCKET_ERROR
                    Else
                        If apiWSAAsyncSelect(tmpSocket, UserControl.hWnd, ByVal soxSERVER, ByVal FD_ACCEPT Or FD_CLOSE Or FD_CONNECT Or FD_READ Or FD_WRITE) = SOCKET_ERROR Then ' Reassign this Socket to Send and Receive on the DATA channel
                            apiCloseSocket tmpSocket
                            Let Listen = SOCKET_ERROR
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
                                    ReDim Preserve Recv(Listen) As typBuffer
                                    ReDim Preserve Send(Listen) As typBuffer
                                    Let Portal.Sockets = Listen
                                End If
                            End If
                            Let Sockets(Listen).Socket = tmpSocket
                            Let Sockets(Listen).SocketAddr = tmpSocketAddr 'Set the details of the new socket/client
                            Let Sockets(Listen).uMsg = soxSERVER
                            Let Recv(Listen).Size = -1
                            Let Recv(Listen).Pos = -1
                            Erase Recv(Listen).Buffer
                            Let Send(Listen).Size = -1
                            Let Send(Listen).Pos = -1
                            Erase Send(Listen).Buffer
                            Call RaiseState(Listen, soxListening)
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Public Function LocalHostName() As String ' The PC's Name eg. RonaldR (Needs a successful WSAStartup to function because it gets the PC name from WinSock)

    Let LocalHostName = Space(256) ' Create a 'buffer' for the API call
    If apiGetHostName(LocalHostName, 256) = SOCKET_ERROR Then
        Let LocalHostName = vbNullString
    Else
        Let LocalHostName = Trim$(LocalHostName)
    End If

End Function

Private Sub Long2Byte2(inLong As Long, inByte() As Byte) ' Similar to the above, but places the bytes direcly into the given array

    Call apiCopyMemory(inByte(0), inLong, 4)

End Sub

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

Private Function RecvSize(inSox As Long) As Long  ' Given 4 bytes, will directly copy them to a long! WARNING - To speed it up, I have no UBound checks, therefore you MUST send it 4 bytes

    Call apiCopyMemory(RecvSize, Recv(inSox).Buffer(0), 4)

End Function

Private Sub SendBuffer(inSox As Long)   'Data to be sent. For binary data, byte array should be used (for optimal performace, change inData to a byte array and only allow that datatype to be sent)

    If Not (inSox < 0 Or inSox > Portal.Sockets) Then
        If Not Send(inSox).Size = soxERROR Then ' If there is data in the Buffer ...
            If Sockets(inSox).State = soxIdle Then
                Call RaiseState(inSox, soxSend)
                apiWSAAsyncSelect Sockets(inSox).Socket, UserControl.hWnd, 0&, 0&
                Select Case Send(inSox).Size - Send(inSox).Pos
                Case Is < 0 ' We have no more data in buffer,
                Case Is = 0 ' What was this for again ???
                Case Is < BufferSize ' We have less data than our buffer size
                    If apiSend(Sockets(inSox).Socket, Send(inSox).Buffer(Send(inSox).Pos + 1), Send(inSox).Size - Send(inSox).Pos, 0) <> SOCKET_ERROR Then
                        Let Send(inSox).Pos = Send(inSox).Size ' We have sent all the data in the Buffer
                    End If
                Case Else
                    If apiSend(Sockets(inSox).Socket, Send(inSox).Buffer(Send(inSox).Pos + 1), BufferSize, 0) <> SOCKET_ERROR Then
                        Let Send(inSox).Pos = Send(inSox).Pos + BufferSize
                    End If
                End Select
                Do While Send(inSox).Size > 2 ' Meaning we can extract SendSize from it which needs a minimum of 4 (0 to 3) so we test > 2
                    DoEvents
                    If Send(inSox).Pos - 3 > SendSize(inSox) Then ' Have we sent an entire SendData command ?
                        Call ExtractSend(inSox)
                    Else
                        Exit Do
                    End If
                Loop
                If Sockets(inSox).State = soxSend Then Call RaiseState(inSox, soxIdle) ' If this socket is closing ... we could cause HAVOK ??? Right ???
                apiWSAAsyncSelect Sockets(inSox).Socket, UserControl.hWnd, ByVal Sockets(inSox).uMsg, ByVal FD_CLOSE Or FD_READ Or FD_WRITE
            End If
        End If
    End If

End Sub

Public Function SendData(inSox As Long, inData() As Byte) As Long
Dim tmpSize(0 To 3) As Byte

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let SendData = soxERROR
    Else
        If Not (Sockets(inSox).State = soxIdle Or Sockets(inSox).State = soxSend Or Sockets(inSox).State = soxRecv) Then ' If we have initiated a ShutDown, the state would change to Closing
            Let SendData = soxERROR
        Else
            If UBound(inData) = soxERROR Then ' A value of -1 is returned from UBound if there was no data
                Let SendData = soxERROR
            Else
                Call Long2Byte2(UBound(inData), tmpSize) ' I use UBound here instead of UBound + 1 to test the buffer on the other side EXACTLY !
                ReDim Preserve Send(inSox).Buffer(Send(inSox).Size + 4 + UBound(inData) + 1) As Byte ' 4 = Sized, UBound + 1 = DataLength
                Call apiCopyMemory(Send(inSox).Buffer(Send(inSox).Size + 1), tmpSize(0), 4)
                Let Send(inSox).Size = Send(inSox).Size + 4
                Call apiCopyMemory(Send(inSox).Buffer(Send(inSox).Size + 1), inData(0), UBound(inData) + 1)
                Let Send(inSox).Size = Send(inSox).Size + UBound(inData) + 1
                apiWSAAsyncSelect Sockets(inSox).Socket, UserControl.hWnd, ByVal Sockets(inSox).uMsg, ByVal FD_CLOSE Or FD_READ Or FD_WRITE
            End If
        End If
    End If

End Function

Private Function SendSize(inSox As Long) As Long  ' Given 4 bytes, will directly copy them to a long! WARNING - To speed it up, I have no UBound checks, therefore you MUST send it 4 bytes

    Call apiCopyMemory(SendSize, Send(inSox).Buffer(0), 4)

End Function

Public Function SetOption(inSox As Long, inOption As enmSoxOptions, inValue As Long) As Long

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let SetOption = soxERROR
    Else
        Select Case inOption
        Case soxSO_TCP_NODELAY
            If apiSetSockOpt(Sockets(inSox).Socket, IPPROTO_TCP, Not inOption, inValue, 4) = SOCKET_ERROR Then Let SetOption = SOCKET_ERROR
        Case Else
            If apiSetSockOpt(Sockets(inSox).Socket, SOL_SOCKET, inOption, inValue, 4) = SOCKET_ERROR Then Let SetOption = SOCKET_ERROR
        End Select
    End If

End Function

Public Function Shut(inSox As Long) As Long ' Initiates ShutDown procedure for a Socket

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let Shut = INVALID_SOCKET
    Else
        Select Case Sockets(inSox).State
        Case soxDisconnected
        Case soxClosing
        Case soxBound
            If apiWSAAsyncSelect(Sockets(inSox).Socket, UserControl.hWnd, 0&, 0&) = SOCKET_ERROR Then       'FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
                Let Shut = SOCKET_ERROR
            Else
                If apiCloseSocket(Sockets(inSox).Socket) = SOCKET_ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                    Let Shut = SOCKET_ERROR
                Else
                    Call RaiseState(inSox, soxDisconnected)
                    RaiseEvent OnClose(inSox)
                End If
            End If
        Case soxListening
            If apiWSAAsyncSelect(Sockets(inSox).Socket, UserControl.hWnd, 0&, 0&) = SOCKET_ERROR Then       'FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
                Let Shut = SOCKET_ERROR
            Else
                If apiCloseSocket(Sockets(inSox).Socket) = SOCKET_ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                    Let Shut = SOCKET_ERROR
                Else
                    Call RaiseState(inSox, soxDisconnected)
                    RaiseEvent OnClose(inSox)
                End If
            End If
        Case Else
            If apiGetSockOpt(Sockets(inSox).Socket, SOL_SOCKET, soxSO_ERROR, Shut, 4) = SOCKET_ERROR Then
                Let Shut = SOCKET_ERROR
            Else
                If apiShutDown(Sockets(inSox).Socket, SD_SEND) = SOCKET_ERROR Then
                    Let Shut = SOCKET_ERROR
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
            If apiWSAAsyncSelect(Sockets(tmpSox).Socket, UserControl.hWnd, 0&, 0&) = SOCKET_ERROR Then       'FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
                Call RaiseState(tmpSox, soxDisconnected)
                RaiseEvent OnClose(tmpSox)
                Let ShutDown = SOCKET_ERROR
            Else
                If apiCloseSocket(Sockets(tmpSox).Socket) = SOCKET_ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                    Call RaiseState(tmpSox, soxDisconnected)
                    RaiseEvent OnClose(tmpSox)
                    Let ShutDown = SOCKET_ERROR
                Else
                    Call RaiseState(tmpSox, soxDisconnected)
                    RaiseEvent OnClose(tmpSox)
                End If
            End If
        Case soxListening ' Same as soxBound
            If apiWSAAsyncSelect(Sockets(tmpSox).Socket, UserControl.hWnd, 0&, 0&) = SOCKET_ERROR Then       'FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
                Call RaiseState(tmpSox, soxDisconnected)
                RaiseEvent OnClose(tmpSox)
                Let ShutDown = SOCKET_ERROR
            Else
                If apiCloseSocket(Sockets(tmpSox).Socket) = SOCKET_ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                    Call RaiseState(tmpSox, soxDisconnected)
                    RaiseEvent OnClose(tmpSox)
                    Let ShutDown = SOCKET_ERROR
                Else
                    Call RaiseState(tmpSox, soxDisconnected)
                    RaiseEvent OnClose(tmpSox)
                End If
            End If
        Case Else
            If apiShutDown(Sockets(tmpSox).Socket, SD_SEND) = SOCKET_ERROR Then
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
        Erase Recv
        Erase Send
    End If

End Function

Private Function Socket2Sox(inSocket As Long) As Long ' Returns the Sockets() Array address of a WinSock Socket

    For Socket2Sox = 0 To Portal.Sockets
        If Sockets(Socket2Sox).Socket = inSocket Then Exit For
    Next Socket2Sox
    If Socket2Sox = Portal.Sockets + 1 Then Let Socket2Sox = INVALID_SOCKET

End Function

Public Function SocketHandle(inSox As Long) As Long

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let SocketHandle = soxERROR
    Else
        Let SocketHandle = Sockets(inSox).Socket
    End If

End Function

Public Function State(inSox As Long) As enmSoxState

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

    Call apiSetWindowLong(UserControl.hWnd, GWL_WNDPROC, Portal.WndProc)
    Let Portal.WndProc = 0

End Sub

Private Sub UserControl_Initialize() ' NB: This code actually gets processed in the IDE as well!

    If Not InIDE Then ' The following code gets run EVERY time we double click the Control and frmMain, this creates problems for us!!! So, this prevents us from using WinSock and SubClassing from within the IDE!!!
        Set SoxControl = Me
        If apiWSAStartup(&H101, WSAData) = SOCKET_ERROR Then
            Call MsgBox("WinSock failed to initialize properly - Error#: " & Err.LastDllError, vbApplicationModal + vbCritical, "Critical Error")  'Creates an 'application instance' and memory space in the WinSock DLL (MUST be cleaned up later)
        Else
            Let Portal.WndProc = apiSetWindowLong(UserControl.hWnd, GWL_WNDPROC, AddressOf WindowProc)
            Let Portal.Sockets = -1 ' Initialize our socket count ... NB - WE HAVE NONE, used wherever we Redim the Sockets Array
        End If
    Else
        Call MsgBox("NB: WinSock API is currently disabled!" & vbCrLf & "Possible VB crash avoided!", vbApplicationModal + vbExclamation, "Warning") ' Disable this line if you get tired of seeing this EVERY SINGLE friggan time this code runs!!!
        Let Portal.Sockets = -1 ' Initialize our socket (the array) count ... NB - WE HAVE NONE, used wherever we Redim the Sockets Array
    End If

End Sub

Private Sub UserControl_Resize()

    Let UserControl.Width = Screen.TwipsPerPixelX * 28
    Let UserControl.Height = Screen.TwipsPerPixelX * 28

End Sub

Private Sub UserControl_Terminate()

    If Not InIDE Then
        'Correctly replaces/reattaches the origional WindowProc procedure to our 'hidden' handle (Basically what the UnHook command does!)
        Call apiSetWindowLong(UserControl.hWnd, GWL_WNDPROC, Portal.WndProc)
        If apiWSACleanup = SOCKET_ERROR Then Call MsgBox("WinSock failed to terminate properly, memory leak imminent - Error#: " & Err.LastDllError, vbApplicationModal + vbCritical, "Critical Error")   'If cleanup failed, does not / cannot raise errors
    End If

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

Friend Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case uMsg
    Case soxSERVER
        Select Case WinSockEvent(lParam)
        Case FD_ACCEPT
            If WinSockError(lParam) = 0 Then Call Accept(wParam)
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
    Case Else: Let WndProc = apiCallWindowProc(Portal.WndProc, hWnd, uMsg, wParam, lParam)
    End Select

End Function
