Attribute VB_Name = "GOREsock"
Option Explicit

Private Const BufferSize As Long = 8192

Private WindowhWnd As Long
Public Loaded As Byte   'Whether the socket is loaded or not

Private Type typPortal 'Class specific variables
    GOREsock_WndProc As Long 'Pointer to the origional WindowProc of our window (We need to give control of ALL messages back to it before we destroy it)
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
    soxBound = 10& ' The socket has been bound to its current GOREsock_Port and GOREsock_Address
    soxERROR = -1&
End Enum
#If False Then
Private soxDisconnected, soxListening, soxConnecting, soxIdle, soxSend, soxRecv, soxClosing, soxBound, soxERROR
#End If

Public Enum enmSoxOptions
    soxSO_BROADCAST = &H20& 'BOOL Allow transmission of broadcast messages on the GOREsock_
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
    Socket As Long ' The actual WinSock API socket number
    SocketAddr As typSocketAddr ' Info about the connection
    GOREsock_State As enmSoxState
    GOREsock_uMsg As Long ' Server (-1) / Client (0) Socket (Server = A Socket that has a connection to the Server / Client = A Socket that was created in GOREsock_Accept that connected to us)
End Type

Private Type typBuffer ' The advantage of using this is if we sent exactly 8K on the other side, when we receive 8K, FD_READ will not be sent again so we won't get an error like when we use a loop
    Size As Long ' Array Size (To check if there is incomming data, we can check the size of this variable, if -1 then we are not receiving anything)
    Pos As Long
    buffer() As Byte
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
Private Const GOREsock_ERROR As Long = -1&
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
Private Const GWL_GOREsock_WndProc As Long = (-4)

'Combinations of flags (process it as a const instead of real-time - slightly (oh so slightly) faster)
Private Const FD_CLOSEREADWRITE As Long = FD_CLOSE Or FD_READ Or FD_WRITE
Private Const FD_CONNECTLISTEN As Long = FD_ACCEPT Or FD_CLOSE Or FD_CONNECT Or FD_READ Or FD_WRITE

Private Declare Function apiWSAStartup Lib "WS2_32" Alias "WSAStartup" (ByVal wVersionRequired As Long, lpWSADATA As typWSAData) As Long
Private Declare Function apiWSACleanup Lib "WS2_32" Alias "WSACleanup" () As Long
Private Declare Function apiSocket Lib "WS2_32" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Private Declare Function apiCloseSocket Lib "WS2_32" Alias "closesocket" (ByVal S As Long) As Long
Private Declare Function apiBind Lib "WS2_32" Alias "bind" (ByVal S As Long, addr As typSocketAddr, ByVal namelen As Long) As Long
Private Declare Function apiListen Lib "WS2_32" Alias "listen" (ByVal S As Long, ByVal backlog As Long) As Long
Private Declare Function apiConnect Lib "WS2_32" Alias "connect" (ByVal S As Long, name As typSocketAddr, ByVal namelen As Long) As Long
Private Declare Function apiAccept Lib "WS2_32" Alias "accept" (ByVal S As Long, addr As typSocketAddr, addrlen As Long) As Long
Private Declare Function apiWSAAsyncSelect Lib "WS2_32" Alias "WSAAsyncSelect" (ByVal S As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function apiRecv Lib "WS2_32" Alias "recv" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function apiSend Lib "WS2_32" Alias "send" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function apiGetSockOpt Lib "WS2_32" Alias "getsockopt" (ByVal S As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function apiSetSockOpt Lib "WS2_32" Alias "setsockopt" (ByVal S As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function apiHToNS Lib "WS2_32" Alias "htons" (ByVal hostshort As Long) As Integer 'Host To Network Short
Private Declare Function apiNToHS Lib "WS2_32" Alias "ntohs" (ByVal netshort As Long) As Integer 'Network To Host Short
Private Declare Function apiIPToNL Lib "WS2_32" Alias "inet_addr" (ByVal cp As String) As Long
Private Declare Function apiNLToIP Lib "WS2_32" Alias "inet_ntoa" (ByVal inn As Long) As Long
Private Declare Function apiShutDown Lib "WS2_32" Alias "shutdown" (ByVal S As Long, ByVal how As Long) As Long
Private Declare Function apiCallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function apiSetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function apiLStrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function apiLstrCpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Sub apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private WSAData As typWSAData 'Stores WinSock data on initialization of WinSock 2
Private Portal As typPortal ' Sorta used for general variables (The word Portal came from my use in previous versions of a STATIC window, do you know what that is? :)
Private Sockets() As typSocket
Private Recv() As typBuffer ' Receive Buffer
Private Send() As typBuffer ' Send Buffer

Private Function GOREsock_Accept(inSocket As Long) As Long 'Returns: New Sox Number -- inSocket is the listening WinSocket, not Sox number ...
Dim tmpSocket As Long
Dim tmpSocketAddr As typSocketAddr 'This stores the details of our new socket/client, including the client IP address

    Let tmpSocket = apiAccept(inSocket, tmpSocketAddr, SOCKADDR_SIZE) 'GOREsock_Accept API returns a valid, random, unused socket for us to use for the new client
    If tmpSocket = INVALID_SOCKET Then 'GOREsock_Accept API may not give us a valid socket eg. when all sockets are full, you may have to add additional error trapping if you believe you will use over 32,767 sockets
        'Since a socket was not commited for the new Connection ... we don't have to close it (Since the socket was never even created)
        Let GOREsock_Accept = INVALID_SOCKET
    Else ' Success, A new connection ... GOREsock_Accept now contains the new Socket number
        For GOREsock_Accept = 0 To Portal.Sockets ' First search to see if the socket already exists
            If Sockets(GOREsock_Accept).Socket = tmpSocket Then Exit For
        Next GOREsock_Accept
        If GOREsock_Accept = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
            For GOREsock_Accept = 0 To Portal.Sockets ' First search to see if the socket already exists
                If Sockets(GOREsock_Accept).Socket = soxDisconnected Then Exit For ' Found an open Socket
            Next GOREsock_Accept
            If GOREsock_Accept = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                ReDim Preserve Sockets(GOREsock_Accept) As typSocket
                ReDim Preserve Recv(GOREsock_Accept) As typBuffer
                ReDim Preserve Send(GOREsock_Accept) As typBuffer
                Let Portal.Sockets = GOREsock_Accept
            End If
        End If
        Let Sockets(GOREsock_Accept).Socket = tmpSocket
        Let Sockets(GOREsock_Accept).SocketAddr = tmpSocketAddr 'Set the details of the new socket/client
        Let Sockets(GOREsock_Accept).GOREsock_uMsg = soxSERVER  'This is a Client Socket - It has connected to US
        Let Recv(GOREsock_Accept).Size = -1
        Let Recv(GOREsock_Accept).Pos = -1
        Erase Recv(GOREsock_Accept).buffer
        Let Send(GOREsock_Accept).Size = -1
        Let Send(GOREsock_Accept).Pos = -1
        Erase Send(GOREsock_Accept).buffer
        Call GOREsock_RaiseState(GOREsock_Accept, soxConnecting) ' Could possibly leave this on soxDisconnected, and on Select Case GOREsock_State, thurn it on and set it ready to send data (Or set it to connecting)
        Call GOREsock_Connection(GOREsock_Accept)
    End If

End Function

Public Function GOREsock_Address(inSox As Long) As String ' Returns the address used by a Socket (Either Local or Remote)

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let GOREsock_Address = soxERROR
    Else
        Let GOREsock_Address = GOREsock_StringFromPointer(apiNLToIP(Sockets(inSox).SocketAddr.sin_addr))
    End If

End Function

Public Function GOREsock_Bind(LocalIP As String, LocalPort As Integer) As Long
Dim tmpSocket As Long
Dim tmpSocketAddr As typSocketAddr

    If LocalPort = 0 Or LocalIP = vbNullString Then
        Let GOREsock_Bind = soxERROR
    Else
        Let tmpSocketAddr.sin_family = AF_INET
        Let tmpSocketAddr.sin_port = apiHToNS(LocalPort)
        If tmpSocketAddr.sin_port = INVALID_SOCKET Then
            Let GOREsock_Bind = INVALID_SOCKET
        Else
            Let tmpSocketAddr.sin_addr = apiIPToNL(LocalIP) 'If this is Zero, it will assign 0.0.0.0 !!!
            If tmpSocketAddr.sin_addr = INADDR_NONE Then 'If 255.255.255.255 is returned ... we have a problem ... I think :)
                Let GOREsock_Bind = INVALID_SOCKET
            Else
                Let tmpSocket = apiSocket(AF_INET, SOCK_STREAM, IPPROTO_TCP) 'This is where you specify what type of protocol to use and what type of Streaming to use, returns a new socket number 4 us (NB - From here, if any further steps fail after this one succeeds, we must close the socket)
                If tmpSocket = INVALID_SOCKET Then
                    Let GOREsock_Bind = INVALID_SOCKET
                Else
                    If apiBind(tmpSocket, tmpSocketAddr, SOCKADDR_SIZE) = GOREsock_ERROR Then 'Socket Number, Socket GOREsock_Address space / Name, Name Length ...
                        apiCloseSocket tmpSocket
                        Let GOREsock_Bind = GOREsock_ERROR
                    Else
                        For GOREsock_Bind = 0 To Portal.Sockets ' First search to see if the socket already exists
                            If Sockets(GOREsock_Bind).Socket = tmpSocket Then Exit For
                        Next GOREsock_Bind
                        If GOREsock_Bind = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                            For GOREsock_Bind = 0 To Portal.Sockets ' First search to see if the socket already exists
                                If Sockets(GOREsock_Bind).Socket = soxDisconnected Then Exit For ' Found an open Socket
                            Next GOREsock_Bind
                            If GOREsock_Bind = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                                ReDim Preserve Sockets(GOREsock_Bind) As typSocket
                                ReDim Preserve Recv(GOREsock_Bind) As typBuffer
                                ReDim Preserve Send(GOREsock_Bind) As typBuffer
                                Let Portal.Sockets = GOREsock_Bind
                            End If
                        End If
                        Let Sockets(GOREsock_Bind).Socket = tmpSocket
                        Let Sockets(GOREsock_Bind).SocketAddr = tmpSocketAddr 'Set the details of the new socket/client
                        Call GOREsock_RaiseState(GOREsock_Bind, soxBound)
                    End If
                End If
            End If
        End If
    End If

End Function

Private Sub GOREsock_Closed(inSox As Long) ' This Socket has successfully closed ... free resources (No need to check if it exists, cause we call this internally)

    If Not (inSox < 0 Or inSox > Portal.Sockets) Then ' Detect out of Range of our Array ...
        If apiWSAAsyncSelect(Sockets(inSox).Socket, WindowhWnd, 0&, 0&) = GOREsock_ERROR Then       'FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
            Call GOREsock_RaiseState(inSox, soxDisconnected) ' Force disconnected status, dunno what the implications are!
        Else
            If apiCloseSocket(Sockets(inSox).Socket) = GOREsock_ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                Call GOREsock_RaiseState(inSox, soxDisconnected) ' Force disconnected status, dunno what the implications are!
            Else
                Call GOREsock_RaiseState(inSox, soxDisconnected)
                Call GOREsock_Close(inSox)
            End If
        End If
    End If

End Sub

Public Function GOREsock_Connect(RemoteHost As String, RemotePort As Integer) As Long 'Returns the new Sox Number / GOREsock_ERROR On Error
Dim tmpSocket As Long
Dim tmpSocketAddr As typSocketAddr

    Let tmpSocketAddr.sin_family = AF_INET
    Let tmpSocketAddr.sin_port = apiHToNS(RemotePort) ' apiHToNS(RemotePort)
    If tmpSocketAddr.sin_port = INVALID_SOCKET Then
        Let GOREsock_Connect = INVALID_SOCKET
    Else
        Let tmpSocketAddr.sin_addr = apiIPToNL(RemoteHost) 'If this is Zero, it will assign 0.0.0.0 !!!
        If tmpSocketAddr.sin_addr = INADDR_NONE Then 'If 255.255.255.255 is returned ... we have a problem ... I think :)
            Let GOREsock_Connect = INVALID_SOCKET
        Else
            Let tmpSocket = apiSocket(AF_INET, SOCK_STREAM, IPPROTO_TCP) 'This is where you specify what type of protocol to use and what type of Streaming to use, returns a new socket number 4 us (NB - From here, if any further steps fail after this one succeeds, we must close the socket)
            If tmpSocket = INVALID_SOCKET Then
                Let GOREsock_Connect = INVALID_SOCKET
            Else
                If apiConnect(tmpSocket, tmpSocketAddr, SOCKADDR_SIZE) = GOREsock_ERROR Then
                    apiCloseSocket tmpSocket
                    Let GOREsock_Connect = GOREsock_ERROR
                Else
                    If apiWSAAsyncSelect(tmpSocket, WindowhWnd, ByVal soxCLIENT, ByVal FD_CONNECTLISTEN) = GOREsock_ERROR Then ' Reassign this Socket to Send and Receive on the DATA channel
                        apiCloseSocket tmpSocket
                        Let GOREsock_Connect = GOREsock_ERROR
                    Else
                        For GOREsock_Connect = 0 To Portal.Sockets ' First search to see if the socket already exists
                            If Sockets(GOREsock_Connect).Socket = tmpSocket Then Exit For
                        Next GOREsock_Connect
                        If GOREsock_Connect = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                            For GOREsock_Connect = 0 To Portal.Sockets ' First search to see if the socket already exists
                                If Sockets(GOREsock_Connect).Socket = soxDisconnected Then Exit For ' Found an open Socket
                            Next GOREsock_Connect
                            If GOREsock_Connect = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                                ReDim Preserve Sockets(GOREsock_Connect) As typSocket
                                ReDim Preserve Recv(GOREsock_Connect) As typBuffer
                                ReDim Preserve Send(GOREsock_Connect) As typBuffer
                                Let Portal.Sockets = GOREsock_Connect
                            End If
                        End If
                        Let Sockets(GOREsock_Connect).Socket = tmpSocket
                        Let Sockets(GOREsock_Connect).SocketAddr = tmpSocketAddr 'Set the details of the new socket/client
                        Let Sockets(GOREsock_Connect).GOREsock_uMsg = soxCLIENT ' This is a Server connection - We have connected to it (Could even be another Client computer but the fact is we connected to it)
                        Let Recv(GOREsock_Connect).Size = -1
                        Let Recv(GOREsock_Connect).Pos = -1
                        Erase Recv(GOREsock_Connect).buffer
                        Let Send(GOREsock_Connect).Size = -1
                        Let Send(GOREsock_Connect).Pos = -1
                        Erase Send(GOREsock_Connect).buffer
                        Call GOREsock_RaiseState(GOREsock_Connect, soxConnecting)
                    End If
                End If
            End If
        End If
    End If

End Function

Public Function GOREsock_Connections() As Long ' Returns the number of clients connected to Sox

Dim tmpLoop As Long

    For tmpLoop = 0 To Portal.Sockets
        If Not Sockets(tmpLoop).GOREsock_State = soxDisconnected Then
            If Sockets(tmpLoop).GOREsock_uMsg = soxSERVER Then
                Select Case Sockets(tmpLoop).GOREsock_State ' These are the valid states for 'connected' sockets
                Case soxConnecting: Let GOREsock_Connections = GOREsock_Connections + 1
                Case soxIdle: Let GOREsock_Connections = GOREsock_Connections + 1
                Case soxClosing: Let GOREsock_Connections = GOREsock_Connections + 1
                Case soxRecv: Let GOREsock_Connections = GOREsock_Connections + 1
                Case soxSend: Let GOREsock_Connections = GOREsock_Connections + 1
                End Select
            End If
        End If
    Next tmpLoop

End Function

Private Function GOREsock_ExtractRecv(inSox As Long) As Byte() ' Extracts complete received data from the buffer

Dim tmpBuffer() As Byte

    ReDim tmpBuffer(GOREsock_RecvSize(inSox)) As Byte
    Call apiCopyMemory(tmpBuffer(0), Recv(inSox).buffer(4), GOREsock_RecvSize(inSox) + 1)
    If Recv(inSox).Size - GOREsock_RecvSize(inSox) - 5 = -1 Then
        Let Recv(inSox).Size = -1
        Erase Recv(inSox).buffer
    Else
        Let Recv(inSox).Size = Recv(inSox).Size - GOREsock_RecvSize(inSox) - 5
        Call apiCopyMemory(Recv(inSox).buffer(0), Recv(inSox).buffer(GOREsock_RecvSize(inSox) + 5), UBound(Recv(inSox).buffer) - (GOREsock_RecvSize(inSox) + 4))
        ReDim Preserve Recv(inSox).buffer(Recv(inSox).Size)
    End If
    Let GOREsock_ExtractRecv = tmpBuffer

End Function

Private Sub GOREsock_ExtractSend(inSox As Long) ' Just Extracts the Data from the array, no need to send it to the client like GOREsock_ExtractRecv as the client knows what it sent

    If Send(inSox).Size = Send(inSox).Pos Then
        Let Send(inSox).Size = -1
        Let Send(inSox).Pos = -1
        Erase Send(inSox).buffer
    Else
        Let Send(inSox).Pos = Send(inSox).Pos - GOREsock_SendSize(inSox) - 5
        Let Send(inSox).Size = Send(inSox).Size - GOREsock_SendSize(inSox) - 5
        Call apiCopyMemory(Send(inSox).buffer(0), Send(inSox).buffer(GOREsock_SendSize(inSox) + 5), UBound(Send(inSox).buffer) - (GOREsock_SendSize(inSox) + 4))
        ReDim Preserve Send(inSox).buffer(Send(inSox).Size)
    End If

End Sub

Private Sub GOREsock_GetData(inSox As Long) ' Extracts data from the WinSock Recv buffers and places it in our local buffer (data() array)
Dim tmpRecvSize As Long
Dim tmpBuffer(0 To (BufferSize - 1)) As Byte 'This buffer could be optimized for small data, eg. A chat program, if you set it's size, to say 256 (0 TO 255), it could retrieve data faster

    If Not (inSox < 0 Or inSox > Portal.Sockets) Then ' Detect out of Range of our Array ...
        If Sockets(inSox).GOREsock_State = soxIdle Then
            Call GOREsock_RaiseState(inSox, soxRecv)
            If apiWSAAsyncSelect(Sockets(inSox).Socket, WindowhWnd, 0&, 0&) <> GOREsock_ERROR Then
                Let tmpRecvSize = apiRecv(Sockets(inSox).Socket, tmpBuffer(0), BufferSize, 0)
                DoEvents
                Select Case tmpRecvSize
                Case GOREsock_ERROR
                Case 0 ' The Socket was Gracefully closed (Never seen this happen!!! Maybe it happens in some older/newer version of WinSock API???)
                    Call GOREsock_RaiseState(inSox, soxDisconnected)
                    Call GOREsock_Close(inSox)
                Case Else
                    ReDim Preserve Recv(inSox).buffer(Recv(inSox).Size + tmpRecvSize)
                    Call apiCopyMemory(Recv(inSox).buffer(Recv(inSox).Size + 1), tmpBuffer(0), tmpRecvSize)
                    Let Recv(inSox).Size = Recv(inSox).Size + tmpRecvSize
                    Do While Recv(inSox).Size > 2 ' If for example we received many small 'packets' of data, this will loop until we have returned/extracted all of them!
                        DoEvents
                        If Recv(inSox).Size - 3 > GOREsock_RecvSize(inSox) Then
                            Call GOREsock_RecvProgress(inSox, Recv(inSox).Size - 3, 0)
                            Call GOREsock_DataArrival(inSox, GOREsock_ExtractRecv(inSox))
                        Else
                            Call GOREsock_RecvProgress(inSox, Recv(inSox).Size - 3, (GOREsock_RecvSize(inSox) + 1) - (Recv(inSox).Size - 3))
                            Exit Do
                        End If
                    Loop
                End Select
                If Sockets(inSox).GOREsock_State = soxRecv Then Call GOREsock_RaiseState(inSox, soxIdle) ' If this socket is closing ... we could cause HAVOK too
                apiWSAAsyncSelect Sockets(inSox).Socket, WindowhWnd, ByVal Sockets(inSox).GOREsock_uMsg, ByVal FD_CLOSEREADWRITE
            End If
        End If
    End If

End Sub

Public Function GOREsock_GetOption(inSox As Long, inOption As enmSoxOptions) As Long

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let GOREsock_GetOption = soxERROR
    Else
        Select Case inOption
        Case soxSO_TCP_NODELAY
            If apiGetSockOpt(Sockets(inSox).Socket, IPPROTO_TCP, Not inOption, GOREsock_GetOption, 4) = GOREsock_ERROR Then Let GOREsock_GetOption = GOREsock_ERROR
        Case Else
            If apiGetSockOpt(Sockets(inSox).Socket, SOL_SOCKET, inOption, GOREsock_GetOption, 4) = GOREsock_ERROR Then Let GOREsock_GetOption = GOREsock_ERROR
        End Select
    End If

End Function

Public Sub GOREsock_Hook() ' WinSock is told to send it's messages to the Sox Control, but we need to intercept these messages!

    If Portal.GOREsock_WndProc = 0 Then ' If it's already hooked to our WindowProc function, we could have problems, this will make sure we've UnHooked before
        Let Portal.GOREsock_WndProc = apiSetWindowLong(WindowhWnd, GWL_GOREsock_WndProc, AddressOf GOREsock_WndProc)
    End If

End Sub

'Creates a socket and sets it in listen mode. This method works only for TCP connections

Public Function GOREsock_Listen(inAddress As String, inPort As Integer) As Long   'Returns Sox number / GOREsock_ERROR On Error
Dim tmpSocket As Long
Dim tmpSocketAddr As typSocketAddr

    Let tmpSocketAddr.sin_family = AF_INET
    Let tmpSocketAddr.sin_port = apiHToNS(inPort)
    If tmpSocketAddr.sin_port = INVALID_SOCKET Then
        Let GOREsock_Listen = INVALID_SOCKET
    Else
        Let tmpSocketAddr.sin_addr = apiIPToNL(inAddress) 'If this is Zero, it will assign 0.0.0.0 !!!
        If tmpSocketAddr.sin_addr = INADDR_NONE Then 'If 255.255.255.255 is returned ... we have a problem ... I think :)
            Let GOREsock_Listen = INVALID_SOCKET
        Else
            Let tmpSocket = apiSocket(AF_INET, SOCK_STREAM, IPPROTO_TCP) 'This is where you specify what type of protocol to use and what type of Streaming to use, returns a new socket number 4 us (NB - From here, if any further steps fail after this one succeeds, we must close the socket)
            If tmpSocket = INVALID_SOCKET Then
                Let GOREsock_Listen = INVALID_SOCKET
            Else
                If apiBind(tmpSocket, tmpSocketAddr, SOCKADDR_SIZE) = GOREsock_ERROR Then 'Socket Number, Socket GOREsock_Address space / Name, Name Length ...
                    apiCloseSocket tmpSocket
                    Let GOREsock_Listen = GOREsock_ERROR
                Else
                    If apiListen(ByVal tmpSocket, ByVal 5) = GOREsock_ERROR Then ' 5 = Maximum connections
                        apiCloseSocket tmpSocket
                        Let GOREsock_Listen = GOREsock_ERROR
                    Else
                        If apiWSAAsyncSelect(tmpSocket, WindowhWnd, ByVal soxSERVER, ByVal FD_CONNECTLISTEN) = GOREsock_ERROR Then ' Reassign this Socket to Send and Receive on the DATA channel
                            apiCloseSocket tmpSocket
                            Let GOREsock_Listen = GOREsock_ERROR
                        Else
                            For GOREsock_Listen = 0 To Portal.Sockets ' First search to see if the socket already exists
                                If Sockets(GOREsock_Listen).Socket = tmpSocket Then Exit For
                            Next GOREsock_Listen
                            If GOREsock_Listen = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                                For GOREsock_Listen = 0 To Portal.Sockets ' First search to see if the socket already exists
                                    If Sockets(GOREsock_Listen).Socket = soxDisconnected Then Exit For ' Found an open Socket
                                Next GOREsock_Listen
                                If GOREsock_Listen = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                                    ReDim Preserve Sockets(GOREsock_Listen) As typSocket
                                    ReDim Preserve Recv(GOREsock_Listen) As typBuffer
                                    ReDim Preserve Send(GOREsock_Listen) As typBuffer
                                    Let Portal.Sockets = GOREsock_Listen
                                End If
                            End If
                            Let Sockets(GOREsock_Listen).Socket = tmpSocket
                            Let Sockets(GOREsock_Listen).SocketAddr = tmpSocketAddr 'Set the details of the new socket/client
                            Let Sockets(GOREsock_Listen).GOREsock_uMsg = soxSERVER
                            Let Recv(GOREsock_Listen).Size = -1
                            Let Recv(GOREsock_Listen).Pos = -1
                            Erase Recv(GOREsock_Listen).buffer
                            Let Send(GOREsock_Listen).Size = -1
                            Let Send(GOREsock_Listen).Pos = -1
                            Erase Send(GOREsock_Listen).buffer
                            Call GOREsock_RaiseState(GOREsock_Listen, soxListening)
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Private Sub GOREsock_Long2Byte2(inLong As Long, inByte() As Byte) ' Similar to the above, but places the bytes direcly into the given array

    Call apiCopyMemory(inByte(0), inLong, 4)

End Sub

Public Function GOREsock_Port(inSox As Long) As Long

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let GOREsock_Port = soxERROR
    Else
        Let GOREsock_Port = apiNToHS(Sockets(inSox).SocketAddr.sin_port)
    End If

End Function

Private Sub GOREsock_RaiseState(inSox As Long, inState As enmSoxState)

    Let Sockets(inSox).GOREsock_State = inState
    If inState = soxConnecting Then Call GOREsock_Connecting(inSox)

End Sub

Private Function GOREsock_RecvSize(inSox As Long) As Long  ' Given 4 bytes, will directly copy them to a long! WARNING - To speed it up, I have no UBound checks, therefore you MUST send it 4 bytes

    Call apiCopyMemory(GOREsock_RecvSize, Recv(inSox).buffer(0), 4)

End Function

Private Sub GOREsock_SendBuffer(inSox As Long)   'Data to be sent. For binary data, byte array should be used (for optimal performace, change inData to a byte array and only allow that datatype to be sent)

    If Not (inSox < 0 Or inSox > Portal.Sockets) Then
        If Not Send(inSox).Size = soxERROR Then ' If there is data in the Buffer ...
            If Sockets(inSox).GOREsock_State = soxIdle Then
                Call GOREsock_RaiseState(inSox, soxSend)
                apiWSAAsyncSelect Sockets(inSox).Socket, WindowhWnd, 0&, 0&
                Select Case Send(inSox).Size - Send(inSox).Pos
                Case Is < 0 ' We have no more data in buffer,
                Case Is = 0 ' What was this for again ???
                Case Is < BufferSize ' We have less data than our buffer size
                    If apiSend(Sockets(inSox).Socket, Send(inSox).buffer(Send(inSox).Pos + 1), Send(inSox).Size - Send(inSox).Pos, 0) <> GOREsock_ERROR Then
                        Let Send(inSox).Pos = Send(inSox).Size ' We have sent all the data in the Buffer
                    End If
                Case Else
                    If apiSend(Sockets(inSox).Socket, Send(inSox).buffer(Send(inSox).Pos + 1), BufferSize, 0) <> GOREsock_ERROR Then
                        Let Send(inSox).Pos = Send(inSox).Pos + BufferSize
                    End If
                End Select
                Do While Send(inSox).Size > 2 ' Meaning we can extract GOREsock_SendSize from it which needs a minimum of 4 (0 to 3) so we test > 2
                    DoEvents
                    If Send(inSox).Pos - 3 > GOREsock_SendSize(inSox) Then ' Have we sent an entire GOREsock_SendData command ?
                        Call GOREsock_SendProgress(inSox, Send(inSox).Pos - 3, 0)
                        Call GOREsock_ExtractSend(inSox)
                        Call GOREsock_SendComplete(inSox)
                    Else
                        Call GOREsock_SendProgress(inSox, Send(inSox).Pos - 3, (GOREsock_SendSize(inSox) + 1) - (Send(inSox).Pos - 3))
                        Exit Do
                    End If
                Loop
                If Sockets(inSox).GOREsock_State = soxSend Then Call GOREsock_RaiseState(inSox, soxIdle) ' If this socket is closing ... we could cause HAVOK ??? Right ???
                apiWSAAsyncSelect Sockets(inSox).Socket, WindowhWnd, ByVal Sockets(inSox).GOREsock_uMsg, ByVal FD_CLOSEREADWRITE
            End If
        End If
    End If

End Sub

Public Function GOREsock_SendData(inSox As Long, inData() As Byte) As Long
Dim tmpSize(0 To 3) As Byte

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let GOREsock_SendData = soxERROR
    Else
        If Not (Sockets(inSox).GOREsock_State = soxIdle Or Sockets(inSox).GOREsock_State = soxSend Or Sockets(inSox).GOREsock_State = soxRecv) Then ' If we have initiated a GOREsock_ShutDown, the state would change to Closing
            Let GOREsock_SendData = soxERROR
        Else
            If UBound(inData) = soxERROR Then ' A value of -1 is returned from UBound if there was no data
                Let GOREsock_SendData = soxERROR
            Else
                Call GOREsock_Long2Byte2(UBound(inData), tmpSize) ' I use UBound here instead of UBound + 1 to test the buffer on the other side EXACTLY !
                ReDim Preserve Send(inSox).buffer(Send(inSox).Size + 4 + UBound(inData) + 1) As Byte ' 4 = Sized, UBound + 1 = DataLength
                Call apiCopyMemory(Send(inSox).buffer(Send(inSox).Size + 1), tmpSize(0), 4)
                Let Send(inSox).Size = Send(inSox).Size + 4
                Call apiCopyMemory(Send(inSox).buffer(Send(inSox).Size + 1), inData(0), UBound(inData) + 1)
                Let Send(inSox).Size = Send(inSox).Size + UBound(inData) + 1
                apiWSAAsyncSelect Sockets(inSox).Socket, WindowhWnd, ByVal Sockets(inSox).GOREsock_uMsg, ByVal FD_CLOSEREADWRITE
            End If
        End If
    End If

End Function

Private Function GOREsock_SendSize(inSox As Long) As Long  ' Given 4 bytes, will directly copy them to a long! WARNING - To speed it up, I have no UBound checks, therefore you MUST send it 4 bytes

    Call apiCopyMemory(GOREsock_SendSize, Send(inSox).buffer(0), 4)

End Function

Public Function GOREsock_SetOption(inSox As Long, inOption As enmSoxOptions, inValue As Long) As Long

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let GOREsock_SetOption = soxERROR
    Else
        Select Case inOption
        Case soxSO_TCP_NODELAY
            If apiSetSockOpt(Sockets(inSox).Socket, IPPROTO_TCP, Not inOption, inValue, 4) = GOREsock_ERROR Then Let GOREsock_SetOption = GOREsock_ERROR
        Case Else
            If apiSetSockOpt(Sockets(inSox).Socket, SOL_SOCKET, inOption, inValue, 4) = GOREsock_ERROR Then Let GOREsock_SetOption = GOREsock_ERROR
        End Select
    End If

End Function

Public Function GOREsock_Shut(inSox As Long) As Long ' Initiates GOREsock_ShutDown procedure for a Socket

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let GOREsock_Shut = INVALID_SOCKET
    Else
        Select Case Sockets(inSox).GOREsock_State
        Case soxDisconnected
        Case soxClosing
        Case soxBound
            If apiWSAAsyncSelect(Sockets(inSox).Socket, WindowhWnd, 0&, 0&) = GOREsock_ERROR Then
                Let GOREsock_Shut = GOREsock_ERROR
            Else
                If apiCloseSocket(Sockets(inSox).Socket) = GOREsock_ERROR Then
                    Let GOREsock_Shut = GOREsock_ERROR
                Else
                    Call GOREsock_RaiseState(inSox, soxDisconnected)
                    Call GOREsock_Close(inSox)
                End If
            End If
        Case soxListening
            If apiWSAAsyncSelect(Sockets(inSox).Socket, WindowhWnd, 0&, 0&) = GOREsock_ERROR Then
                Let GOREsock_Shut = GOREsock_ERROR
            Else
                If apiCloseSocket(Sockets(inSox).Socket) = GOREsock_ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                    Let GOREsock_Shut = GOREsock_ERROR
                Else
                    Call GOREsock_RaiseState(inSox, soxDisconnected)
                    Call GOREsock_Close(inSox)
                End If
            End If
        Case Else
            If apiGetSockOpt(Sockets(inSox).Socket, SOL_SOCKET, soxSO_ERROR, GOREsock_Shut, 4) = GOREsock_ERROR Then
                Let GOREsock_Shut = GOREsock_ERROR
            Else
                If apiShutDown(Sockets(inSox).Socket, SD_SEND) = GOREsock_ERROR Then
                    Let GOREsock_Shut = GOREsock_ERROR
                Else
                    Call GOREsock_RaiseState(inSox, soxClosing)
                End If
            End If
        End Select
    End If

End Function

' Returns the GOREsock_ShutDown 'Go Ahead' ...

Public Function GOREsock_ShutDown() As Long ' Closes Listening and Bound Sockets immediately, sends apiShutDown to the rest
Dim tmpSox As Long

    For tmpSox = 0 To Portal.Sockets
        Select Case Sockets(tmpSox).GOREsock_State
        Case soxDisconnected
        Case soxClosing ' No need to close a closing Socket
        Case soxBound ' Same as soxListening
            If apiWSAAsyncSelect(Sockets(tmpSox).Socket, WindowhWnd, 0&, 0&) = GOREsock_ERROR Then
                Call GOREsock_RaiseState(tmpSox, soxDisconnected)
                Call GOREsock_Close(tmpSox)
                Let GOREsock_ShutDown = GOREsock_ERROR
            Else
                If apiCloseSocket(Sockets(tmpSox).Socket) = GOREsock_ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                    Call GOREsock_RaiseState(tmpSox, soxDisconnected)
                    Call GOREsock_Close(tmpSox)
                    Let GOREsock_ShutDown = GOREsock_ERROR
                Else
                    Call GOREsock_RaiseState(tmpSox, soxDisconnected)
                    Call GOREsock_Close(tmpSox)
                End If
            End If
        Case soxListening ' Same as soxBound
            If apiWSAAsyncSelect(Sockets(tmpSox).Socket, WindowhWnd, 0&, 0&) = GOREsock_ERROR Then
                Call GOREsock_RaiseState(tmpSox, soxDisconnected)
                Call GOREsock_Close(tmpSox)
                Let GOREsock_ShutDown = GOREsock_ERROR
            Else
                If apiCloseSocket(Sockets(tmpSox).Socket) = GOREsock_ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                    Call GOREsock_RaiseState(tmpSox, soxDisconnected)
                    Call GOREsock_Close(tmpSox)
                    Let GOREsock_ShutDown = GOREsock_ERROR
                Else
                    Call GOREsock_RaiseState(tmpSox, soxDisconnected)
                    Call GOREsock_Close(tmpSox)
                End If
            End If
        Case Else
            If apiShutDown(Sockets(tmpSox).Socket, SD_SEND) = GOREsock_ERROR Then
                Let GOREsock_ShutDown = soxERROR
            Else
                Call GOREsock_RaiseState(tmpSox, soxClosing)
            End If
        End Select
    Next tmpSox
    DoEvents ' There could be an incomming FD_CLOSE
    For tmpSox = 0 To Portal.Sockets
        If Not Sockets(tmpSox).GOREsock_State = soxDisconnected Then Exit For
    Next tmpSox
    If Not tmpSox = Portal.Sockets + 1 Then
        Let GOREsock_ShutDown = soxERROR ' This could also be set by all the soxClosing sockets
    Else
        DoEvents
        Let Portal.Sockets = -1
        Erase Sockets
        Erase Recv
        Erase Send
    End If

End Function

Private Function GOREsock_Socket2Sox(inSocket As Long) As Long ' Returns the Sockets() Array address of a WinSock Socket

    For GOREsock_Socket2Sox = 0 To Portal.Sockets
        If Sockets(GOREsock_Socket2Sox).Socket = inSocket Then Exit For
    Next GOREsock_Socket2Sox
    If GOREsock_Socket2Sox = Portal.Sockets + 1 Then Let GOREsock_Socket2Sox = INVALID_SOCKET

End Function

Public Function GOREsock_SocketHandle(inSox As Long) As Long

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let GOREsock_SocketHandle = soxERROR
    Else
        Let GOREsock_SocketHandle = Sockets(inSox).Socket
    End If

End Function

Public Function GOREsock_State(inSox As Long) As enmSoxState

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let GOREsock_State = soxERROR
    Else
        Let GOREsock_State = Sockets(inSox).GOREsock_State
    End If

End Function

Private Function GOREsock_StringFromPointer(ByVal lPointer As Long) As String

    Let GOREsock_StringFromPointer = Space$(apiLStrLen(ByVal lPointer))
    Call apiLstrCpy(ByVal GOREsock_StringFromPointer, ByVal lPointer)

End Function

Public Function GOREsock_uMsg(inSox As Long) As Long  ' This is the closest I can probably get to defining the type of Sox

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Let GOREsock_uMsg = soxERROR
    Else
        Let GOREsock_uMsg = Sockets(inSox).GOREsock_uMsg
    End If

End Function

Public Sub GOREsock_UnHook() ' Once the Control is UnHooked, we will not be able to intercept messages from WinSock API and process them according to our needs!
    
    Call apiSetWindowLong(WindowhWnd, GWL_GOREsock_WndProc, Portal.GOREsock_WndProc)
    Let Portal.GOREsock_WndProc = 0

End Sub

Sub GOREsock_Initialize(ByVal hWnd As Long)

    WindowhWnd = hWnd

    If apiWSAStartup(&H101, WSAData) = GOREsock_ERROR Then
        Call MsgBox("WinSock failed to initialize properly - Error#: " & Err.LastDllError, vbApplicationModal + vbCritical, "Critical Error")  'Creates an 'application instance' and memory space in the WinSock DLL (MUST be cleaned up later)
    Else
        Let Portal.GOREsock_WndProc = apiSetWindowLong(WindowhWnd, GWL_GOREsock_WndProc, AddressOf GOREsock_WndProc)
        Let Portal.Sockets = -1 ' GOREsock_Initialize our socket count ... NB - WE HAVE NONE, used wherever we Redim the Sockets Array
    End If
    
    Loaded = 1
    
End Sub

Public Sub GOREsock_Terminate()

    'Correctly replaces/reattaches the origional WindowProc procedure to our 'hidden' handle (Basically what the GOREsock_UnHook command does!)
    Call apiSetWindowLong(WindowhWnd, GWL_GOREsock_WndProc, Portal.GOREsock_WndProc)
    apiWSACleanup

    Loaded = 0

End Sub

Private Function GOREsock_WinSockError(ByVal lParam As Long) As Integer 'WSAGETSELECTERROR

    Let GOREsock_WinSockError = (lParam And &HFFFF0000) \ &H10000

End Function

Private Function GOREsock_WinSockEvent(ByVal lParam As Long) As Integer 'WSAGETSELECTEVENT

    If (lParam And &HFFFF&) > &H7FFF Then
        Let GOREsock_WinSockEvent = (lParam And &HFFFF&) - &H10000
    Else
        Let GOREsock_WinSockEvent = lParam And &HFFFF&
    End If

End Function

Public Function GOREsock_WndProc(ByVal hWnd As Long, ByVal GOREsock_uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case GOREsock_uMsg
    Case soxSERVER
        Select Case GOREsock_WinSockEvent(lParam)
        Case FD_ACCEPT
            If GOREsock_WinSockError(lParam) = 0 Then Call GOREsock_Accept(wParam)
        Case FD_CLOSE
            Select Case GOREsock_WinSockError(lParam)
            Case 0
                Select Case Sockets(GOREsock_Socket2Sox(wParam)).GOREsock_State
                Case soxClosing: Call GOREsock_Closed(GOREsock_Socket2Sox(wParam))
                Case Else
                    Call GOREsock_Shut(GOREsock_Socket2Sox(wParam))
                    Call GOREsock_Closed(GOREsock_Socket2Sox(wParam))
                End Select
            Case Else
                Call GOREsock_Closed(GOREsock_Socket2Sox(wParam))
            End Select
        Case FD_READ
            Call GOREsock_GetData(GOREsock_Socket2Sox(wParam))
        Case FD_WRITE
            If GOREsock_WinSockError(lParam) = 0 Then
                Select Case Sockets(GOREsock_Socket2Sox(wParam)).GOREsock_State
                Case soxConnecting
                    Call GOREsock_RaiseState(GOREsock_Socket2Sox(wParam), soxIdle)
                Case soxIdle
                    Call GOREsock_SendBuffer(GOREsock_Socket2Sox(wParam))
                Case soxClosing
                    Call GOREsock_Closed(GOREsock_Socket2Sox(wParam))
                End Select
            End If
        End Select
    Case soxCLIENT
        Select Case GOREsock_WinSockEvent(lParam)
        Case FD_CLOSE
            Select Case GOREsock_WinSockError(lParam)
            Case 0
                Select Case Sockets(GOREsock_Socket2Sox(wParam)).GOREsock_State
                Case soxClosing: Call GOREsock_Closed(GOREsock_Socket2Sox(wParam))
                Case Else
                    Call GOREsock_Shut(GOREsock_Socket2Sox(wParam))
                    Call GOREsock_Closed(GOREsock_Socket2Sox(wParam))
                End Select
            Case Else
                Call GOREsock_Closed(GOREsock_Socket2Sox(wParam))
            End Select
        Case FD_READ
            Call GOREsock_GetData(GOREsock_Socket2Sox(wParam))
        Case FD_WRITE
            If GOREsock_WinSockError(lParam) = 0 Then
                Select Case Sockets(GOREsock_Socket2Sox(wParam)).GOREsock_State
                Case soxConnecting
                    Call GOREsock_RaiseState(GOREsock_Socket2Sox(wParam), soxIdle)
                Case soxIdle
                    Call GOREsock_SendBuffer(GOREsock_Socket2Sox(wParam))
                Case soxClosing
                    Call GOREsock_Closed(GOREsock_Socket2Sox(wParam))
                End Select
            End If
        End Select
    Case Else
        Let GOREsock_WndProc = apiCallWindowProc(Portal.GOREsock_WndProc, hWnd, GOREsock_uMsg, wParam, lParam)
    End Select

End Function



