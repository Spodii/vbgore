VERSION 5.00
Begin VB.UserControl Socket 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   660
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   FontTransparent =   0   'False
   Picture         =   "Sox.ctx":0000
   ScaleHeight     =   44
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   44
   ToolboxBitmap   =   "Sox.ctx":16F2
End
Attribute VB_Name = "Socket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'GORESOCK is a modification of SOX - credits go to the creator of SOX

Private BufferSize As Long
Private tmpGetBuffer() As Byte
Private tmpSendSize(0 To 3) As Byte

Public Enum enmSoxState
    soxDisconnected = 0&
    soxListening = 1&
    soxConnecting = 2&
    soxIdle = 3&
    soxSend = 4&
    soxRecv = 5&
    soxClosing = 6&
    soxBound = 10&
    soxERROR = -1&
End Enum
#If False Then
Private soxDisconnected, soxListening, soxConnecting, soxIdle, soxSend, soxRecv, soxClosing, soxBound, soxERROR
#End If

Public Enum enmSoxOptions
    soxSO_BROADCAST = &H20&
    soxSO_DEBUG = &H1&
    soxSO_DONTROUTE = &H10&
    soxSO_KEEPALIVE = &H8&
    soxSO_LINGER = &H80&
    soxSO_OOBINLINE = &H100&
    soxSO_RCVBUF = &H1002&
    soxSO_REUSEADDR = &H4&
    soxSO_SNDBUF = &H1001&
    soxSO_TCP_NODELAY = Not &H1&
    soxSO_USELOOPBACK = &H40&
    soxSO_ACCEPTCONN = &H2&
    soxSO_ERROR = &H1007&
    soxSO_TYPE = &H1008&
End Enum
#If False Then
Private soxSO_BROADCAST, soxSO_DEBUG, soxSO_DONTROUTE, soxSO_KEEPALIVE, soxSO_LINGER, soxSO_OOBINLINE, soxSO_RCVBUF, soxSO_REUSEADDR, soxSO_SNDBUF, _
        soxSO_TCP_NODELAY, soxSO_USELOOPBACK, soxSO_ACCEPTCONN, soxSO_ERROR, soxSO_TYPE
#End If

Public Enum enmSoxTypes
    soxSERVER = 4026&
    soxCLIENT = 4027&
End Enum
#If False Then
Private soxSERVER, soxCLIENT
#End If

Private Type typPortal
    WndProc As Long
    Sockets As Long
End Type

Private Const SOCKADDR_SIZE As Long = 16&
Private Type typSocketAddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(7) As Byte
End Type

Private Type typSocket
    Socket As Long
    SocketAddr As typSocketAddr
    State As enmSoxState
    uMsg As Long
End Type

Private Type typBuffer
    Size As Long
    Pos As Long
    Buffer() As Byte
End Type

Private Const WSADESCRIPTION_LEN As Long = 255&
Private Const WSASYS_STATUS_LEN As Long = 127&
Private Type typWSAData
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADESCRIPTION_LEN) As Byte
    szSystemStatus(0 To WSASYS_STATUS_LEN) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Const WSABASEERR As Long = 10000
Private Const INVALID_SOCKET As Long = -1&
Private Const SOCKET_ERROR As Long = -1&
Private Const INADDR_NONE As Long = &HFFFFFFFF
Private Const AF_INET As Long = 2
Private Const SOCK_STREAM As Long = 1
Private Const SOL_SOCKET As Long = &HFFFF&
Private Const FD_READ As Long = &H1
Private Const FD_WRITE As Long = &H2
Private Const FD_ACCEPT As Long = &H8
Private Const FD_CONNECT As Long = &H10
Private Const FD_CLOSE As Long = &H20
Private Const SD_SEND As Long = &H1
Private Const IPPROTO_TCP As Long = 6

Private Const GWL_WNDPROC As Long = (-4)

Private Const WSAEINTR = (WSABASEERR + 4) 'Interrupted function call
Private Const WSAEBADF = (WSABASEERR + 9)
Private Const WSAEACCES = (WSABASEERR + 13) 'Permission Denied
Private Const WSAEFAULT = (WSABASEERR + 14) 'Bad address
Private Const WSAEINVAL = (WSABASEERR + 22) 'Invalid argument
Private Const WSAEMFILE = (WSABASEERR + 24) 'Too many open files
Private Const WSAEWOULDBLOCK = (WSABASEERR + 35) 'Resource temporarily unavailable
Private Const WSAEINPROGRESS = (WSABASEERR + 36) 'Operation now in progress
Private Const WSAEALREADY = (WSABASEERR + 37) 'Operation already in progress
Private Const WSAENOTSOCK = (WSABASEERR + 38) 'Socket operation on non-socket
Private Const WSAEDESTADDRREQ = (WSABASEERR + 39) 'Destination address required
Private Const WSAEMSGSIZE = (WSABASEERR + 40) 'Message too long
Private Const WSAEPROTOTYPE = (WSABASEERR + 41) 'Protocol wrong type for socket
Private Const WSAENOPROTOOPT = (WSABASEERR + 42) 'Bad protocol option
Private Const WSAEPROTONOSUPPORT = (WSABASEERR + 43) 'Protocol not supported
Private Const WSAESOCKTNOSUPPORT = (WSABASEERR + 44) 'Socket type not supported
Private Const WSAEOPNOTSUPP = (WSABASEERR + 45) 'Operation not supported
Private Const WSAEPFNOSUPPORT = (WSABASEERR + 46) 'Protocol family not supported
Private Const WSAEAFNOSUPPORT = (WSABASEERR + 47) 'Address family not supported by protocol family
Private Const WSAEADDRINUSE = (WSABASEERR + 48) 'Address already in use
Private Const WSAEADDRNOTAVAIL = (WSABASEERR + 49) 'Cannot assign requested address
Private Const WSAENETDOWN = (WSABASEERR + 50) 'Network is down
Private Const WSAENETUNREACH = (WSABASEERR + 51) 'Network is unreachable
Private Const WSAENETRESET = (WSABASEERR + 52) 'Network dropped connection on reset
Private Const WSAECONNABORTED = (WSABASEERR + 53) 'Software caused connection abort
Private Const WSAECONNRESET = (WSABASEERR + 54) 'Connection reset by peer
Private Const WSAENOBUFS = (WSABASEERR + 55) 'No buffer space available
Private Const WSAEISCONN = (WSABASEERR + 56) 'Socket is already connected
Private Const WSAENOTCONN = (WSABASEERR + 57) 'Socket is not connected
Private Const WSAESHUTDOWN = (WSABASEERR + 58) 'Cannot send after socket shutdown
Private Const WSAETOOMANYREFS = (WSABASEERR + 59) 'Too many references: can't splice (UnConfirmed Description)
Private Const WSAETIMEDOUT = (WSABASEERR + 60) 'Connection timed out
Private Const WSAECONNREFUSED = (WSABASEERR + 61) 'Connection refused
Private Const WSAELOOP = (WSABASEERR + 62) 'Too many levels of symbolic links (UnConfirmed Description)
Private Const WSAENAMETOOLONG = (WSABASEERR + 63) 'File name too long (UnConfirmed Description)
Private Const WSAEHOSTDOWN = (WSABASEERR + 64) 'Host is down
Private Const WSAEHOSTUNREACH = (WSABASEERR + 65) 'No route to host
Private Const WSAENOTEMPTY = (WSABASEERR + 66) 'Directory not empty (UnConfirmed Description)
Private Const WSAEPROCLIM = (WSABASEERR + 67) 'Too many processes
Private Const WSAEUSERS = (WSABASEERR + 68) 'Too many users (UnConfirmed Description)
Private Const WSAEDQUOT = (WSABASEERR + 69) 'Disk quota exceeded (UnConfirmed Description)
Private Const WSAESTALE = (WSABASEERR + 70) 'Stale NFS file handle (UnConfirmed Description)
Private Const WSAEREMOTE = (WSABASEERR + 71) 'Too many levels of remote in path (UnConfirmed Description)
Private Const WSASYSNOTREADY = (WSABASEERR + 91) 'Network subsystem is unavailable
Private Const WSAVERNOTSUPPORTED = (WSABASEERR + 92) 'WINSOCK.DLL version out of range
Private Const WSANOTINITIALISED = (WSABASEERR + 93) 'Successful WSAStartup not yet performed
Private Const WSAEDISCON1 = (WSABASEERR + 94) 'Graceful shutdown in progress
Private Const WSAEDISCON2 = (WSABASEERR + 101) 'Graceful shutdown in progress
Private Const WSAENOMORE = (WSABASEERR + 102)
Private Const WSAECANCELLED = (WSABASEERR + 103)
Private Const WSAEINVALIDPROCTABLE = (WSABASEERR + 104)
Private Const WSAEINVALIDPROVIDER = (WSABASEERR + 105)
Private Const WSAEPROVIDERFAILEDINIT = (WSABASEERR + 106)
Private Const WSASYSCALLFAILURE = (WSABASEERR + 107) '(OS Dependent) System call failure
Private Const WSASERVICE_NOT_FOUND = (WSABASEERR + 108)
Private Const WSATYPE_NOT_FOUND = (WSABASEERR + 109) 'Class type not found
Private Const WSA_E_NO_MORE = (WSABASEERR + 110)
Private Const WSA_E_CANCELLED = (WSABASEERR + 111)
Private Const WSAEREFUSED = (WSABASEERR + 112)
Private Const WSAHOST_NOT_FOUND = (WSABASEERR + 1001) 'Host not found
Private Const WSATRY_AGAIN = (WSABASEERR + 1002) 'Non-authoritative host not found
Private Const WSANO_RECOVERY = (WSABASEERR + 1003) 'This is a non-recoverable error
Private Const WSANO_DATA = (WSABASEERR + 1004) 'Valid name, no data record of requested type

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
Private Declare Function apiHToNS Lib "WS2_32" Alias "htons" (ByVal hostshort As Long) As Integer
Private Declare Function apiNToHS Lib "WS2_32" Alias "ntohs" (ByVal netshort As Long) As Integer
Private Declare Function apiIPToNL Lib "WS2_32" Alias "inet_addr" (ByVal cp As String) As Long
Private Declare Function apiNLToIP Lib "WS2_32" Alias "inet_ntoa" (ByVal inn As Long) As Long
Private Declare Function apiGetHostName Lib "WS2_32" Alias "gethostname" (ByVal name As String, ByVal namelen As Long) As Long
Private Declare Function apiShutDown Lib "WS2_32" Alias "shutdown" (ByVal S As Long, ByVal how As Long) As Long
Private Declare Function apiCallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function apiSetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function apiLStrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function apiLstrCpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Sub apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Event OnClose(inSox As Long)
Public Event OnConnect(inSox As Long) 'Notification of connection established to a Server
Public Event OnConnection(inSox As Long) 'Notification of a new connection (From a Listening Port)
Public Event OnDataArrival(inSox As Long, inData() As Byte) ' Unlike the origional WinSock OCX, a byte Array is used as the Data type instead of a Variant as this is a faster way of getting data to you directly
' This is the Old WinSock OCX Error Event ... Too complicated and unnecessary ... who used all this crap anyway ???
' Public Event Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Public Event OnError(inSox As Long, inError As Long, inDescription As String, inSource As String, inSnipet As String)
Public Event OnRecvProgress(inSox As Long, bytesRecv As Long, bytesRemaining As Long)
Public Event OnSendComplete(inSox As Long)
Public Event OnSendProgress(inSox As Long, bytesSent As Long, bytesRemaining As Long) ' Currently unused
Public Event OnState(inSox As Long, inState As enmSoxState) ' Notification of a new State Change!
Public Event OnStatus(inSox As Long, inSource As String, inStatus As String) ' Really useful to track general Sox State information showing where we are currently in code, I use it instead of Debug.Print because much of Sox testing must be done from the EXE not IDE!!! And there ain't no Debug.Print in the EXE :)))


Private WSAData As typWSAData
Private Portal As typPortal
Private Sockets() As typSocket
Private Recv() As typBuffer
Private Send() As typBuffer

'Variables stored to prevent recreation upon sub calls
'Easier to use a few extra bytes of RAM then to recreate it for every packet
Private a As Long
Private b As Long
Private c As Long
Private d As Byte

'***** ENCRYPTIONS *****

'Encryptions module for GORESOCK
'Simplified and trimmed version from the vbGORE encryptiosn module
'Only XOR and RC4 were left in since these are fast and have no size inflation!
'Only byte array support was left in since we don't send/rec in anything else - RAW POWER!!! >:D

'***** RC4 *****
Private m_sBoxRC4(0 To 255) As Integer
Private m_sBox(0 To 255) As Integer
Private m_KeyS As String

'***** SIMPLE XOR *****
Private m_XORKey() As Byte
Private m_XORKeyLen As Long
Private m_XORKeyValue As String


Public Enum EncryptType
    eNone = 0
    eXOR = 1
    eRC4 = 2
End Enum
#If False Then
Private eNone, eXOR, eRC4
#End If
Public EncryptMethod As EncryptType

Public Sub SetBufferSize(ByVal Size As Long)

    BufferSize = Size
    ReDim tmpGetBuffer(0 To BufferSize - 1)

End Sub

Public Sub SetEncryption(Method As EncryptType, Key As String)

    Select Case Method
        Case eNone
            EncryptMethod = eNone
        Case eXOR
            EncryptMethod = eXOR
            Encryption_XOR_SetKey Key
        Case eRC4
            EncryptMethod = eRC4
            Encryption_RC4_SetKey Key
    End Select

End Sub

Private Sub Encryption_RC4(ByteArray() As Byte, Optional Key As String)

    If (Len(Key) > 0) Then Encryption_RC4_SetKey Key
    
    Call apiCopyMemory(m_sBox(0), m_sBoxRC4(0), 512)

    a = 0
    b = 0
    For c = 0 To UBound(ByteArray)
        a = (a + 1) Mod 256
        b = (b + m_sBox(a)) Mod 256
        d = m_sBox(a)
        m_sBox(a) = m_sBox(b)
        m_sBox(b) = d
        ByteArray(c) = ByteArray(c) Xor (m_sBox((m_sBox(a) + m_sBox(b)) Mod 256))
    Next

End Sub

Private Sub DoEncryption(bArray() As Byte)
 
    Select Case EncryptMethod
        Case eXOR
            Encryption_XOR bArray()
        Case eRC4
            Encryption_RC4 bArray()
    End Select
            
End Sub

Private Sub Encryption_RC4_SetKey(New_Value As String)
Dim Key() As Byte

    If (m_KeyS = New_Value) Then Exit Sub
    m_KeyS = New_Value

    Key() = StrConv(m_KeyS, vbFromUnicode)
    c = Len(m_KeyS)

    For a = 0 To 255
        m_sBoxRC4(a) = a
    Next a
    
    b = 0
    For a = 0 To 255
        b = (b + m_sBoxRC4(a) + Key(a Mod c)) Mod 256
        d = m_sBoxRC4(a)
        m_sBoxRC4(a) = m_sBoxRC4(b)
        m_sBoxRC4(b) = d
    Next

End Sub

Private Sub Encryption_XOR(ByteArray() As Byte, Optional Key As String)

    If (Len(Key) > 0) Then Encryption_XOR_SetKey Key

    For a = 0 To UBound(ByteArray)
        ByteArray(a) = ByteArray(a) Xor m_XORKey(a Mod m_XORKeyLen)
    Next

End Sub

Private Sub Encryption_XOR_SetKey(New_Value As String)

    If (m_XORKeyValue = New_Value) Then Exit Sub

    m_XORKeyValue = New_Value
    m_XORKeyLen = Len(New_Value)
    m_XORKey() = StrConv(m_XORKeyValue, vbFromUnicode)

End Sub

Private Function Accept(inSocket As Long) As Long 'Returns: New Sox Number -- inSocket is the listening WinSocket, not Sox number ...

Const Procedure As String = "Accept"
Dim tmpSocket As Long
Dim tmpSocketAddr As typSocketAddr 'This stores the details of our new socket/client, including the client IP address

    Let tmpSocket = apiAccept(inSocket, tmpSocketAddr, SOCKADDR_SIZE) 'Accept API returns a valid, random, unused socket for us to use for the new client
    If tmpSocket = INVALID_SOCKET Then 'Accept API may not give us a valid socket eg. when all sockets are full, you may have to add additional error trapping if you believe you will use over 32,767 sockets
        'Since a socket was not commited for the new Connection ... we don't have to close it (Since the socket was never even created)
        Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "On Accept")
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
        RaiseEvent OnStatus(Accept, Procedure, "New Connection")
        RaiseEvent OnConnection(Accept)
    End If

End Function

Public Function Address(inSox As Long) As String ' Returns the address used by a Socket (Either Local or Remote)

Const Procedure As String = "Address"

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(soxERROR, 0, Procedure, "Invalid Sox: " & inSox)
        Let Address = soxERROR
    Else
        Let Address = StringFromPointer(apiNLToIP(Sockets(inSox).SocketAddr.sin_addr))
    End If

End Function

Public Function Bind(LocalIP As String, LocalPort As Integer) As Long

Const Procedure As String = "Bind"
Dim tmpSocket As Long
Dim tmpSocketAddr As typSocketAddr

    If LocalPort = 0 Or LocalIP = vbNullString Then
        Call RaiseError(soxERROR, 0, Procedure, "Invalid Port or IP address")
        Let Bind = soxERROR
    Else
        Let tmpSocketAddr.sin_family = AF_INET
        Let tmpSocketAddr.sin_port = apiHToNS(LocalPort)
        If tmpSocketAddr.sin_port = INVALID_SOCKET Then
            Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "Invalid Port: " & tmpSocketAddr.sin_port)
            Let Bind = INVALID_SOCKET
        Else
            Let tmpSocketAddr.sin_addr = apiIPToNL(LocalIP) 'If this is Zero, it will assign 0.0.0.0 !!!
            If tmpSocketAddr.sin_addr = INADDR_NONE Then 'If 255.255.255.255 is returned ... we have a problem ... I think :)
                Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "Invalid NL Address: " & tmpSocketAddr.sin_addr) ' NL = Network Long
                Let Bind = INVALID_SOCKET
            Else
                Let tmpSocket = apiSocket(AF_INET, SOCK_STREAM, IPPROTO_TCP) 'This is where you specify what type of protocol to use and what type of Streaming to use, returns a new socket number 4 us (NB - From here, if any further steps fail after this one succeeds, we must close the socket)
                If tmpSocket = INVALID_SOCKET Then
                    Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "On Creation")
                    Let Bind = INVALID_SOCKET
                Else
                    If apiBind(tmpSocket, tmpSocketAddr, SOCKADDR_SIZE) = SOCKET_ERROR Then 'Socket Number, Socket Address space / Name, Name Length ...
                        Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Bind")
                        If apiCloseSocket(tmpSocket) = SOCKET_ERROR Then Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Close") ' We MUST start closing the Socket handle from this point (Unless we store the number and force WinSock to use it later ... Nah ... too much codeing :)))
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
                        RaiseEvent OnStatus(Bind, Procedure, "Bound")
                    End If
                End If
            End If
        End If
    End If

End Function

' This function origionally (in the old OCX) used the apiRecv call with the MSG_PEEK flag!
' Microsoft does not recommend this procedure as it slows things down and is very potentially
' a misleading indicator! Therefore the next 2 functions are similar to the On~~~Progress Events
' I do not recommend their use, cause all you will are the Events for progress!

Public Function BytesReceived(inSox As Long) As Long

Const Procedure As String = "BytesReceived"

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(soxERROR, 0, Procedure, "Invalid Sox: " & inSox)
        Let BytesReceived = soxERROR
    Else
        ' Also check the RaiseEvent OnRecvProgress in GetData
        If Recv(inSox).Size > 2 Then Let BytesReceived = Recv(inSox).Size - 3
    End If

End Function

Public Function bytesSent(inSox As Long) As Long

Const Procedure As String = "BytesSent"

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(soxERROR, 0, Procedure, "Invalid Sox: " & inSox)
        Let bytesSent = soxERROR
    Else
        ' Also check the RaiseEvent OnSendProgress in SendBuffer
        If Send(inSox).Size > 2 Then Let bytesSent = Send(inSox).Pos - 3
    End If

End Function

Private Sub Closed(inSox As Long) ' This Socket has successfully closed ... free resources (No need to check if it exists, cause we call this internally)

Const Procedure As String = "Closed"

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(INVALID_SOCKET, 0, Procedure, "Critical Error, Detected an Invalid Sox: " & inSox) ' This should NEVER happen, why? Because WinSock API closed a valid Socket that we didn't even know about!
    Else
        If apiWSAAsyncSelect(Sockets(inSox).Socket, UserControl.hWnd, 0&, 0&) = SOCKET_ERROR Then       'FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
            Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiWSAAsyncSelect")
            Call RaiseState(inSox, soxDisconnected) ' Force disconnected status, dunno what the implications are!
        Else
            If apiCloseSocket(Sockets(inSox).Socket) = SOCKET_ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiCloseSocket")
                Call RaiseState(inSox, soxDisconnected) ' Force disconnected status, dunno what the implications are!
            Else
                Call RaiseState(inSox, soxDisconnected)
                RaiseEvent OnStatus(inSox, Procedure, "Successfully Closed")
                RaiseEvent OnClose(inSox)
            End If
        End If
    End If

End Sub

Public Function Connect(RemoteHost As String, RemotePort As Integer) As Long 'Returns the new Sox Number / SOCKET_ERROR On Error

Const Procedure As String = "Connect"
Dim tmpSocket As Long
Dim tmpSocketAddr As typSocketAddr

    Let tmpSocketAddr.sin_family = AF_INET
    Let tmpSocketAddr.sin_port = apiHToNS(RemotePort) ' apiHToNS(RemotePort)
    If tmpSocketAddr.sin_port = INVALID_SOCKET Then
        Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "Port: " & tmpSocketAddr.sin_port)
        Let Connect = INVALID_SOCKET
    Else
        Let tmpSocketAddr.sin_addr = apiIPToNL(RemoteHost) 'If this is Zero, it will assign 0.0.0.0 !!!
        If tmpSocketAddr.sin_addr = INADDR_NONE Then 'If 255.255.255.255 is returned ... we have a problem ... I think :)
            Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "NL Address: " & tmpSocketAddr.sin_addr) ' NL = Network Long
            Let Connect = INVALID_SOCKET
        Else
            Let tmpSocket = apiSocket(AF_INET, SOCK_STREAM, IPPROTO_TCP) 'This is where you specify what type of protocol to use and what type of Streaming to use, returns a new socket number 4 us (NB - From here, if any further steps fail after this one succeeds, we must close the socket)
            If tmpSocket = INVALID_SOCKET Then
                Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "On Socket Creation")
                Let Connect = INVALID_SOCKET
            Else
                If apiConnect(tmpSocket, tmpSocketAddr, SOCKADDR_SIZE) = SOCKET_ERROR Then
                    Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Connect")
                    If apiCloseSocket(tmpSocket) = SOCKET_ERROR Then Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Close") ' We MUST start closing the Socket handle from this point (Unless we store the number and force WinSock to use it later ... Nah ... too much codeing :)))
                    Let Connect = SOCKET_ERROR
                Else
                    If apiWSAAsyncSelect(tmpSocket, UserControl.hWnd, ByVal soxCLIENT, ByVal FD_ACCEPT Or FD_CLOSE Or FD_CONNECT Or FD_READ Or FD_WRITE) = SOCKET_ERROR Then ' Reassign this Socket to Send and Receive on the DATA channel
                        Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On AsyncSelect")
                        If apiCloseSocket(tmpSocket) = SOCKET_ERROR Then Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Close")
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
                        RaiseEvent OnStatus(Connect, Procedure, "Connecting")
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
    DoEncryption tmpBuffer
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

Const Procedure As String = "GetData"
Dim tmpRecvSize As Long

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(soxERROR, 0, Procedure, "FATAL ERROR: Invalid Sox: " & inSox) ' This should NEVER happen (never has though!) because it indicates that Sox is sending itself an invalid Sox :(((
    Else
        Select Case Sockets(inSox).State
        Case soxRecv ' If another Receive is being processed ... this will cause HAVOK with our data
        Case soxSend
        Case soxClosing
        Case soxIdle
            Call RaiseState(inSox, soxRecv) ' If this socket is closing ... we could cause HAVOK too
            ' First we will disable further notification of FD_READ, because if we extract data with the Recv function, WinSock API posts ANOTHER FD_READ notification to say there's more ...
            ' This is a valid (dare I say recommended) procedure according to WinSock API documentation on MSDN
            If apiWSAAsyncSelect(Sockets(inSox).Socket, UserControl.hWnd, 0&, 0&) = SOCKET_ERROR Then  ' Reassign this Socket to Send and Receive on the DATA channel
                Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiWSAAsyncSelect")
            Else
                Let tmpRecvSize = apiRecv(Sockets(inSox).Socket, tmpGetBuffer(0), BufferSize, 0)
                DoEvents
                Select Case tmpRecvSize
                Case SOCKET_ERROR ' Houston, we have a problem :)))
                    Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiRecv") ' If we send the exact same amount of data as the buffer, we get error 10035 = Resource temporarily unavailable ... not 0
                Case 0 ' The Socket was Gracefully closed (Never seen this happen!!! Maybe it happens in some older/newer version of WinSock API???)
                    RaiseEvent OnStatus(inSox, Procedure, "Gracefully Closed")
                    Call RaiseState(inSox, soxDisconnected)
                    RaiseEvent OnClose(inSox)
                Case Else
                    ReDim Preserve Recv(inSox).Buffer(Recv(inSox).Size + tmpRecvSize)
                    Call apiCopyMemory(Recv(inSox).Buffer(Recv(inSox).Size + 1), tmpGetBuffer(0), tmpRecvSize)
                    Let Recv(inSox).Size = Recv(inSox).Size + tmpRecvSize
                    Do While Recv(inSox).Size > 2 ' If for example we received many small 'packets' of data, this will loop until we have returned/extracted all of them!
                        DoEvents
                        If Recv(inSox).Size - 3 > RecvSize(inSox) Then
                            RaiseEvent OnRecvProgress(inSox, Recv(inSox).Size - 3, 0)
                            RaiseEvent OnDataArrival(inSox, ExtractRecv(inSox))
                        Else
                            RaiseEvent OnRecvProgress(inSox, Recv(inSox).Size - 3, (RecvSize(inSox) + 1) - (Recv(inSox).Size - 3))
                            Exit Do
                        End If
                    Loop
                End Select
                If Sockets(inSox).State = soxRecv Then Call RaiseState(inSox, soxIdle) ' If this socket is closing ... we could cause HAVOK too
                If apiWSAAsyncSelect(Sockets(inSox).Socket, UserControl.hWnd, ByVal Sockets(inSox).uMsg, ByVal FD_CLOSE Or FD_READ Or FD_WRITE) = SOCKET_ERROR Then   ' Reassign this Socket to Send and Receive on the DATA channel
                    Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiWSAAsyncSelect")
                End If
            End If
        Case Else
        End Select
    End If

End Sub

Public Function GetOption(inSox As Long, inOption As enmSoxOptions) As Long

Const Procedure As String = "GetOption"

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(soxERROR, 0, Procedure, "Invalid Sox: " & inSox)
        Let GetOption = soxERROR
    Else
        Select Case inOption
        Case soxSO_TCP_NODELAY
            If apiGetSockOpt(Sockets(inSox).Socket, IPPROTO_TCP, Not inOption, GetOption, 4) = SOCKET_ERROR Then
                Call RaiseError(inSox, Err.LastDllError, Procedure, "Option: " & inOption)
                Let GetOption = SOCKET_ERROR
            End If
        Case Else
            If apiGetSockOpt(Sockets(inSox).Socket, SOL_SOCKET, inOption, GetOption, 4) = SOCKET_ERROR Then
                Call RaiseError(inSox, Err.LastDllError, Procedure, "Option: " & inOption)
                Let GetOption = SOCKET_ERROR
            End If
        End Select
    End If

End Function

Public Sub Hook() ' WinSock is told to send it's messages to the Sox Control, but we need to intercept these messages!

Const Procedure As String = "Hook"

    If Portal.WndProc = 0 Then ' If it's already hooked to our WindowProc function, we could have problems, this will make sure we've UnHooked before
        Let Portal.WndProc = apiSetWindowLong(UserControl.hWnd, GWL_WNDPROC, AddressOf WindowProc)
        RaiseEvent OnStatus(soxERROR, Procedure, "Message hook enabled, Sox Control Hooked") ' soxERROR doesn't actually indicate an error here at all!!!
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

Const Procedure As String = "Listen"
Dim tmpSocket As Long
Dim tmpSocketAddr As typSocketAddr

    Let tmpSocketAddr.sin_family = AF_INET
    Let tmpSocketAddr.sin_port = apiHToNS(inPort)
    If tmpSocketAddr.sin_port = INVALID_SOCKET Then
        Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "Invalid Port: " & tmpSocketAddr.sin_port)
        Let Listen = INVALID_SOCKET
    Else
        Let tmpSocketAddr.sin_addr = apiIPToNL(inAddress) 'If this is Zero, it will assign 0.0.0.0 !!!
        If tmpSocketAddr.sin_addr = INADDR_NONE Then 'If 255.255.255.255 is returned ... we have a problem ... I think :)
            Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "Invalid NL Address: " & tmpSocketAddr.sin_addr) ' NL = Network Long
            Let Listen = INVALID_SOCKET
        Else
            Let tmpSocket = apiSocket(AF_INET, SOCK_STREAM, IPPROTO_TCP) 'This is where you specify what type of protocol to use and what type of Streaming to use, returns a new socket number 4 us (NB - From here, if any further steps fail after this one succeeds, we must close the socket)
            If tmpSocket = INVALID_SOCKET Then
                Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "On Creation")
                Let Listen = INVALID_SOCKET
            Else
                If apiBind(tmpSocket, tmpSocketAddr, SOCKADDR_SIZE) = SOCKET_ERROR Then 'Socket Number, Socket Address space / Name, Name Length ...
                    Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Bind")
                    If apiCloseSocket(tmpSocket) = SOCKET_ERROR Then Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Close") ' We MUST start closing the Socket handle from this point (Unless we store the number and force WinSock to use it later ... Nah ... too much codeing :)))
                    Let Listen = SOCKET_ERROR
                Else
                    If apiListen(ByVal tmpSocket, ByVal 5) = SOCKET_ERROR Then ' 5 = Maximum connections
                        Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Listen")
                        If apiCloseSocket(tmpSocket) = SOCKET_ERROR Then Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Close") ' We MUST start closing the Socket handle from this point (Unless we store the number and force WinSock to use it later ... Nah ... too much codeing :)))
                        Let Listen = SOCKET_ERROR
                    Else
                        If apiWSAAsyncSelect(tmpSocket, UserControl.hWnd, ByVal soxSERVER, ByVal FD_ACCEPT Or FD_CLOSE Or FD_CONNECT Or FD_READ Or FD_WRITE) = SOCKET_ERROR Then ' Reassign this Socket to Send and Receive on the DATA channel
                            Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On AsyncSelect")
                            If apiCloseSocket(tmpSocket) = SOCKET_ERROR Then Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Close")
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
                            RaiseEvent OnStatus(Listen, Procedure, "Listening")
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Public Function LocalHostName() As String ' The PC's Name eg. RonaldR (Needs a successful WSAStartup to function because it gets the PC name from WinSock)

Const Procedure As String = "LocalHostName"

    Let LocalHostName = Space(256) ' Create a 'buffer' for the API call
    If apiGetHostName(LocalHostName, 256) = SOCKET_ERROR Then
        Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On apiGetHostName")
        Let LocalHostName = vbNullString
    Else
        Let LocalHostName = Trim$(LocalHostName)
    End If

End Function

Private Sub Long2Byte2(inLong As Long, inByte() As Byte) ' Similar to the above, but places the bytes direcly into the given array

    Call apiCopyMemory(inByte(0), inLong, 4)

End Sub

Public Function Port(inSox As Long) As Long

Const Procedure As String = "Address"

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(soxERROR, 0, Procedure, "Invalid Sox: " & inSox)
        Let Port = soxERROR
    Else
        Let Port = apiNToHS(Sockets(inSox).SocketAddr.sin_port)
    End If

End Function

Private Sub RaiseError(inSox As Long, inCode As Long, inProcedure As String, inSnipet As String)   'Returns EXACTLY the same value as inError but raises the corresponding event if this is an error

    Select Case inCode
    Case WSABASEERR: RaiseEvent OnError(inSox, inCode, "General WinSock subsystem failure", inProcedure, inSnipet) 'Just sounds cool :)))
    Case WSAEINTR: RaiseEvent OnError(inSox, inCode, "Interrupted function call", inProcedure, inSnipet)
    Case WSAEBADF: RaiseEvent OnError(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
    Case WSAEACCES: RaiseEvent OnError(inSox, inCode, "Permission Denied", inProcedure, inSnipet)
    Case WSAEFAULT: RaiseEvent OnError(inSox, inCode, "Bad address", inProcedure, inSnipet)
    Case WSAEINVAL: RaiseEvent OnError(inSox, inCode, "Invalid argument", inProcedure, inSnipet)
    Case WSAEMFILE: RaiseEvent OnError(inSox, inCode, "Too many open files", inProcedure, inSnipet)
    Case WSAEWOULDBLOCK: RaiseEvent OnError(inSox, inCode, "Resource temporarily unavailable", inProcedure, inSnipet)
    Case WSAEINPROGRESS: RaiseEvent OnError(inSox, inCode, "Operation now in progress", inProcedure, inSnipet)
    Case WSAEALREADY: RaiseEvent OnError(inSox, inCode, "Operation already in progress", inProcedure, inSnipet)
    Case WSAENOTSOCK: RaiseEvent OnError(inSox, inCode, "Socket operation on non-socket", inProcedure, inSnipet)
    Case WSAEDESTADDRREQ: RaiseEvent OnError(inSox, inCode, "Destination address required", inProcedure, inSnipet)
    Case WSAEMSGSIZE: RaiseEvent OnError(inSox, inCode, "Message too long", inProcedure, inSnipet)
    Case WSAEPROTOTYPE: RaiseEvent OnError(inSox, inCode, "Protocol wrong type for socket", inProcedure, inSnipet)
    Case WSAENOPROTOOPT: RaiseEvent OnError(inSox, inCode, "Bad protocol option", inProcedure, inSnipet)
    Case WSAEPROTONOSUPPORT: RaiseEvent OnError(inSox, inCode, "Protocol not supported", inProcedure, inSnipet)
    Case WSAESOCKTNOSUPPORT: RaiseEvent OnError(inSox, inCode, "Socket type not supported", inProcedure, inSnipet)
    Case WSAEOPNOTSUPP: RaiseEvent OnError(inSox, inCode, "Operation not supported", inProcedure, inSnipet)
    Case WSAEPFNOSUPPORT: RaiseEvent OnError(inSox, inCode, "Protocol family not supported", inProcedure, inSnipet)
    Case WSAEAFNOSUPPORT: RaiseEvent OnError(inSox, inCode, "Address family not supported by protocol family", inProcedure, inSnipet)
    Case WSAEADDRINUSE: RaiseEvent OnError(inSox, inCode, "Address already in use", inProcedure, inSnipet)
    Case WSAEADDRNOTAVAIL: RaiseEvent OnError(inSox, inCode, "Cannot assign requested address", inProcedure, inSnipet)
    Case WSAENETDOWN: RaiseEvent OnError(inSox, inCode, "Network is down", inProcedure, inSnipet)
    Case WSAENETUNREACH: RaiseEvent OnError(inSox, inCode, "Network is unreachable", inProcedure, inSnipet)
    Case WSAENETRESET: RaiseEvent OnError(inSox, inCode, "Network dropped connection on reset", inProcedure, inSnipet)
    Case WSAECONNABORTED: RaiseEvent OnError(inSox, inCode, "Software caused connection abort", inProcedure, inSnipet)
    Case WSAECONNRESET: RaiseEvent OnError(inSox, inCode, "Connection reset by peer", inProcedure, inSnipet)
    Case WSAENOBUFS: RaiseEvent OnError(inSox, inCode, "No buffer space available", inProcedure, inSnipet)
    Case WSAEISCONN: RaiseEvent OnError(inSox, inCode, "Socket is already connected", inProcedure, inSnipet)
    Case WSAENOTCONN: RaiseEvent OnError(inSox, inCode, "Socket is not connected", inProcedure, inSnipet)
    Case WSAESHUTDOWN: RaiseEvent OnError(inSox, inCode, "Cannot send after socket shutdown", inProcedure, inSnipet)
    Case WSAETOOMANYREFS: RaiseEvent OnError(inSox, inCode, "Too many references: can't splice", inProcedure, inSnipet)  ' UnConfirmed Description
    Case WSAETIMEDOUT: RaiseEvent OnError(inSox, inCode, "Connection timed out", inProcedure, inSnipet)
    Case WSAECONNREFUSED: RaiseEvent OnError(inSox, inCode, "Connection refused", inProcedure, inSnipet)
    Case WSAELOOP: RaiseEvent OnError(inSox, inCode, "Too many levels of symbolic links", inProcedure, inSnipet)  ' UnConfirmed Description
    Case WSAENAMETOOLONG: RaiseEvent OnError(inSox, inCode, "File name too long", inProcedure, inSnipet)  ' UnConfirmed Description
    Case WSAEHOSTDOWN: RaiseEvent OnError(inSox, inCode, "Host is down", inProcedure, inSnipet)
    Case WSAEHOSTUNREACH: RaiseEvent OnError(inSox, inCode, "No route to host", inProcedure, inSnipet)
    Case WSAENOTEMPTY: RaiseEvent OnError(inSox, inCode, "Directory not empty", inProcedure, inSnipet)  ' UnConfirmed Description
    Case WSAEPROCLIM: RaiseEvent OnError(inSox, inCode, "Too many processes", inProcedure, inSnipet)
    Case WSAEUSERS: RaiseEvent OnError(inSox, inCode, "Too many users", inProcedure, inSnipet)  ' UnConfirmed Description
    Case WSAEDQUOT: RaiseEvent OnError(inSox, inCode, "Disk quota exceeded", inProcedure, inSnipet)  ' UnConfirmed Description
    Case WSAESTALE: RaiseEvent OnError(inSox, inCode, "Stale NFS file handle", inProcedure, inSnipet) ' UnConfirmed Description
    Case WSAEREMOTE: RaiseEvent OnError(inSox, inCode, "Too many levels of remote in path", inProcedure, inSnipet)  ' UnConfirmed Description
    Case WSASYSNOTREADY: RaiseEvent OnError(inSox, inCode, "Network subsystem is unavailable", inProcedure, inSnipet)
    Case WSAVERNOTSUPPORTED: RaiseEvent OnError(inSox, inCode, "WinSock.DLL version out of range", inProcedure, inSnipet)
    Case WSANOTINITIALISED: RaiseEvent OnError(inSox, inCode, "Successful WSAStartup not yet performed", inProcedure, inSnipet)
    Case WSAEDISCON1: RaiseEvent OnError(inSox, inCode, "Graceful shutdown in progress", inProcedure, inSnipet)
    Case WSAEDISCON2: RaiseEvent OnError(inSox, inCode, "Graceful shutdown in progress", inProcedure, inSnipet)
    Case WSAENOMORE: RaiseEvent OnError(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
    Case WSAECANCELLED: RaiseEvent OnError(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
    Case WSAEINVALIDPROCTABLE: RaiseEvent OnError(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
    Case WSAEINVALIDPROVIDER: RaiseEvent OnError(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
    Case WSAEPROVIDERFAILEDINIT: RaiseEvent OnError(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
    Case WSASYSCALLFAILURE: RaiseEvent OnError(inSox, inCode, "System call failure", inProcedure, inSnipet)
    Case WSASERVICE_NOT_FOUND: RaiseEvent OnError(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
    Case WSATYPE_NOT_FOUND: RaiseEvent OnError(inSox, inCode, "Class type not found", inProcedure, inSnipet)
    Case WSA_E_NO_MORE: RaiseEvent OnError(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
    Case WSA_E_CANCELLED: RaiseEvent OnError(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
    Case WSAEREFUSED: RaiseEvent OnError(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
    Case WSAHOST_NOT_FOUND: RaiseEvent OnError(inSox, inCode, "Host not found", inProcedure, inSnipet)
    Case WSATRY_AGAIN: RaiseEvent OnError(inSox, inCode, "Non-authoritative host not found", inProcedure, inSnipet)
    Case WSANO_RECOVERY: RaiseEvent OnError(inSox, inCode, "This is a non-recoverable error", inProcedure, inSnipet)
    Case WSANO_DATA: RaiseEvent OnError(inSox, inCode, "Valid name, no data record of requested type", inProcedure, inSnipet)
    Case Else: RaiseEvent OnError(inSox, inCode, "Unrecognized WinSock error", inProcedure, inSnipet)
    End Select

End Sub

Private Sub RaiseState(inSox As Long, inState As enmSoxState)

    Let Sockets(inSox).State = inState
    RaiseEvent OnState(inSox, inState)

End Sub

Private Function RecvSize(inSox As Long) As Long  ' Given 4 bytes, will directly copy them to a long! WARNING - To speed it up, I have no UBound checks, therefore you MUST send it 4 bytes

    Call apiCopyMemory(RecvSize, Recv(inSox).Buffer(0), 4)

End Function

Private Sub SendBuffer(inSox As Long)   'Data to be sent. For binary data, byte array should be used (for optimal performace, change inData to a byte array and only allow that datatype to be sent)

Const Procedure As String = "SendBuffer"

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(INVALID_SOCKET, 0, Procedure, "Invalid Sox: " & inSox)
    Else
        If Not Send(inSox).Size = soxERROR Then ' If there is data in the Buffer ...
            Select Case Sockets(inSox).State
            Case soxRecv ' Just terminate cause a Receive is being processed
            Case soxClosing ' Shouldn't / Cannot send while closing
                Call RaiseError(SOCKET_ERROR, 0, Procedure, "Invalid Sox State: " & Sockets(inSox).State)
            Case soxIdle
                Call RaiseState(inSox, soxSend)
                If apiWSAAsyncSelect(Sockets(inSox).Socket, UserControl.hWnd, 0&, 0&) = SOCKET_ERROR Then
                    Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiWSAAsyncSelect")
                End If
                Select Case Send(inSox).Size - Send(inSox).Pos
                Case Is < 0 ' We have no more data in buffer,
                Case Is = 0 ' What was this for again ???
                Case Is < BufferSize ' We have less data than our buffer size
                    If apiSend(Sockets(inSox).Socket, Send(inSox).Buffer(Send(inSox).Pos + 1), Send(inSox).Size - Send(inSox).Pos, 0) = SOCKET_ERROR Then
                        Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiSend")
                    Else
                        Let Send(inSox).Pos = Send(inSox).Size ' We have sent all the data in the Buffer
                    End If
                Case Else ' We have more data than our specified buffer size, or we have exactly BufferSize bytes
                    If apiSend(Sockets(inSox).Socket, Send(inSox).Buffer(Send(inSox).Pos + 1), BufferSize, 0) = SOCKET_ERROR Then
                        Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiSend")
                    Else
                        Let Send(inSox).Pos = Send(inSox).Pos + BufferSize ' If we sent exactly the last BufferSize byte, Pos and Size will be the same, indicating all data has been sent!
                    End If
                End Select
                Do While Send(inSox).Size > 2 ' Meaning we can extract SendSize from it which needs a minimum of 4 (0 to 3) so we test > 2
                    DoEvents
                    If Send(inSox).Pos - 3 > SendSize(inSox) Then ' Have we sent an entire SendData command ?
                        RaiseEvent OnSendProgress(inSox, Send(inSox).Pos - 3, 0)
                        Call ExtractSend(inSox)
                        RaiseEvent OnSendComplete(inSox)
                    Else
                        RaiseEvent OnSendProgress(inSox, Send(inSox).Pos - 3, (SendSize(inSox) + 1) - (Send(inSox).Pos - 3))
                        Exit Do
                    End If
                Loop
                If Sockets(inSox).State = soxSend Then Call RaiseState(inSox, soxIdle) ' If this socket is closing ... we could cause HAVOK ??? Right ???
                If apiWSAAsyncSelect(Sockets(inSox).Socket, UserControl.hWnd, ByVal Sockets(inSox).uMsg, ByVal FD_CLOSE Or FD_READ Or FD_WRITE) = SOCKET_ERROR Then
                    Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiWSAAsyncSelect")
                End If
            Case Else
                Call RaiseError(SOCKET_ERROR, 0, Procedure, "Invalid Sox State: " & Sockets(inSox).State)
            End Select
        End If
    End If

End Sub

'When a UNICODE string is passed in, it is converted to an ANSI string before being sent out on the network
'Hint - All data eventually get converted to Byte Arrays before being sent, therefore this is the most efficient data type, and if this is going to be the only data type, then you can improve send performance dramatically!

Public Function SendData(inSox As Long, inData() As Byte) As Long   'Data to be sent. For binary data, byte array should be used (for optimal performace, change inData to a byte array and only allow that datatype to be sent)

Const Procedure As String = "SendData"

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(soxERROR, 0, Procedure, "Invalid Sox: " & inSox)
        Let SendData = soxERROR
    Else
        If Not (Sockets(inSox).State = soxIdle Or Sockets(inSox).State = soxSend Or Sockets(inSox).State = soxRecv) Then ' If we have initiated a ShutDown, the state would change to Closing
            Call RaiseError(soxERROR, 0, Procedure, "Invalid Sox State: " & Sockets(inSox).State)
            Let SendData = soxERROR
        Else
            If UBound(inData) = soxERROR Then ' A value of -1 is returned from UBound if there was no data
                Call RaiseError(soxERROR, 0, Procedure, "Invalid Data. Possible cause: No Data to send eg. a blank string")
                Let SendData = soxERROR
            Else
                DoEncryption inData
                Call Long2Byte2(UBound(inData), tmpSendSize) ' I use UBound here instead of UBound + 1 to test the buffer on the other side EXACTLY !
                ReDim Preserve Send(inSox).Buffer(Send(inSox).Size + 4 + UBound(inData) + 1) As Byte ' 4 = Sized, UBound + 1 = DataLength
                Call apiCopyMemory(Send(inSox).Buffer(Send(inSox).Size + 1), tmpSendSize(0), 4)
                Let Send(inSox).Size = Send(inSox).Size + 4
                Call apiCopyMemory(Send(inSox).Buffer(Send(inSox).Size + 1), inData(0), UBound(inData) + 1)
                Let Send(inSox).Size = Send(inSox).Size + UBound(inData) + 1
                If apiWSAAsyncSelect(Sockets(inSox).Socket, UserControl.hWnd, ByVal Sockets(inSox).uMsg, ByVal FD_CLOSE Or FD_READ Or FD_WRITE) = SOCKET_ERROR Then
                    Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiWSAAsyncSelect")
                End If
            End If
        End If
    End If

End Function

Private Function SendSize(inSox As Long) As Long  ' Given 4 bytes, will directly copy them to a long! WARNING - To speed it up, I have no UBound checks, therefore you MUST send it 4 bytes

    Call apiCopyMemory(SendSize, Send(inSox).Buffer(0), 4)

End Function

Public Function SetOption(inSox As Long, inOption As enmSoxOptions, inValue As Long) As Long

Const Procedure As String = "SetOption"

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(soxERROR, 0, Procedure, "Invalid Sox: " & inSox)
        Let SetOption = soxERROR
    Else
        Select Case inOption
        Case soxSO_TCP_NODELAY
            If apiSetSockOpt(Sockets(inSox).Socket, IPPROTO_TCP, Not inOption, inValue, 4) = SOCKET_ERROR Then
                Call RaiseError(inSox, Err.LastDllError, Procedure, "Option: " & inOption & " & Value: " & inValue)
                Let SetOption = SOCKET_ERROR
            End If
        Case Else
            If apiSetSockOpt(Sockets(inSox).Socket, SOL_SOCKET, inOption, inValue, 4) = SOCKET_ERROR Then
                Call RaiseError(inSox, Err.LastDllError, Procedure, "Option: " & inOption & " & Value: " & inValue)
                Let SetOption = SOCKET_ERROR
            End If
        End Select
    End If

End Function

Public Function Shut(inSox As Long) As Long ' Initiates ShutDown procedure for a Socket

Const Procedure As String = "Shut"

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(INVALID_SOCKET, 0, Procedure, "Invalid Sox: " & inSox)
        Let Shut = INVALID_SOCKET
    Else
        Select Case Sockets(inSox).State
        Case soxDisconnected: Call RaiseError(INVALID_SOCKET, 0, Procedure, "Sox: " & inSox & " already closed")
        Case soxClosing: Call RaiseError(INVALID_SOCKET, 0, Procedure, "Sox: " & inSox & " already closing")
        Case soxBound
            If apiWSAAsyncSelect(Sockets(inSox).Socket, UserControl.hWnd, 0&, 0&) = SOCKET_ERROR Then       'FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
                Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiWSAAsyncSelect")
                Let Shut = SOCKET_ERROR
            Else
                If apiCloseSocket(Sockets(inSox).Socket) = SOCKET_ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                    Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiCloseSocket")
                    Let Shut = SOCKET_ERROR
                Else
                    Call RaiseState(inSox, soxDisconnected)
                    RaiseEvent OnStatus(inSox, Procedure, "Successfully Closed")
                    RaiseEvent OnClose(inSox)
                End If
            End If
        Case soxListening
            If apiWSAAsyncSelect(Sockets(inSox).Socket, UserControl.hWnd, 0&, 0&) = SOCKET_ERROR Then       'FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
                Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiWSAAsyncSelect")
                Let Shut = SOCKET_ERROR
            Else
                If apiCloseSocket(Sockets(inSox).Socket) = SOCKET_ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                    Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiCloseSocket")
                    Let Shut = SOCKET_ERROR
                Else
                    Call RaiseState(inSox, soxDisconnected)
                    RaiseEvent OnStatus(inSox, Procedure, "Successfully Closed")
                    RaiseEvent OnClose(inSox)
                End If
            End If
        Case Else
            If apiGetSockOpt(Sockets(inSox).Socket, SOL_SOCKET, soxSO_ERROR, Shut, 4) = SOCKET_ERROR Then
                Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiGetSockOpt")
                Let Shut = SOCKET_ERROR
            Else
                If apiShutDown(Sockets(inSox).Socket, SD_SEND) = SOCKET_ERROR Then
                    Call RaiseError(inSox, Err.LastDllError, Procedure, "On apiShutDown")
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

Const Procedure As String = "ShutDown"
Dim tmpSox As Long

    For tmpSox = 0 To Portal.Sockets
        Select Case Sockets(tmpSox).State
        Case soxDisconnected
        Case soxClosing ' No need to close a closing Socket
        Case soxBound ' Same as soxListening
            If apiWSAAsyncSelect(Sockets(tmpSox).Socket, UserControl.hWnd, 0&, 0&) = SOCKET_ERROR Then       'FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
                Call RaiseError(tmpSox, Err.LastDllError, Procedure, "On apiWSAAsyncSelect")
                Call RaiseState(tmpSox, soxDisconnected)
                RaiseEvent OnClose(tmpSox)
                Let ShutDown = SOCKET_ERROR
            Else
                If apiCloseSocket(Sockets(tmpSox).Socket) = SOCKET_ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                    Call RaiseError(tmpSox, Err.LastDllError, Procedure, "On apiCloseSocket")
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
                Call RaiseError(tmpSox, Err.LastDllError, Procedure, "On apiWSAAsyncSelect")
                Call RaiseState(tmpSox, soxDisconnected)
                RaiseEvent OnClose(tmpSox)
                Let ShutDown = SOCKET_ERROR
            Else
                If apiCloseSocket(Sockets(tmpSox).Socket) = SOCKET_ERROR Then  ' I can't get the API that checks the current status of the socket to work :(((
                    Call RaiseError(tmpSox, Err.LastDllError, Procedure, "On apiCloseSocket")
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
                Call RaiseError(tmpSox, Err.LastDllError, Procedure, "On apiShutDown")
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
        Call RaiseError(soxERROR, Err.LastDllError, Procedure, "ShutDown Pending!")
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

Const Procedure As String = "Socket"

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(soxERROR, 0, Procedure, "Invalid Sox: " & inSox)
        Let SocketHandle = soxERROR
    Else
        Let SocketHandle = Sockets(inSox).Socket
    End If

End Function

Public Function State(inSox As Long) As enmSoxState

Const Procedure As String = "State"

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(soxERROR, 0, Procedure, "Invalid Sox: " & inSox)
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

Const Procedure As String = "uMsg"

    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(inSox, 0, Procedure, "Critical Error, Invalid Sox: " & inSox) ' This should NEVER happen, why? Because WinSock API closed a valid Socket that we didn't even know about!
        Let uMsg = soxERROR
    Else
        Let uMsg = Sockets(inSox).uMsg
    End If

End Function

Public Sub UnHook() ' Once the Control is UnHooked, we will not be able to intercept messages from WinSock API and process them according to our needs!

Const Procedure As String = "UnHook" ' The messages WinSock API will be sending us are FD_ACCEPT, FD_WRITE, FD_READ etc. Notifying us of socket state changes

    Call apiSetWindowLong(UserControl.hWnd, GWL_WNDPROC, Portal.WndProc)
    Let Portal.WndProc = 0
    RaiseEvent OnStatus(soxERROR, Procedure, "Message hook disabled, Sox Control UnHooked") ' soxERROR doesn't actually indicate an error here at all!!!

End Sub

Private Sub UserControl_Initialize() ' NB: This code actually gets processed in the IDE as well!

    BufferSize = 8192
    If Not InIDE Then ' The following code gets run EVERY time we double click the Control and frmMain, this creates problems for us!!! So, this prevents us from using WinSock and SubClassing from within the IDE!!!
        Set SocketControl = Me
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

    Let UserControl.Width = Screen.TwipsPerPixelX * 44
    Let UserControl.Height = Screen.TwipsPerPixelX * 44

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

' First time I've ever NEEDED 'Friend' :))) Because this is not for general use! Only used by modSox!!!

Friend Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const Procedure As String = "WndProc"

    Select Case uMsg
    Case soxSERVER
        Select Case WinSockEvent(lParam)
        Case FD_ACCEPT: RaiseEvent OnStatus(Socket2Sox(wParam), Procedure, "Server received: FD_ACCEPT")
            Select Case WinSockError(lParam)
            Case 0: Call Accept(wParam)
            Case Else: Call RaiseError(Socket2Sox(wParam), WinSockError(lParam), Procedure, "On FD_ACCEPT -- lParam: " & lParam)
            End Select
        Case FD_CLOSE: RaiseEvent OnStatus(Socket2Sox(wParam), Procedure, "Server received: FD_CLOSE")
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
                Call RaiseError(Socket2Sox(wParam), WinSockError(lParam), Procedure, "On FD_CLOSE -- lParam: " & lParam)
            End Select
        Case FD_CONNECT: RaiseEvent OnStatus(Socket2Sox(wParam), Procedure, "Server received: FD_CONNECT")
        Case FD_READ: RaiseEvent OnStatus(Socket2Sox(wParam), Procedure, "Server received: FD_READ")
            Call GetData(Socket2Sox(wParam))
        Case FD_WRITE: RaiseEvent OnStatus(Socket2Sox(wParam), Procedure, "Server received: FD_WRITE") ' A Server Client is ready to Send
            Select Case WinSockError(lParam)
            Case 0
                Select Case Sockets(Socket2Sox(wParam)).State
                Case soxConnecting
                    RaiseEvent OnConnect(Socket2Sox(wParam))
                    Call RaiseState(Socket2Sox(wParam), soxIdle)
                Case soxIdle: Call SendBuffer(Socket2Sox(wParam))
                Case soxClosing: Call Closed(Socket2Sox(wParam)) 'Call RaiseState(Socket2Sox(wParam), soxDisconnected)
                End Select
            Case Else: Call RaiseError(Socket2Sox(wParam), WinSockError(lParam), Procedure, "On FD_WRITE -- lParam: " & lParam)
            End Select
        Case Else
        End Select
    Case soxCLIENT
        Select Case WinSockEvent(lParam)
        Case FD_ACCEPT: RaiseEvent OnStatus(Socket2Sox(wParam), Procedure, "Client received: FD_ACCEPT") ' This should never happen!
        Case FD_CLOSE: RaiseEvent OnStatus(Socket2Sox(wParam), Procedure, "Client received: FD_CLOSE")
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
                Call RaiseError(Socket2Sox(wParam), WinSockError(lParam), Procedure, "On FD_CLOSE -- lParam: " & lParam)
            End Select
        Case FD_CONNECT: RaiseEvent OnStatus(Socket2Sox(wParam), Procedure, "Client received: FD_CONNECT")
        Case FD_READ: RaiseEvent OnStatus(Socket2Sox(wParam), Procedure, "Client received: FD_READ")
            Call GetData(Socket2Sox(wParam))
        Case FD_WRITE: RaiseEvent OnStatus(Socket2Sox(wParam), Procedure, "Client received: FD_WRITE")
            Select Case WinSockError(lParam)
            Case 0
                Select Case Sockets(Socket2Sox(wParam)).State
                Case soxConnecting
                    RaiseEvent OnConnect(Socket2Sox(wParam))
                    Call RaiseState(Socket2Sox(wParam), soxIdle)
                Case soxIdle: Call SendBuffer(Socket2Sox(wParam))
                Case soxClosing: Call Closed(Socket2Sox(wParam)) 'Call RaiseState(Socket2Sox(wParam), soxDisconnected)
                End Select
            Case Else: Call RaiseError(Socket2Sox(wParam), WinSockError(lParam), Procedure, "On FD_WRITE -- lParam: " & lParam)
            End Select
        Case Else
        End Select
    Case Else: Let WndProc = apiCallWindowProc(Portal.WndProc, hWnd, uMsg, wParam, lParam)
    End Select

End Function

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 18:16)  Decl: 468  Code: 1055  Total: 1523 Lines
':) CommentOnly: 157 (10.3%)  Commented: 381 (25%)  Empty: 113 (7.4%)  Max Logic Depth: 9

