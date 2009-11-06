VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This byte buffer was created for being able to automatically convert variables to byte arrays
'and to be able to access them by retrieving them in the order stored. The main intentions of this
'was for creating packets, so you could easily send binary packets (in a byte array) instead of
'having to send strings (which are larger then a byte array, along with slower).
'IE:
'The value "12340" in a string would be 5 bytes, while as an integer, only be 2 bytes. The really
'noticeable values come in at Longs, such as "1293049392" as a string (10 bytes)
'VS as a long (4 bytes).
'This buffer supports strings only of 256 characters long because the length of a string is stored
'as a byte. Though, this can easily be changed by changing Put_Byte / Get_Byte in
'Put_String / Get_String to Put_Integer / Get_Integer, or if you are sending huge strings, you can
'use longs. Try to keep it to the smallest variable you can, though, since it increases the length
'of every string you use.
'This is required to be in there since variable-length strings (in oppose to fixed length strings)
'can not be stored in the same manner as integers, bytes and longs. How would the program know
'if "ABC" was 2 strings "AB" and "C", or "A" and "BC". Instead, it appears "3ABC" so the program
'knows that the next 3 bytes are part of the same string.
'This works very similar to file I/O in binary. You have to retrieve in the same order as you store,
'and you can only retrieve what you have already stored. I did not add in error checking for
'recieving past the stored size just for the slightest speed increase (every ms counts!), along
'with that you shouldn't be even trying to retrieve past what you have stored.
'
'NEW: Allocate and Pre-Allocate
'These routines are put in to prevent having to use ReDim Preserve over and over again, as it is VERY
'slow to constantly use (especially byte-by-byte). Both of the commands will resize the array by the
'amount of bytes that you specify. In an optimal situation, you will call PreAllocate once and fill
'the whole buffer. It is okay to be short some bytes (such as if you know a string will always be
'longer then 10 bytes, but no idea how long in total) since it will auto-resize as needed if you don't
'specify enough of a preallocating. Though, it is best to never over-size the array, or else you will
'be sending extra, empty data.
'Allocate: Saves the buffer currently existing, then allocates X amount of bytes.
'PreAllocate: ReDims (no preserve) the buffer, and clears the buffer variables. If you use this, only
' use it at the start of filling the buffer, and in subsitute of using .Clear.

Private PutBytePos As Long
Private GetBytePos As Long
Private ByteBuffer() As Byte
Private ByteBufferUbound As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub Overflow()

    'Force the buffer to overflow to break the packet reading loop
    GetBytePos = ByteBufferUbound + 10

End Sub

Public Function HasBuffer() As Byte

    'If there is a buffer or not
    If PutBytePos > 0 Then HasBuffer = 1

End Function

Public Sub Clear()

'Clear all the values so we can use the buffer from the start again

    PutBytePos = 0
    GetBytePos = 0
    ByteBufferUbound = -1
    Erase ByteBuffer

End Sub

Public Function Get_Buffer_Remainder() As Byte()
Dim b() As Byte

'Return the remainder of the byte array buffer

    'Check if we already hit the end of the buffer
    If UBound(ByteBuffer) - GetBytePos + 1 = 0 Then
        Get_Buffer_Remainder = b
    Else
        ReDim b(0 To (UBound(ByteBuffer) - GetBytePos + 1))
        CopyMemory b(0), ByteBuffer(GetBytePos), UBound(ByteBuffer) - GetBytePos + 1
        Get_Buffer_Remainder = ByteBuffer
    End If
    
End Function

Public Function Get_Buffer() As Byte()

    'Return the byte array buffer
    Get_Buffer = ByteBuffer

End Function

Public Function Get_Byte() As Byte

'Retrieve a byte from the buffer (1 byte)

    On Error GoTo ErrOut '//\\LOGLINE//\\
    If GetBytePos > ByteBufferUbound Then GoTo ErrOut '//\\LOGLINE//\\
    If GetBytePos > ByteBufferUbound Then Exit Function   'Non-log style
    
    CopyMemory Get_Byte, ByteBuffer(GetBytePos), 1
    GetBytePos = GetBytePos + 1
    On Error GoTo 0 '//\\LOGLINE//\\

Exit Function '//\\LOGLINE//\\

ErrOut: '//\\LOGLINE//\\

    Log "Get_Byte: Could not acquire BYTE from buffer. GetBytePos: " & GetBytePos & " Ubound(ByteBuffer): " & UBound(ByteBuffer), CriticalError '//\\LOGLINE//\\

End Function

Public Function Get_Integer() As Integer

'Retrieve an integer from the buffer (2 bytes)

    On Error GoTo ErrOut '//\\LOGLINE//\\
    If GetBytePos + 1 > ByteBufferUbound Then GoTo ErrOut '//\\LOGLINE//\\
    If GetBytePos + 1 > ByteBufferUbound Then Exit Function  'Non-log style

    CopyMemory Get_Integer, ByteBuffer(GetBytePos), 2
    GetBytePos = GetBytePos + 2
    On Error GoTo 0 '//\\LOGLINE//\\
    
Exit Function '//\\LOGLINE//\\

ErrOut: '//\\LOGLINE//\\

    Log "Get_Integer: Could not acquire INTEGER from buffer. GetBytePos: " & GetBytePos & " Ubound(ByteBuffer): " & UBound(ByteBuffer), CriticalError '//\\LOGLINE//\\

End Function

Public Function Get_Long() As Long

'Retrieve a long from the buffer (4 bytes)

    On Error GoTo ErrOut '//\\LOGLINE//\\
    If GetBytePos + 3 > ByteBufferUbound Then GoTo ErrOut '//\\LOGLINE//\\
    If GetBytePos + 3 > ByteBufferUbound Then Exit Function  'Non-log style

    CopyMemory Get_Long, ByteBuffer(GetBytePos), 4
    GetBytePos = GetBytePos + 4
    On Error GoTo 0 '//\\LOGLINE//\\

Exit Function '//\\LOGLINE//\\

ErrOut: '//\\LOGLINE//\\

    Log "Get_Long: Could not acquire INTEGER from buffer. GetBytePos: " & GetBytePos & " Ubound(ByteBuffer): " & UBound(ByteBuffer), CriticalError '//\\LOGLINE//\\

End Function

Private Function Get_PutPos() As Long

'Return the put byte position

    Get_PutPos = PutBytePos

End Function

Public Function Get_ReadPos() As Long

'Return the read byte position

    Get_ReadPos = GetBytePos

End Function

Public Function Get_String() As String

Dim tempB() As Byte
Dim ArraySize As Byte

    On Error GoTo ErrOut

    'Retrieve a string from the buffer
    ArraySize = Get_Byte    'Get the size of the string

    'Check for a valid size before sizing the array
    If ArraySize = 0 Then Exit Function

    'Resize the temp byte array to fit the size of the string
    ReDim tempB(ArraySize - 1)

    'Copy the bytes for the string in the buffer to the temp byte array
    CopyMemory tempB(0), ByteBuffer(GetBytePos), ArraySize

    'Convert the byte array to Unicode
    Get_String = StrConv(tempB, vbUnicode)
    GetBytePos = GetBytePos + ArraySize
    
    On Error GoTo 0

Exit Function '//\\LOGLINE//\\

ErrOut:

    Log "Get_String: Could not acquire STRING from buffer. GetBytePos: " & GetBytePos & " Ubound(ByteBuffer): " & UBound(ByteBuffer) & " Size: " & ArraySize, CriticalError '//\\LOGLINE//\\

End Function

Public Function Get_StringEX() As String

Dim tempB() As Byte
Dim ArraySize As Integer

    On Error GoTo ErrOut

    'Retrieve a very long string from the buffer
    ArraySize = Get_Integer 'Get the size of the string

    'Check for a valid size before sizing the array
    If ArraySize = 0 Then Exit Function

    'Resize the temp byte array to fit the size of the string
    ReDim tempB(ArraySize - 1)

    'Copy the bytes for the string in the buffer to the temp byte array
    CopyMemory tempB(0), ByteBuffer(GetBytePos), ArraySize

    'Convert the byte array to Unicode
    Get_StringEX = StrConv(tempB, vbUnicode)
    GetBytePos = GetBytePos + ArraySize
    
    On Error GoTo 0

Exit Function '//\\LOGLINE//\\

ErrOut:

    Log "Get_StringEX: Could not acquire STRING from buffer. GetBytePos: " & GetBytePos & " Ubound(ByteBuffer): " & UBound(ByteBuffer) & " Size: " & ArraySize, CriticalError '//\\LOGLINE//\\

End Function

Public Sub Put_Byte(ByVal Value As Byte)

    'Store a byte (1 byte)
    If ByteBufferUbound < PutBytePos Then
        ReDim Preserve ByteBuffer(0 To PutBytePos)
        ByteBufferUbound = PutBytePos
    End If
    CopyMemory ByteBuffer(PutBytePos), Value, 1
    PutBytePos = PutBytePos + 1

End Sub

Public Sub Put_Integer(ByVal Value As Integer)

    'Store an integer (2 bytes)
    If ByteBufferUbound < PutBytePos + 1 Then
        ReDim Preserve ByteBuffer(0 To PutBytePos + 1)
        ByteBufferUbound = PutBytePos + 1
    End If
    CopyMemory ByteBuffer(PutBytePos), Value, 2
    PutBytePos = PutBytePos + 2

End Sub

Public Sub Put_Long(ByVal Value As Long)

    'Store a long (4 bytes)
    If ByteBufferUbound < PutBytePos + 3 Then
        ReDim Preserve ByteBuffer(0 To PutBytePos + 3)
        ByteBufferUbound = PutBytePos + 3
    End If
    CopyMemory ByteBuffer(PutBytePos), Value, 4
    PutBytePos = PutBytePos + 4

End Sub

Public Sub Put_String(ByRef Value As String)
Dim tempB() As Byte
Dim i As Long

    'Store a string

    'Check for an empty string
    If Value = vbNullString Then
        Put_Byte 0
        Exit Sub
    End If

    'Cache the UBound
    i = Len(Value) - 1
    
    'Store a byte-long value that represents the size of the string
    If i > 254 Then
        Log "Put_String: The following string was too large to pack into the Put_String buffer! " & Value, CriticalError '//\\LOGLINE//\\
        Exit Sub
    End If
    Put_Byte i + 1
    
    'Convert the string to a byte array
    tempB = StrConv(Value, vbFromUnicode)

    'Resize the array to fit the string
    If ByteBufferUbound < PutBytePos + i Then
        ReDim Preserve ByteBuffer(0 To PutBytePos + i)
        ByteBufferUbound = PutBytePos + i
    End If
    
    'Store the byte array of the string into the buffer byte array
    CopyMemory ByteBuffer(PutBytePos), tempB(0), i + 1
    PutBytePos = PutBytePos + i + 1

End Sub

Public Sub Put_StringEX(ByRef Value As String)
Dim tempB() As Byte
Dim i As Long

    'Store a very long string
    
    'Check for an empty string
    If Value = vbNullString Then
        Put_Integer 0
        Exit Sub
    End If
    
    'Cache the UBound
    i = Len(Value) - 1

    'Store a byte-long value that represents the size of the string
    If i > 32760 Then
        Log "Put_StringEX: The following string was too large to pack into the Put_StringEX buffer! " & Value, CriticalError '//\\LOGLINE//\\
        Exit Sub
    End If
    Put_Integer i + 1
    
    'Convert the string to a byte array
    tempB = StrConv(Value, vbFromUnicode)
 
    'Resize the array to fit the string
    If ByteBufferUbound < PutBytePos + i Then
        ReDim Preserve ByteBuffer(0 To PutBytePos + i)
        ByteBufferUbound = PutBytePos + i
    End If
    
    'Store the byte array of the string into the buffer byte array
    CopyMemory ByteBuffer(PutBytePos), tempB(0), i + 1
    PutBytePos = PutBytePos + i + 1

End Sub

Public Sub Set_Buffer(ByRef Value() As Byte)

    'Clear the values
    Clear

    'Cache the UBound
    ByteBufferUbound = UBound(Value)

    'Set the byte buffer to the size of the array being sent in
    ReDim ByteBuffer(0 To ByteBufferUbound)

    'Copy to the byte buffer
    CopyMemory ByteBuffer(0), Value(0), ByteBufferUbound + 1

End Sub

Public Sub Allocate(ByVal NumBytes As Long)

    'Allocate the memory in bulk
    ByteBufferUbound = ByteBufferUbound + NumBytes
    ReDim Preserve ByteBuffer(0 To ByteBufferUbound)

End Sub

Public Sub PreAllocate(ByVal NumBytes As Long)

    'Allocate the memory in bulk without preserving data
    ByteBufferUbound = NumBytes - 1
    PutBytePos = 0
    GetBytePos = 0
    ReDim ByteBuffer(0 To ByteBufferUbound)
    
End Sub

Private Sub Class_Initialize()
    
    'Set the buffer UBound
    ByteBufferUbound = -1

End Sub
