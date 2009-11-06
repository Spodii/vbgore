Attribute VB_Name = "Sound"
Option Explicit

Public Const SoundBufferTimerMax As Long = 300000   'How long a sound stays in memory unused (miliseconds)
Public SoundBufferTimer() As Long                   'How long until the sound buffer unloads
Public DS As DirectSound8
Public DSBDesc As DSBUFFERDESC
Public DSBuffer() As DirectSoundSecondaryBuffer8

Public Sub Sound_Init()

'************************************************************
'Initialize the 3D sound device
'************************************************************

    'Make sure we try not to load a file while the engine is unloading
    If IsUnloading Then Exit Sub
    
    On Error GoTo ErrOut
    
    If UseSfx = 0 Then Exit Sub
    
    'Create the DirectSound device (with the default device)
    Set DS = DX.DirectSoundCreate("")
    DS.SetCooperativeLevel frmMain.hwnd, DSSCL_PRIORITY
    
    'Set up the buffer description for later use
    'We are only using panning and volume - combined, we will use this to create a custom 3D effect
    DSBDesc.lFlags = DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    
    'Check if the texture exists
    If Engine_FileExist(SfxPath & "Sfx.ini", vbNormal) = False Then
        MsgBox "Error! Could not find the following data file:" & vbCrLf & SfxPath & "Sfx.ini", vbOKOnly
        IsUnloading = 1
        Exit Sub
    End If

    'Get the number of sound effects
    NumSfx = Val(Var_Get(SfxPath & "Sfx.ini", "INIT", "NumSfx"))
    
    'Resize the sound buffer array
    If NumSfx > 0 Then
        ReDim DSBuffer(1 To NumSfx)
        ReDim SoundBufferTimer(1 To NumSfx)
    End If
    
    On Error GoTo 0
    
    Exit Sub
    
ErrOut:

    'Failure loading sounds, so we won't use them
    UseSfx = 0
    UseMusic = 0

End Sub

Public Sub Sound_SetToMap(ByVal SoundID As Integer, ByVal TileX As Byte, ByVal TileY As Byte)

'************************************************************
'Create a looping sound on the tile
'************************************************************

    If UseSfx = 0 Then Exit Sub

    'Make sure the sound isn't already going
    If Not MapData(TileX, TileY).Sfx Is Nothing Then
        MapData(TileX, TileY).Sfx.Stop
        Set MapData(TileX, TileY).Sfx = Nothing
    End If
    
    'Create the buffer
    Sound_Set MapData(TileX, TileY).Sfx, SoundID
    
    'Exit if theres an error
    If MapData(TileX, TileY).Sfx Is Nothing Then Exit Sub

    'Start the loop
    MapData(TileX, TileY).Sfx.Play DSBPLAY_LOOPING
    
    'Since we dont want to start hearing the sound until we have calculated the panning/volume, we set the volume to off for now
    MapData(TileX, TileY).Sfx.SetVolume -10000

End Sub

Public Sub Sound_UpdateMap()

'************************************************************
'Update the panning and volume on the map's sfx
'************************************************************
Dim SX As Integer
Dim SY As Integer
Dim X As Byte
Dim Y As Byte
Dim L As Long

    If UseSfx = 0 Then Exit Sub

    'Set the user's position to sX/sY
    SX = CharList(UserCharIndex).Pos.X
    SY = CharList(UserCharIndex).Pos.Y
    
    'Loop through all the map tiles
    For X = 1 To MapInfo.Width
        For Y = 1 To MapInfo.Height
            
            'Only update used tiles
            If Not MapData(X, Y).Sfx Is Nothing Then
                
                'Calculate the volume and check for valid range
                L = Sound_CalcVolume(SX, SY, X, Y)
                If L < -5000 Then
                    MapData(X, Y).Sfx.Stop
                Else
                    If L > 0 Then L = 0
                    If MapData(X, Y).Sfx.GetStatus <> DSBSTATUS_LOOPING Then MapData(X, Y).Sfx.Play DSBPLAY_LOOPING
                    MapData(X, Y).Sfx.SetVolume L
                End If
                
                'Calculate the panning and check for a valid range
                L = Sound_CalcPan(SX, X)
                If L > 10000 Then L = 10000
                If L < -10000 Then L = -10000
                MapData(X, Y).Sfx.SetPan L
                
            End If
            
        Next Y
    Next X

End Sub

Public Sub Sound_Play(ByRef SoundBuffer As DirectSoundSecondaryBuffer8, Optional ByVal flags As CONST_DSBPLAYFLAGS = DSBPLAY_DEFAULT)
'************************************************************
'Used for non area-specific sound effects, such as weather
'************************************************************

    If UseSfx = 0 Then Exit Sub

    'Play the sound
    If Not SoundBuffer Is Nothing Then SoundBuffer.Play flags
    
End Sub

Public Sub Sound_Erase(ByRef SoundBuffer As DirectSoundSecondaryBuffer8)

'************************************************************
'Erase the sound buffer
'************************************************************

    If UseSfx = 0 Then Exit Sub
    
    'Make sure the object exists
    If Not SoundBuffer Is Nothing Then
    
        'If it is playing, we have to stop it first
        If SoundBuffer.GetStatus > 0 Then SoundBuffer.Stop
        
        'Clear the object
        Set SoundBuffer = Nothing
        
    End If

End Sub

Public Sub Sound_Set(ByRef SoundBuffer As DirectSoundSecondaryBuffer8, ByVal SoundID As Integer)

'************************************************************
'Set the SoundID to the sound buffer
'************************************************************

    If UseSfx = 0 Then Exit Sub

    'Check if the sound buffer is in use
    Sound_Erase SoundBuffer
    
    'Set the buffer
    If Engine_FileExist(SfxPath & SoundID & ".wav", vbNormal) Then Set SoundBuffer = DS.CreateSoundBufferFromFile(SfxPath & SoundID & ".wav", DSBDesc)

End Sub

Public Sub Sound_Play3D(ByVal SoundID As Integer, ByVal TileX As Integer, ByVal TileY As Integer)

'************************************************************
'Play a pseudo-3D sound by the sound buffer ID
'************************************************************
Dim SX As Integer
Dim SY As Integer

    If UseSfx = 0 Then Exit Sub

    'Make sure we have the UserCharIndex, or else we cant play the sound! :o
    If UserCharIndex = 0 Then Exit Sub

    'Check for a valid sound
    If SoundID <= 0 Then Exit Sub

    'Create the buffer if needed
    If SoundBufferTimer(SoundID) < timeGetTime Then
        If DSBuffer(SoundID) Is Nothing Then Sound_Set DSBuffer(SoundID), SoundID
    End If
    
    'Update the timer
    SoundBufferTimer(SoundID) = timeGetTime + SoundBufferTimerMax
    
    'Clear the position (used in case the sound was already playing - we can only have one of each sound play at a time)
    DSBuffer(SoundID).SetCurrentPosition 0
    
    'Set the user's position to sX/sY
    SX = CharList(UserCharIndex).Pos.X
    SY = CharList(UserCharIndex).Pos.Y
    
    'Calculate the panning
    Sound_Pan DSBuffer(SoundID), Sound_CalcPan(SX, TileX)
    
    'Calculate the volume
    Sound_Volume DSBuffer(SoundID), Sound_CalcVolume(SX, SY, TileX, TileY)
    
    'Play the sound
    DSBuffer(SoundID).Play DSBPLAY_DEFAULT
    
End Sub

Public Function Sound_CalcPan(ByVal x1 As Integer, ByVal x2 As Integer) As Long

'************************************************************
'Calculate the panning for 3D sound based on the user's position and the sound's position
'************************************************************

    If UseSfx = 0 Then Exit Function

    Sound_CalcPan = (x1 - x2) * 75 * ReverseSound
    
End Function

Public Function Sound_CalcVolume(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer) As Long

'************************************************************
'Calculate the volume for 3D sound based on the user's position and the sound's position
'the (Abs(sX - TileX) * 25) is put on the end to make up for the simulated
' volume loss during panning (since one speaker gets muted to create the panning)
'************************************************************
Dim Dist As Single

    If UseSfx = 0 Then Exit Function

    'Store the distance
    Dist = Sqr(((Y1 - Y2) * (Y1 - Y2)) + ((x1 - x2) * (x1 - x2)))
    
    'Apply the initial value
    Sound_CalcVolume = -(Dist * 80) + (Abs(x1 - x2) * 25)
    
    'Once we get out of the screen (>= 13 tiles away) then we want to fade fast
    If Dist > 13 Then Sound_CalcVolume = Sound_CalcVolume - ((Dist - 13) * 180)
    
End Function

Private Sub Sound_Pan(ByRef SoundBuffer As DirectSoundSecondaryBuffer8, ByVal Value As Long)

'************************************************************
'Pan the selected SoundID (-10,000 to 10,000)
'************************************************************

    If UseSfx = 0 Then Exit Sub

    If SoundBuffer Is Nothing Then Exit Sub
    SoundBuffer.SetPan Value

End Sub

Private Sub Sound_Volume(ByRef SoundBuffer As DirectSoundSecondaryBuffer8, ByVal Value As Long)

'************************************************************
'Pan the selected SoundID (-10,000 to 0)
'************************************************************

    If UseSfx = 0 Then Exit Sub

    If SoundBuffer Is Nothing Then Exit Sub
    If Value > 0 Then Value = 0
    If Value < -10000 Then Value = -10000
    SoundBuffer.SetVolume Value

End Sub

Public Sub Music_Load(ByVal FilePath As String, ByVal BufferNumber As Long)

'************************************************************
'Loads a mp3 by the specified path
'************************************************************

    If UseMusic = 0 Then Exit Sub

    On Error GoTo Error_Handler
                
    If Right$(FilePath, 4) = ".mp3" Then
    
        Set DirectShow_Control(BufferNumber) = New FilgraphManager
        DirectShow_Control(BufferNumber).RenderFile FilePath
    
        Set DirectShow_Audio(BufferNumber) = DirectShow_Control(BufferNumber)
        
        DirectShow_Audio(BufferNumber).Volume = 0
        DirectShow_Audio(BufferNumber).Balance = 0
    
        Set DirectShow_Event(BufferNumber) = DirectShow_Control(BufferNumber)
        Set DirectShow_Position(BufferNumber) = DirectShow_Control(BufferNumber)
        
        DirectShow_Position(BufferNumber).Rate = 1
        
        DirectShow_Position(BufferNumber).CurrentPosition = 0
    
    End If

Error_Handler:

End Sub

Public Sub Music_Play(ByVal BufferNumber As Long)

'************************************************************
'Plays the mp3 in the specified buffer
'************************************************************
    
    On Error GoTo Error_Handler
    
    If UseMusic = 0 Then Exit Sub
    
    DirectShow_Control(BufferNumber).Run

Error_Handler:

End Sub

Public Sub Music_Stop(ByVal BufferNumber As Long)

'************************************************************
'Stops the mp3 in the specified buffer
'************************************************************

    On Error GoTo Error_Handler
    
    If UseMusic = 0 Then Exit Sub
    
    DirectShow_Control(BufferNumber).Stop
    
    DirectShow_Position(BufferNumber).CurrentPosition = 0

    Exit Sub

Error_Handler:

End Sub

Public Sub Music_Pause(ByVal BufferNumber As Long)

'************************************************************
'Pause the music in the specified buffer
'************************************************************

    On Error GoTo Error_Handler
    
    If UseMusic = 0 Then Exit Sub
    
    DirectShow_Control(BufferNumber).Stop
    
Error_Handler:

End Sub

Public Sub Music_Volume(ByVal Volume As Long, ByVal BufferNumber As Long)

'************************************************************
'Set the volume of the music in the specified buffer
'************************************************************

    On Error GoTo Error_Handler
    
    If UseMusic = 0 Then Exit Sub
    
    If Volume >= Music_MaxVolume Then Volume = Music_MaxVolume
    
    If Volume <= 0 Then Volume = 0
    
    DirectShow_Audio(BufferNumber).Volume = (Volume * Music_MaxVolume) - 10000
    
Error_Handler:

End Sub

Public Sub Music_Balance(ByVal Balance As Long, ByVal BufferNumber As Long)

'************************************************************
'Set the balance of the music in the specified buffer
'************************************************************

    On Error GoTo Error_Handler
    
    If UseMusic = 0 Then Exit Sub
    
    If Balance >= Music_MaxBalance Then Balance = Music_MaxBalance
    
    If Balance <= -Music_MaxBalance Then Balance = -Music_MaxBalance
    
    DirectShow_Audio(BufferNumber).Balance = Balance * Music_MaxBalance

Error_Handler:

End Sub

Public Sub Music_Speed(ByVal Speed As Single, ByVal BufferNumber As Long)

'************************************************************
'Set the speed of the music in the specified buffer
'************************************************************

    On Error GoTo Error_Handler
    
    If UseMusic = 0 Then Exit Sub

    If Speed >= Music_MaxSpeed Then Speed = Music_MaxSpeed
    
    If Speed <= 0 Then Speed = 0

    DirectShow_Position(BufferNumber).Rate = Speed / 100

Error_Handler:

End Sub

Public Sub Music_SetPosition(ByVal Hours As Long, ByVal Minutes As Long, ByVal Seconds As Long, Milliseconds As Single, ByVal BufferNumber As Long)
    
'************************************************************
'Set the speed of the music in the specified buffer
'************************************************************
    
    On Error GoTo Error_Handler
    
    Dim Max_Position As Single
    
    Dim Position As Double
    
    Dim Decimal_Milliseconds As Single
    
    If UseMusic = 0 Then Exit Sub
    
    'Keep minutes within range
    Minutes = Minutes Mod 60
        
    'Keep seconds within range
    Seconds = Seconds Mod 60
        
    'Keep milliseconds within range and keep decimal
    Decimal_Milliseconds = Milliseconds - Int(Milliseconds)
    Milliseconds = Milliseconds Mod 1000
    Milliseconds = Milliseconds + Decimal_Milliseconds
    
    'Convert Minutes & Seconds to Position time
    Position = (Hours * 3600) + (Minutes * 60) + Seconds + (Milliseconds * 0.001)
    
    Max_Position = DirectShow_Position(BufferNumber).StopTime

    If Position >= Max_Position Then
        Position = 0
        GoTo Error_Handler
    End If
    
    If Position <= 0 Then
        Position = 0
        GoTo Error_Handler
    End If
    
    DirectShow_Position(BufferNumber).CurrentPosition = Position

Error_Handler:

End Sub

Public Sub Music_End(ByVal BufferNumber As Long)

'************************************************************
'End the music in the specified buffer
'************************************************************

    On Error GoTo Error_Handler
    
    If UseMusic = 0 Then Exit Sub
    
    'Check if the buffer is looping
    If Not Music_Loop(BufferNumber) Then
    
        'Check if the current position is past the stop time
        If DirectShow_Position(BufferNumber).CurrentPosition >= DirectShow_Position(BufferNumber).StopTime Then Music_Stop BufferNumber
    
    End If

Error_Handler:

End Sub

Public Function Music_Loop(ByVal Media_Number As Long) As Boolean

'************************************************************
'Loop the music in the specified buffer
'************************************************************

    On Error GoTo Error_Handler
    
    If UseMusic = 0 Then Exit Function
    
    'Check if the current position is past the stop time - if so, reset it
    If DirectShow_Position(Media_Number).CurrentPosition >= DirectShow_Position(Media_Number).StopTime Then
        DirectShow_Position(Media_Number).CurrentPosition = 0
    End If
    
    Music_Loop = True

    Exit Function

Error_Handler:

    Music_Loop = False

End Function
