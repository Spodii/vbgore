Attribute VB_Name = "Sound"
Option Explicit
Private Const NumSoundBuffers = 7
'Sound
'//Public DirectSound As DirectSound
'//Dim DSBuffers(1 To NumSoundBuffers) As DirectSoundBuffer
Private LastSoundBufferUsed As Integer
Private SoundPlaying As Boolean

'//Public Sub DeInitiliazeSound()

'***************************************************
'Sets the sound's volume
'***************************************************
'Reset any channels that are done
'//For LoopC = 1 To NumSoundBuffers
'//    Set DSBuffers(LoopC) = Nothing
'//Next LoopC
'//Set DirectSound = Nothing

'//End Sub

'//Public Function InitializeSound(ByRef DX As DirectX7) As Boolean
'***************************************************
'Sets the sound's volume
'***************************************************
'//    Set DirectSound = DX.DirectSoundCreate("")
'//    DirectSound.SetCooperativeLevel DisplayFormhWnd, DSSCL_PRIORITY
'//    LastSoundBufferUsed = 0
'//End Function

Public Function LoadWavetoDSBuffer(ByVal file As String, Optional ByVal LoopSound As Boolean = False) As Boolean

'***************************************************
'Sets the sound's volume
'***************************************************

Dim BufferDesc As DSBUFFERDESC
Dim WAVEFORMAT As WAVEFORMATEX

    If Not Engine_FileExist(SoundPath & file, vbNormal) Then
        LoadWavetoDSBuffer = False
        Exit Function
    End If
    LastSoundBufferUsed = LastSoundBufferUsed + 1
    If LastSoundBufferUsed > NumSoundBuffers Then
        LastSoundBufferUsed = 1
    End If
    BufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    WAVEFORMAT.nFormatTag = WAVE_FORMAT_PCM
    WAVEFORMAT.nChannels = 2
    WAVEFORMAT.lSamplesPerSec = 22050
    WAVEFORMAT.nBitsPerSample = 16
    WAVEFORMAT.nBlockAlign = WAVEFORMAT.nBitsPerSample / 8 * WAVEFORMAT.nChannels
    WAVEFORMAT.lAvgBytesPerSec = WAVEFORMAT.lSamplesPerSec * WAVEFORMAT.nBlockAlign
    '//Set DSBuffers(LastSoundBufferUsed) = DirectSound.CreateSoundBufferFromFile(IniPath & SoundPath & file, BufferDesc, WAVEFORMAT)
    '//PlayWave LoopSound
    LoadWavetoDSBuffer = True

End Function

Private Function WaveIsPlaying() As Boolean

'***************************************************
'Sets the sound's volume
'***************************************************

    WaveIsPlaying = SoundPlaying

End Function

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:36)  Decl: 7  Code: 65  Total: 72 Lines
':) CommentOnly: 29 (40.3%)  Commented: 0 (0%)  Empty: 13 (18.1%)  Max Logic Depth: 2
