Attribute VB_Name = "General"
Option Explicit

Public CurrGrhNum As Long
Public CurrGrh As Grh

'DirectX 8 Objects
Private DX As DirectX8
Private D3D As Direct3D8
Private D3DX As D3DX8
Public D3DDevice As Direct3DDevice8

'Describes a transformable lit vertex
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
Private Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    Rhw As Single
    Color As Long
    tu As Single
    tv As Single
End Type
Private VertexArray(0 To 3) As TLVERTEX

'Grh data
Private Type Grh
    GrhIndex As Long
    LastCount As Long
    FrameCounter As Single
    SpeedCounter As Byte
    Started As Byte
End Type
Private Type GrhData
    SX As Integer
    SY As Integer
    FileNum As Long
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Byte
    Frames() As Long
    Speed As Single
End Type
Private NumGrhs As Long
Private NumGrhFiles As Integer   'Number of pngs
Private GrhData() As GrhData

'Surfaces
Private Const SurfaceTimerMax As Single = 30000      'How long a texture stays in memory unused (miliseconds)
Private SurfaceDB() As Direct3DTexture8          'The list of all the textures
Private SurfaceTimer() As Integer                'How long until the surface unloads
Private LastTexture As Long                      'The last texture used
Private SurfaceSize() As Point

'Describes the return from a texture init
Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Sub Main()

    InitManifest

    Load frmMain
    frmMain.Show
    DoEvents

    'Set the timer resolution
    timeBeginPeriod 1
    
    'Load the file paths
    InitFilePaths
    
    'Resize buffers
    NumGrhFiles = CInt(Engine_Var_Get(DataPath & "Grh.ini", "INIT", "NumGrhFiles"))
    ReDim SurfaceDB(1 To NumGrhFiles)
    ReDim SurfaceSize(1 To NumGrhFiles)
    ReDim SurfaceTimer(1 To NumGrhFiles)
    
    'Load the main form
    Load frmMain
    frmMain.Show
    
    'Load the engine
    Engine_Init_TileEngine
    Engine_Init_GrhData
    
    'Get the first value
    CurrGrhNum = GetNextUncategorizedGrh
    Engine_Init_Grh CurrGrh, CurrGrhNum

End Sub

Function GetNextUncategorizedGrh() As Long
Dim TempSplit() As String
Dim FileNum As Byte
Dim ln As String

    On Error GoTo ErrOut

    'Loop through the GrhRaw.txt
    FileNum = FreeFile
    Open Data2Path & "GrhRaw.txt" For Binary As #FileNum
        While Not EOF(FileNum)
            
            'Get the line
            Line Input #FileNum, ln
            
            'Check if categorized
            If ln <> "" Then
                If InStr(1, ln, "(") = 0 Then   'The ( will only be in there for the categorization
                    If InStr(1, ln, "=") Then
                        TempSplit = Split(ln, "=")
                        GetNextUncategorizedGrh = Val(Right$(TempSplit(0), Len(TempSplit(0)) - 3))
                        frmMain.InfoLbl.Caption = ln
                        Close #FileNum
                        Exit Function
                    End If
                End If
            End If
            
        Wend
        
ErrOut:
        
    Close #FileNum
    
    MsgBox "All grhs are categorized!", vbOKOnly

End Function

Sub Engine_Init_Grh(ByRef Grh As Grh, ByVal GrhIndex As Long, Optional ByVal Started As Byte = 2)

'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************

    If GrhIndex <= 0 Then Exit Sub
    Grh.GrhIndex = GrhIndex
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then
            Started = 0
        End If
        Grh.Started = Started
    End If
    Grh.LastCount = timeGetTime
    Grh.FrameCounter = 1
    Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed

End Sub

Private Sub Engine_Init_Texture(ByVal TextureNum As Integer)

'*****************************************************************
'Loads a texture into memory
'*****************************************************************

Dim TexInfo As D3DXIMAGE_INFO_A
Dim FilePath As String

    'Get the path
    FilePath = GrhPath & TextureNum & ".png"
    
    'Check if the texture exists
    If Engine_FileExist(FilePath, vbNormal) = False Then
        MsgBox "Error! Could not find the following texture file:" & vbCrLf & FilePath, vbOKOnly
        Exit Sub
    End If

    'Set the texture
    Set SurfaceDB(TextureNum) = D3DX.CreateTextureFromFileEx(D3DDevice, FilePath, D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, TexInfo, ByVal 0)

    'Set the size
    SurfaceSize(TextureNum).X = TexInfo.Width
    SurfaceSize(TextureNum).Y = TexInfo.Height

    'Set the texture timer
    SurfaceTimer(TextureNum) = SurfaceTimerMax

End Sub

Function Engine_FileExist(File As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************
On Error GoTo ErrOut

    If Dir$(File, FileType) <> "" Then Engine_FileExist = True

Exit Function

'An error will most likely be caused by invalid filenames (those that do not follow the file name rules)
ErrOut:

    Engine_FileExist = False
    
End Function

Private Function Engine_Init_D3DDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS)

'************************************************************
'Initialize the Direct3D Device - start off trying with the
'best settings and move to the worst until one works
'************************************************************
Dim D3DWindow As D3DPRESENT_PARAMETERS
Dim DispMode As D3DDISPLAYMODE          'Describes the display mode
Dim i As Byte

'When there is an error, destroy the D3D device and get ready to make a new one

    On Error GoTo ErrOut

    'Retrieve current display mode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

    'Set format for windowed mode
    D3DWindow.Windowed = 1  'State that using windowed mode
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
    D3DWindow.BackBufferFormat = DispMode.Format    'Use format just retrieved

    'Set the D3DDevices
    If ObjPtr(D3DDevice) Then Set D3DDevice = Nothing
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.PreviewPic.hWnd, D3DCREATEFLAGS, D3DWindow)
    
    'The Rhw will always be 1, so set it now instead of every call
    For i = 0 To 3
        VertexArray(i).Rhw = 1
    Next i
    
    'Everything was successful
    Engine_Init_D3DDevice = 1

Exit Function

ErrOut:

    'Destroy the D3DDevice so it can be remade
    Set D3DDevice = Nothing

    'Return a failure - 0
    Engine_Init_D3DDevice = 0

End Function

Function Engine_Var_Get(File As String, Main As String, Var As String) As String

'*****************************************************************
'Gets a Var from a text file
'*****************************************************************

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

    szReturn = vbNullString

    sSpaces = Space$(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish
    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File
    Engine_Var_Get = RTrim$(sSpaces)
    If Len(Engine_Var_Get) > 0 Then
        Engine_Var_Get = Left$(Engine_Var_Get, Len(Engine_Var_Get) - 1)
    Else
        Engine_Var_Get = ""
    End If
    
End Function

Sub Engine_Var_Write(File As String, Main As String, Var As String, Value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    writeprivateprofilestring Main, Var, Value, File

End Sub

Sub Engine_Render_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, Optional ByVal LoopAnim As Boolean = True, Optional ByVal Light1 As Long = -1, Optional ByVal Light2 As Long = -1, Optional ByVal Light3 As Long = -1, Optional ByVal Light4 As Long = -1, Optional ByVal Shadow As Byte = 0)

'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
Dim CurrGrhIndex As Long    'The grh index we will be working with (acquired after updating animations)
Dim SrcHeight As Integer
Dim SrcWidth As Integer
Dim FileNum As Integer

    'Check to make sure it is legal
    If Grh.GrhIndex < 1 Then Exit Sub
    If GrhData(Grh.GrhIndex).NumFrames < 1 Then Exit Sub

    'Update the animation frame
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + ((timeGetTime - Grh.LastCount) * GrhData(Grh.GrhIndex).Speed * 0.0009)
            Grh.LastCount = timeGetTime
            If Grh.FrameCounter >= GrhData(Grh.GrhIndex).NumFrames + 1 Then
                If LoopAnim = True Then
                    Do While Grh.FrameCounter >= GrhData(Grh.GrhIndex).NumFrames + 1
                        Grh.FrameCounter = Grh.FrameCounter - GrhData(Grh.GrhIndex).NumFrames
                    Loop
                Else
                    Grh.Started = 0
                    Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames  'Force the last frame and stick with it
                    Exit Sub
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrGrhIndex = GrhData(Grh.GrhIndex).Frames(Int(Grh.FrameCounter))
    
    'Set the file number in a shorter variable
    FileNum = GrhData(CurrGrhIndex).FileNum

    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrGrhIndex).TileWidth * 32 * 0.5) + 32 * 0.5
        End If
        If GrhData(CurrGrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrGrhIndex).TileHeight * 32) + 32
        End If
    End If

    'Check for in-bounds
    If X + GrhData(CurrGrhIndex).pixelWidth > 0 Then
        If Y + GrhData(CurrGrhIndex).pixelHeight > 0 Then
            If X < frmMain.ScaleWidth Then
                If Y < frmMain.ScaleHeight Then
                
                    '***** Render the texture *****
                    'Sorry if this confuses anyone - this code was placed in-line (in opposed to calling another sub/function) to
                    ' speed things up. In-line code is always faster, especially with passing as many arguments as there is
                    ' in Engine_Render_Rectangle. This code should be just about the exact same as Engine_Render_Rectangle
                    ' minus possibly a few changes to work more specified to the manner in which it is called.
                    SrcWidth = GrhData(CurrGrhIndex).pixelWidth
                    SrcHeight = GrhData(CurrGrhIndex).pixelHeight
                    
                    'Load the surface into memory if it is not in memory and reset the timer
                    If FileNum > 0 Then
                        If SurfaceTimer(FileNum) = 0 Then Engine_Init_Texture FileNum
                        SurfaceTimer(FileNum) = SurfaceTimerMax
                    End If
                
                    'Set the texture
                    If FileNum <= 0 Then
                        D3DDevice.SetTexture 0, Nothing
                    Else
                        If LastTexture <> FileNum Then
                            D3DDevice.SetTexture 0, SurfaceDB(FileNum)
                            LastTexture = FileNum
                        End If
                    End If
                
                    'Set shadowed settings - shadows only change on the top 2 points
                    If Shadow Then
                
                        SrcWidth = SrcWidth - 1
                
                        'Set the top-left corner
                        VertexArray(0).X = X + (GrhData(CurrGrhIndex).pixelWidth * 0.5)
                        VertexArray(0).Y = Y - (GrhData(CurrGrhIndex).pixelHeight * 0.5)
                
                        'Set the top-right corner
                        VertexArray(1).X = X + GrhData(CurrGrhIndex).pixelWidth + (GrhData(CurrGrhIndex).pixelWidth * 0.5)
                        VertexArray(1).Y = Y - (GrhData(CurrGrhIndex).pixelHeight * 0.5)
                
                    Else
                
                        SrcWidth = SrcWidth + 1
                        SrcHeight = SrcHeight + 1
                
                        'Set the top-left corner
                        VertexArray(0).X = X
                        VertexArray(0).Y = Y
                
                        'Set the top-right corner
                        VertexArray(1).X = X + GrhData(CurrGrhIndex).pixelWidth
                        VertexArray(1).Y = Y
                
                    End If
                
                    'Set the top-left corner
                    VertexArray(0).Color = Light1
                    VertexArray(0).tu = GrhData(CurrGrhIndex).SX / SurfaceSize(FileNum).X
                    VertexArray(0).tv = GrhData(CurrGrhIndex).SY / SurfaceSize(FileNum).Y
                
                    'Set the top-right corner
                    VertexArray(1).Color = Light2
                    VertexArray(1).tu = (GrhData(CurrGrhIndex).SX + SrcWidth) / SurfaceSize(FileNum).X
                    VertexArray(1).tv = GrhData(CurrGrhIndex).SY / SurfaceSize(FileNum).Y
                
                    'Set the bottom-left corner
                    VertexArray(2).X = X
                    VertexArray(2).Y = Y + GrhData(CurrGrhIndex).pixelHeight
                    VertexArray(2).Color = Light3
                    VertexArray(2).tu = GrhData(CurrGrhIndex).SX / SurfaceSize(FileNum).X
                    VertexArray(2).tv = (GrhData(CurrGrhIndex).SY + SrcHeight) / SurfaceSize(FileNum).Y
                
                    'Set the bottom-right corner
                    VertexArray(3).X = X + GrhData(CurrGrhIndex).pixelWidth
                    VertexArray(3).Y = Y + GrhData(CurrGrhIndex).pixelHeight
                    VertexArray(3).Color = Light4
                    VertexArray(3).tu = (GrhData(CurrGrhIndex).SX + SrcWidth) / SurfaceSize(FileNum).X
                    VertexArray(3).tv = (GrhData(CurrGrhIndex).SY + SrcHeight) / SurfaceSize(FileNum).Y

                    'Render the texture to the device
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), Len(VertexArray(0))
                
                End If
            End If
        End If
    End If

End Sub

Private Sub Engine_Init_GrhData()

'*****************************************************************
'Loads Grh.dat
'*****************************************************************

Dim Grh As Long
Dim Frame As Long

    'Get Number of Graphics
    NumGrhs = CLng(Engine_Var_Get(DataPath & "Grh.ini", "INIT", "NumGrhs"))
    
    'Resize arrays
    ReDim GrhData(1 To NumGrhs) As GrhData
    
    'Open files
    Open DataPath & "Grh.dat" For Binary As #1
    Seek #1, 1
    
    'Fill Grh List
    Get #1, , Grh
    
    Do Until Grh <= 0
    
        'Get number of frames
        Get #1, , GrhData(Grh).NumFrames
        If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
        
        If GrhData(Grh).NumFrames > 1 Then
        
            'Read a animation GRH set
            ReDim GrhData(Grh).Frames(1 To GrhData(Grh).NumFrames)
            For Frame = 1 To GrhData(Grh).NumFrames
                Get #1, , GrhData(Grh).Frames(Frame)
                If GrhData(Grh).Frames(Frame) <= 0 Then
                    GoTo ErrorHandler
                End If
            Next Frame
            
            Get #1, , GrhData(Grh).Speed
            If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
            
            'Compute width and height
            GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
            GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
            GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
            If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
            GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
            If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
            
        Else
        
            'Read in normal GRH data
            ReDim GrhData(Grh).Frames(1 To 1)
            Get #1, , GrhData(Grh).FileNum
            If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
            Get #1, , GrhData(Grh).SX
            If GrhData(Grh).SX < 0 Then GoTo ErrorHandler
            Get #1, , GrhData(Grh).SY
            If GrhData(Grh).SY < 0 Then GoTo ErrorHandler
            Get #1, , GrhData(Grh).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
            Get #1, , GrhData(Grh).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
            
            'Compute width and height
            GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / 32
            GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / 32
            GrhData(Grh).Frames(1) = Grh
            
        End If
        
        'Get Next Grh Number
        Get #1, , Grh
        
    Loop
    '************************************************
    Close #1

Exit Sub

ErrorHandler:
    Close #1
    MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Private Function Engine_Init_TileEngine() As Boolean

'*****************************************************************
'Init Tile Engine
'*****************************************************************

    '****** INIT DirectX ******
    ' Create the root D3D objects
    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate()
    Set D3DX = New D3DX8

    'Create the D3D Device
    If Engine_Init_D3DDevice(D3DCREATE_PUREDEVICE) = 0 Then
        If Engine_Init_D3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) = 0 Then
            If Engine_Init_D3DDevice(D3DCREATE_MIXED_VERTEXPROCESSING) = 0 Then
                If Engine_Init_D3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) = 0 Then
                    If Engine_Init_D3DDevice(D3DCREATE_FPU_PRESERVE) = 0 Then
                        If Engine_Init_D3DDevice(D3DCREATE_MULTITHREADED) = 0 Then
                            MsgBox "Could not init D3DDevice. Exiting..."
                            Engine_UnloadAllForms
                            End
                        End If
                    End If
                End If
            End If
        End If
    End If
    Engine_Init_RenderStates

    'Start the engine
    Engine_Init_TileEngine = True

End Function

Private Sub Engine_Init_RenderStates()

'************************************************************
'Set the render states of the Direct3D Device
'This is in a seperate sub since if using Fullscreen and device is lost
'this is eventually called to restore settings.
'************************************************************
    
    'Set the shader to be used
    D3DDevice.SetVertexShader FVF

    'Set the render states
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

End Sub

Sub Engine_UnloadAllForms()

'*****************************************************************
'Unloads all forms
'*****************************************************************

Dim frm As Form

    For Each frm In VB.Forms
        Unload frm
    Next

End Sub
