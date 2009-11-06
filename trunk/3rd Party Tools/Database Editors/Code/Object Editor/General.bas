Attribute VB_Name = "General"
Option Explicit

'********** Direct X ***********
Public Const SurfaceTimerMax As Single = 30000  'How long a texture stays in memory unused (miliseconds)
Public SurfaceDB() As Direct3DTexture8          'The list of all the textures
Public SurfaceTimer() As Integer                'How long until the surface unloads
Public LastTexture As Long                      'The last texture used
Public D3DWindow As D3DPRESENT_PARAMETERS       'Describes the viewport and used to restore when in fullscreen
Public UsedCreateFlags As CONST_D3DCREATEFLAGS  'The flags we used to create the device when it first succeeded

'DirectX 8 Objects
Private DX As DirectX8
Private D3D As Direct3D8
Public D3DX As D3DX8
Public D3DDevice As Direct3DDevice8

'Describes a transformable lit vertex
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Public Type TLVERTEX
    x As Single
    Y As Single
    Z As Single
    Rhw As Single
    Color As Long
    Specular As Long
    Tu As Single
    Tv As Single
End Type

Private VertexArray(0 To 3) As TLVERTEX

'Holds a position on a 2d grid
Public Type Position
    x As Long
    Y As Long
End Type

'Holds data about where a png can be found,
'How big it is and animation info
Public Type GrhData
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

'Points to a grhData and keeps animation info
Public Type Grh
    GrhIndex As Long
    LastCount As Long
    FrameCounter As Single
    SpeedCounter As Byte
    Started As Byte
End Type

'Bodies list
Public Type BodyData
    Walk(1 To 8) As Grh
    Attack(1 To 8) As Grh
    HeadOffset As Position
End Type

'Wings list
Public Type WingData
    Walk(1 To 8) As Grh
    Attack(1 To 8) As Grh
End Type

'Weapons list
Public Type WeaponData
    Walk(1 To 8) As Grh
    Attack(1 To 8) As Grh
End Type

'Heads list
Public Type HeadData
    Head(1 To 8) As Grh
    Blink(1 To 8) As Grh
    AgrHead(1 To 8) As Grh
    AgrBlink(1 To 8) As Grh
End Type

'Hair list
Public Type HairData
    Hair(1 To 8) As Grh
End Type

'Holds info about a object
Public Type Obj
    ObjIndex As Integer
    Amount As Integer
End Type

'Holds a world position
Public Type WorldPos
    Map As Integer  'Map
    x As Integer       'X coordinate
    Y As Integer       'Y coordinate
End Type

'User skills casted
Type Skills
    IronSkin As Byte
    Bless As Integer
    Protect As Integer
    Strengthen As Integer
    WarCurse As Integer
End Type


Public GrhData() As GrhData         'Holds data for the graphic structure
Public SurfaceSize() As Point       'Holds the size of the surfaces for SurfaceDB()
Public ObjData() As udtObjData

'FPS
Public FPS As Long
Public EndTime As Long
Public ElapsedTime As Single
Public TickPerFrame As Single
Public TimerMultiplier As Single
Public EngineBaseSpeed As Single
Public FramesPerSecCounter As Long

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

'changed from integer to long
Private NumGrhs As Long
Private NumObjs As Integer
Private NumGrhFiles As Integer
Private TilePixelHeight As Integer
Private TilePixelWidth As Integer
Public EngineRun As Boolean

'The object we're editing
Public OpenObj As udtObjData
Public OpenIndex As Integer
Public ObjGrh As Grh

'********** OUTSIDE FUNCTIONS ***********
Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Sub Editor_LoadOBJ(ByVal ObjIndex As Integer)
'Loads an object
Dim TempSplit() As String
Dim TempStr As String
Dim TempStr2 As String
Dim i As Long
Dim here As Long
    
    'Check for valid Object number
    If ObjIndex <= 0 Then Exit Sub
    
    DB_RS.Open "SELECT * FROM objects WHERE id=" & ObjIndex, DB_Conn, adOpenStatic, adLockOptimistic
    
    'Make sure the Object exists
    If DB_RS.EOF Then
        Exit Sub
    End If
    
    Objnumber = ObjIndex
    
    'Loop through every field - match up the names then set the data accordingly
    With frmMain
        
        'Load uncatagorized parts first
        .NameTxt = Trim$(DB_RS!Name)
        .PriceTxt = Val(DB_RS!price)
        .ObjTypeCombo.ListIndex = Val(DB_RS!ObjType)
        .WeaponTypeTxt = Val(DB_RS!WeaponType)
        .RangeTxt = Val(DB_RS!WeaponRange)
        .GrhTxt = Val(DB_RS!GrhIndex)
        .ProjecTxt = Val(DB_RS!UseGrh)
        .FXTxt = Val(DB_RS!UseSfx)
        .RotTxt = Val(DB_RS!ProjectileRotateSpeed)
        .StackTxt = Val(DB_RS!Stacking)
        
        'figure out what classes are checked
        'seprate sub to keep this one cleaner, it's just below
        Load_Classes Val(DB_RS!ClassReq)
        
        'Load the Catagorized parts.
        
        'Make all the text boxes invisible that need to be
        For i = .SpriteTxt.LBound To .SpriteTxt.UBound
            .SpriteTxt(i).Visible = False
            .SpriteLbl(i).Visible = False
            .SpriteTxt(i).Text = vbNullString
        Next i
        For i = .ReqLbl.LBound To .ReqLbl.UBound
            .ReqLbl(i).Visible = False
            .ReqTxt(i).Visible = False
            .ReqTxt(i).Text = vbNullString
        Next i
        For i = .StatLbl.LBound To .StatLbl.UBound
            .StatLbl(i).Visible = False
            .StatTxt(i).Visible = False
            .StatTxt(i).Text = vbNullString
        Next i
        For i = .RepLbl.LBound To .RepLbl.UBound
            .RepLbl(i).Visible = False
            .RepTxt(i).Visible = False
            .RepPercLbl(i).Visible = False
            .RepPercTxt(i).Visible = False
            .RepPercTxt(i).Text = vbNullString
            .RepTxt(i).Text = vbNullString
        Next i
        
        'Load sprite info
            here = 0
            For i = 0 To DB_RS.Fields.Count - 1
                If InStr(1, DB_RS.Fields.Item(i).Name, "sprite_", vbTextCompare) Then
                    .SpriteLbl.Item(here + i).Caption = Replace(DB_RS.Fields.Item(i).Name, "sprite_", "") '!stat_min_atack)
                    .SpriteTxt.Item(here + i).Text = Val(DB_RS(i))
                    .SpriteLbl.Item(here + i).Visible = True
                    .SpriteTxt.Item(here + i).Visible = True
                Else
                    here = here - 1
                End If
            Next i
        
        'Load requirements
            here = 0
            For i = 0 To DB_RS.Fields.Count - 1
                If InStr(1, DB_RS.Fields.Item(i).Name, "req_", vbTextCompare) Then
                    .ReqLbl.Item(here + i).Caption = Replace(DB_RS.Fields.Item(i).Name, "req_", "") '!stat_min_atack)
                    .ReqTxt.Item(here + i).Text = Val(DB_RS(i))
                    .ReqLbl.Item(here + i).Visible = True
                    .ReqTxt.Item(here + i).Visible = True
                Else
                    here = here - 1
                End If
            Next i
        
        'Load stats
            here = 0
            For i = 0 To DB_RS.Fields.Count - 1
                If InStr(1, DB_RS.Fields.Item(i).Name, "stat_", vbTextCompare) Then
                    .StatLbl.Item(here + i).Caption = Replace(DB_RS.Fields.Item(i).Name, "stat_", "") '!stat_min_atack)
                    .StatTxt.Item(here + i).Text = Val(DB_RS(i))
                    .StatLbl.Item(here + i).Visible = True
                    .StatTxt.Item(here + i).Visible = True
                Else
                    here = here - 1
                End If
            Next i
        
        'Load replenish info
        'A little complicated having to split the string up, but I'm sure either you will figure it out
        'or just leave it be and add to the database :-)
        
            here = 0
            For i = 0 To DB_RS.Fields.Count - 1
                TempStr = DB_RS.Fields.Item(i).Name
                If InStr(1, TempStr, "replenish_", vbTextCompare) And Not InStr(1, TempStr, "_percent", vbTextCompare) Then
                    TempStr2 = Replace(TempStr, "replenish_", "")
                    
                    .RepLbl.Item(here + i).Caption = TempStr2
                    .RepTxt.Item(here + i).Text = DB_RS.Fields.Item(i).Value
                    
                    .RepPercLbl.Item(here + i).Caption = TempStr2 & " %"
                
                    .RepPercTxt.Item(here + i).Text = DB_RS("replenish_" & TempStr2 & "_percent")
                    
                    .RepLbl.Item(here + i).Visible = True
                    .RepTxt.Item(here + i).Visible = True
                    .RepPercLbl.Item(here + i).Visible = True
                    .RepPercTxt.Item(here + i).Visible = True
                Else
                    here = here - 1
                End If
            Next i
        
    End With
        DB_RS.Close
    
    OpenIndex = ObjIndex

End Sub


Public Sub Load_Classes(Classes As Byte)

If Classes And 1 Then
    frmMain.Classes(0).Value = 1
Else
    frmMain.Classes(0).Value = 0
End If

If Classes And 2 Then
    frmMain.Classes(1).Value = 1
Else
    frmMain.Classes(1).Value = 0
End If

If Classes And 4 Then
    frmMain.Classes(2).Value = 1
Else
    frmMain.Classes(2).Value = 0
End If

End Sub





Function ObjExist(ByVal ObjNum As Integer, Optional ByVal DeleteIfExists As Boolean = False) As Boolean
'*****************************************************************
'Checks the database for if a user exists by the specified name
'*****************************************************************

    'Make the query
    DB_RS.Open "SELECT * FROM objects WHERE id=" & ObjNum, DB_Conn, adOpenStatic, adLockOptimistic

    'If End Of File = true, then the user doesn't exist
    If DB_RS.EOF = True Then ObjExist = False Else ObjExist = True
    
    'Close the recordset
    DB_RS.Close
    
    'Delete the npc so we can update it if it exists.
    If DeleteIfExists Then
        If ObjExist Then DB_Conn.Execute "DELETE FROM objects WHERE id=" & ObjNum
    End If

End Function

Sub Editor_SaveOBJ(ByVal ObjIndex As Integer)
'Dim FileNum As Byte
'Dim Count As Integer
Dim i As Integer
Dim here As Integer
Dim TempStr As String
 
    With frmMain
    
        'If we are updating the object, then the record must be deleted, so make sure it isn't there (or else we get a duplicate key entry error)
        ObjExist ObjIndex, True
        'Open the database with an empty table
        DB_RS.Open "SELECT * FROM objects WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
        DB_RS.AddNew
        
        'Put the data in the recordset
        DB_RS!id = ObjIndex
        DB_RS!Name = .NameTxt
    
    
    'Save the stat info
    i = .StatLbl.UBound
    For i = 0 To .StatLbl.UBound - 1
        'If it wasn't visible it didn't have a value so don't save it
        If .StatLbl.Item(i).Visible = True Then
            DB_RS("stat_" & .StatLbl(i).Caption) = .StatTxt(i).Text
        End If
    Next i

    
    'Save the sprite info
    For i = 0 To .SpriteLbl.UBound - 1
        If .SpriteLbl.Item(i).Visible = True Then
            DB_RS("sprite_" & .SpriteLbl(i).Caption) = .SpriteTxt(i).Text
        End If
    Next i
    
    
    'Save the replinish info
    For i = 0 To .RepLbl.UBound - 1
        If .RepLbl.Item(i).Visible = True Then
            DB_RS("replenish_" & .RepLbl(i).Caption) = .RepTxt(i).Text
        End If
    Next i
    
    For i = 0 To .RepPercLbl.UBound - 1
        If .RepPercLbl.Item(i).Visible = True Then
        TempStr = Replace(.RepPercLbl(i).Caption, " %", "")
            DB_RS("replenish_" & TempStr & "_percent") = .RepPercTxt(i).Text
        End If
    Next i

    'Save the requirement info
    For i = 0 To .ReqLbl.UBound - 1
        If .ReqLbl.Item(i).Visible = True Then
            DB_RS("req_" & .ReqLbl(i).Caption) = .ReqTxt(i).Text
        End If
    Next i
    
    'Save the Class info
    i = 0
    If frmMain.Classes(0).Value = 1 Then i = i + 1
    If frmMain.Classes(1).Value = 1 Then i = i + 2
    If frmMain.Classes(2).Value = 1 Then i = i + 4
    DB_RS!ClassReq = i

    'Save the rest of the data
    DB_RS!id = ObjIndex
    DB_RS!Name = .NameTxt
    DB_RS!price = .PriceTxt
    DB_RS!ObjType = Val(.ObjTypeCombo.ListIndex)
    DB_RS!WeaponType = .WeaponTypeTxt
    DB_RS!WeaponRange = .RangeTxt
    DB_RS!GrhIndex = .GrhTxt
    DB_RS!UseGrh = .ProjecTxt
    DB_RS!UseSfx = .FXTxt
    DB_RS!ProjectileRotateSpeed = .RotTxt
    DB_RS!Stacking = .StackTxt


    End With
    'Update the database
    DB_RS.Update
    
    'Close the recordset
    DB_RS.Close
    MsgBox "Object " & ObjIndex & " successfully saved!", vbOKOnly
    
    frmMain.SelectObjCombo.Clear
    Engine_Load_ObjCombo
End Sub

Sub Engine_Render_Grh(ByRef Grh As Grh, ByVal x As Long, ByVal Y As Long, ByVal Center As Byte, ByVal Animate As Byte, Optional ByVal LoopAnim As Boolean = True, Optional ByVal Light1 As Long = -1, Optional ByVal Light2 As Long = -1, Optional ByVal Light3 As Long = -1, Optional ByVal Light4 As Long = -1, Optional ByVal Degrees As Byte = 0, Optional ByVal Shadow As Byte = 0)

'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************

Dim CurrentGrh As Grh

'Check to make sure it is legal

    If Grh.GrhIndex < 1 Then Exit Sub
    If GrhData(Grh.GrhIndex).NumFrames < 1 Then Exit Sub

    'Update the animation frame
    If Animate Then
        Grh.FrameCounter = Grh.FrameCounter + (0.0375 * GrhData(Grh.GrhIndex).Speed)
        If Grh.FrameCounter >= GrhData(Grh.GrhIndex).NumFrames + 1 Then
            Do While Grh.FrameCounter >= GrhData(Grh.GrhIndex).NumFrames + 1
                Grh.FrameCounter = Grh.FrameCounter - GrhData(Grh.GrhIndex).NumFrames
            Loop
            If LoopAnim <> True Then Grh.Started = 0
        End If
    End If

    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Int(Grh.FrameCounter))

    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * TilePixelWidth * 0.5) + TilePixelWidth * 0.5
        End If
        If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If

    'Draw
    Engine_Render_Rectangle x, Y, GrhData(CurrentGrh.GrhIndex).pixelWidth, GrhData(CurrentGrh.GrhIndex).pixelHeight, GrhData(CurrentGrh.GrhIndex).SX, GrhData(CurrentGrh.GrhIndex).SY, GrhData(CurrentGrh.GrhIndex).pixelWidth, GrhData(CurrentGrh.GrhIndex).pixelHeight, , , 0, GrhData(CurrentGrh.GrhIndex).FileNum, Light1, Light2, Light3, Light4, Shadow

End Sub

Sub Engine_Render_Rectangle(ByVal x As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal SrcX As Single, ByVal SrcY As Single, ByVal SrcWidth As Single, ByVal SrcHeight As Single, Optional ByVal SrcBitmapWidth As Long = -1, Optional ByVal SrcBitmapHeight As Long = -1, Optional ByVal Degrees As Single = 0, Optional ByVal TextureNum As Long, Optional ByVal Color0 As Long = -1, Optional ByVal Color1 As Long = -1, Optional ByVal Color2 As Long = -1, Optional ByVal Color3 As Long = -1, Optional ByVal Shadow As Byte = 0)

'************************************************************
'Render a square/rectangle based on the specified values then rotate it if needed
'************************************************************

Dim RadAngle As Single 'The angle in Radians
Dim CenterX As Single
Dim CenterY As Single
Dim Index As Integer
Dim NewX As Single
Dim NewY As Single
Dim SinRad As Single
Dim CosRad As Single

'Load the surface into memory if it is not in memory and reset the timer

    If TextureNum > 0 Then
        If SurfaceTimer(TextureNum) = 0 Then Engine_Init_Texture TextureNum
        SurfaceTimer(TextureNum) = SurfaceTimerMax
    End If

    'Set the texture
    If LastTexture <> TextureNum Then
        If TextureNum <= 0 Then
            D3DDevice.SetTexture 0, Nothing
        Else
            D3DDevice.SetTexture 0, SurfaceDB(TextureNum)
        End If
        LastTexture = TextureNum
    End If

    'Set the bitmap dimensions if needed
    If SrcBitmapWidth = -1 Then SrcBitmapWidth = SurfaceSize(TextureNum).x
    If SrcBitmapHeight = -1 Then SrcBitmapHeight = SurfaceSize(TextureNum).Y

    'Set shadowed settings - shadows only change on the top 2 points
    If Shadow Then

        SrcWidth = SrcWidth - 1

        'Set the top-left corner
        With VertexArray(0)
            .x = x + (Width * 0.5)
            .Y = Y - (Height * 0.5)
        End With

        'Set the top-right corner
        With VertexArray(1)
            .x = x + Width + (Width * 0.5)
            .Y = Y - (Height * 0.5)
        End With

    Else

        SrcWidth = SrcWidth + 1
        SrcHeight = SrcHeight + 1

        'Set the top-left corner
        With VertexArray(0)
            .x = x
            .Y = Y
        End With

        'Set the top-right corner
        With VertexArray(1)
            .x = x + Width
            .Y = Y
        End With

    End If

    'Set the top-left corner
    With VertexArray(0)
        .Color = Color0
        .Tu = SrcX / SrcBitmapWidth
        .Tv = SrcY / SrcBitmapHeight
    End With

    'Set the top-right corner
    With VertexArray(1)
        .Color = Color1
        .Tu = (SrcX + SrcWidth) / SrcBitmapWidth
        .Tv = SrcY / SrcBitmapHeight
    End With

    'Set the bottom-left corner
    With VertexArray(2)
        .x = x
        .Y = Y + Height
        .Color = Color2
        .Tu = SrcX / SrcBitmapWidth
        .Tv = (SrcY + SrcHeight) / SrcBitmapHeight
    End With

    'Set the bottom-right corner
    With VertexArray(3)
        .x = x + Width
        .Y = Y + Height
        .Color = Color3
        .Tu = (SrcX + SrcWidth) / SrcBitmapWidth
        .Tv = (SrcY + SrcHeight) / SrcBitmapHeight
    End With

    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), Len(VertexArray(0))

End Sub

Sub Main()
Dim i As Long

    InitFilePaths
    
    'Use the vbgore root directories if in the 3rd party tools folder
    i = InStr(1, UCase$(DataPath), UCase$("\3rd Party Tools\Database Editors"))
    If i > 0 Then DataPath = Left$(DataPath, i) & Right$(DataPath, Len(DataPath) - i - Len("\3rd Party Tools\Database Editors"))
    i = InStr(1, UCase$(ServerDataPath), UCase$("\3rd Party Tools\Database Editors"))
    If i > 0 Then ServerDataPath = Left$(ServerDataPath, i) & Right$(ServerDataPath, Len(ServerDataPath) - i - Len("\3rd Party Tools\Database Editors"))
    i = InStr(1, UCase$(GrhPath), UCase$("\3rd Party Tools\Database Editors"))
    If i > 0 Then GrhPath = Left$(GrhPath, i) & Right$(GrhPath, Len(GrhPath) - i - Len("\3rd Party Tools\Database Editors"))

    MySQL_Init
    
    frmMain.Show
    
    Engine_Init_TileEngine frmMain.PreviewPic.hWnd, frmMain.PreviewPic.ScaleWidth, frmMain.PreviewPic.ScaleHeight, 32, 32, 1, 0.011
    
    ''Load the object dropdown
    Engine_Load_ObjCombo
    
    ''Fill the object type list box
    Fill_ObjectType

End Sub

Sub Engine_Load_ObjCombo()
Dim NumObjDatas As Integer
Dim TempNum As Integer
Dim TempName As String

    'Retrieve the objects from the database
    DB_RS.Open "SELECT * FROM objects", DB_Conn, adOpenStatic, adLockOptimistic
    
    'Fill the object list. (Just id and name)
    Do While DB_RS.EOF = False  'Loop until we reach the end of the recordset
        TempNum = DB_RS!id
        TempName = DB_RS!Name
        frmMain.SelectObjCombo.AddItem TempNum & "- " & TempName
        DB_RS.MoveNext
    Loop

    'Close the recordset
    DB_RS.Close

End Sub

Sub Fill_ObjectType()
'Fill the types list.(Only so it's easy to read)
frmMain.ObjTypeCombo.AddItem "Select One"
frmMain.ObjTypeCombo.AddItem "Use Once"
frmMain.ObjTypeCombo.AddItem "Weapon"
frmMain.ObjTypeCombo.AddItem "Body"
frmMain.ObjTypeCombo.AddItem "Wings"


End Sub

Sub Engine_Init_Texture(ByVal TextureNum As Integer)

'*****************************************************************
'Loads a texture into memory
'*****************************************************************

Dim TexInfo As D3DXIMAGE_INFO_A
Dim FilePath As String

    'Get the path
    FilePath = GrhPath & TextureNum & ".png"

    'Make sure the texture exists

    If Engine_FileExist(FilePath, vbNormal) Then


        'Set the texture
        Set SurfaceDB(TextureNum) = D3DX.CreateTextureFromFileEx(D3DDevice, FilePath, D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, TexInfo, ByVal 0)

        'Set the size
        SurfaceSize(TextureNum).x = TexInfo.Width
        SurfaceSize(TextureNum).Y = TexInfo.Height

        'Set the texture timer
        SurfaceTimer(TextureNum) = SurfaceTimerMax

    End If

End Sub

Function Engine_Var_Get(File As String, Main As String, Var As String) As String

'*****************************************************************
'Gets a Var from a text file
'*****************************************************************

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

    szReturn = ""

    sSpaces = Space$(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish
    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File
    Engine_Var_Get = RTrim$(sSpaces)
    Engine_Var_Get = Left$(Engine_Var_Get, Len(Engine_Var_Get) - 1)

End Function

Sub Engine_Var_Write(File As String, Main As String, Var As String, Value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    writeprivateprofilestring Main, Var, Value, File

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

Private Function Engine_Init_D3DDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS)

'************************************************************
'Initialize the Direct3D Device - start off trying with the
'best settings and move to the worst until one works
'************************************************************

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
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.PreviewPic.hWnd, D3DCREATEFLAGS, D3DWindow)

    'Store the create flags
    UsedCreateFlags = D3DCREATEFLAGS

    'Everything was successful
    Engine_Init_D3DDevice = 1

    'The Rhw will always be 1, so set it now instead of every call
    For i = 0 To 3
        VertexArray(i).Rhw = 1
    Next i

Exit Function

ErrOut:

    'Destroy the D3DDevice so it can be remade
    Set D3DDevice = Nothing

    'Return a failure - 0
    Engine_Init_D3DDevice = 0

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
    Grh.FrameCounter = 1
    Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed

End Sub

Sub Engine_Init_GrhData()

'*****************************************************************
'Loads Grh.dat
'*****************************************************************

Dim Grh As Long
Dim Frame As Long

    'Get Number of Graphics
    NumGrhs = 40000 'CLng(Engine_Var_Get(DataPath & "Grh.ini", "INIT", "NumGrhs"))
    
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
            GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
            GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
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
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

    'Particle engine settings
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0

    'Set the texture stage stats (filters)
    '//D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    '//D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR

End Sub

Function Engine_Init_TileEngine(ByRef setDisplayFormhWnd As Long, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal Engine_Speed As Single) As Boolean

'*****************************************************************
'Init Tile Engine
'*****************************************************************

    'Set the array sizes by the number of graphic files
    NumGrhFiles = CInt(Engine_Var_Get(DataPath & "Grh.ini", "INIT", "NumGrhFiles"))
    ReDim SurfaceDB(1 To NumGrhFiles)
    ReDim SurfaceSize(1 To NumGrhFiles)
    ReDim SurfaceTimer(1 To NumGrhFiles)
    
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight

    '****** INIT DirectX ******
    ' Create the root D3D objects
    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate()
    Set D3DX = New D3DX8

    'Create the D3D Device
    If Engine_Init_D3DDevice(D3DCREATE_PUREDEVICE) = 0 Then
        If Engine_Init_D3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) = 0 Then
            If Engine_Init_D3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) = 0 Then
                MsgBox "Could not init D3DDevice. Exiting..."
                Engine_Init_UnloadTileEngine
                Engine_UnloadAllForms
                End
            End If
        End If
    End If
    Engine_Init_RenderStates

    'Load graphic data into memory
    Engine_Init_GrhData

    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60

    'Set high resolution timer
    timeBeginPeriod 1
    EndTime = timeGetTime

    'Start the engine
    Engine_Init_TileEngine = True
    EngineRun = True

End Function

Public Sub Engine_Init_UnloadTileEngine()

'*****************************************************************
'Shutsdown engine
'*****************************************************************

    On Error Resume Next

    Dim LoopC As Long

        EngineRun = False

        '****** Clear DirectX objects ******
        Set D3DDevice = Nothing
        Set D3DX = Nothing

        'Clear GRH memory
        For LoopC = 1 To NumGrhFiles
            Set SurfaceDB(LoopC) = Nothing
        Next LoopC

End Sub

Function Engine_FileExist(File As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

    Engine_FileExist = (Dir$(File, FileType) <> "")

End Function

Public Sub Server_Unload()
'*****************************************************************
'Unload the server and all the variables
'*****************************************************************
Engine_UnloadAllForms
End Sub

