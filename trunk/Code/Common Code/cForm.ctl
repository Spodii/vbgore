VERSION 5.00
Begin VB.UserControl cForm 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   ScaleHeight     =   2205
   ScaleWidth      =   3375
   Begin VB.PictureBox pic_RightCaption 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   1440
      Begin VB.Image img_MinimizeBtn 
         Height          =   300
         Left            =   810
         Top             =   0
         Width           =   285
      End
      Begin VB.Image img_MaximizeBtn 
         Height          =   300
         Left            =   540
         Top             =   0
         Width           =   285
      End
      Begin VB.Image img_RestoreBtn 
         Height          =   300
         Left            =   270
         Top             =   0
         Width           =   285
      End
      Begin VB.Image img_CloseBtn 
         Height          =   300
         Left            =   0
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.PictureBox pic_CenterCaption 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   480
      ScaleHeight     =   345
      ScaleWidth      =   855
      TabIndex        =   0
      Top             =   720
      Width           =   855
      Begin VB.Label lbl_Caption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   675
      End
   End
   Begin VB.PictureBox pic_LeftCaption 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   720
      TabIndex        =   6
      Top             =   0
      Width           =   720
   End
   Begin VB.PictureBox pic_DownBorder 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   0
      ScaleHeight     =   150
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.PictureBox pic_RightBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   150
      TabIndex        =   4
      Top             =   1560
      Width           =   150
   End
   Begin VB.PictureBox pic_Borders 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox pic_LeftBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   150
      TabIndex        =   2
      Top             =   1560
      Width           =   150
   End
End
Attribute VB_Name = "cForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Const DefMaximizeBtn = 1
Const DefMinimizeBtn = 1
Const DefCaption = "Caption"
Const DefBackColor = 0
Const DefForeColor = 0
Const DefCaptionTop = 195
Const DefCaptionColor = 0
Const DefAllowResizing = 1

Dim v_bMaximizeBtn As Boolean
Dim v_bMinimizeBtn As Boolean
Dim v_sCaption As String
Dim v_sSkinPath As String
Dim v_oBackColor As OLE_COLOR
Dim v_oForeColor As OLE_COLOR
Dim v_iCaptionTop As Integer
Dim v_oCaptionColor As OLE_COLOR
Dim v_bAllowResizing As Boolean
Dim v_iMouseX As Integer
Dim v_iMouseY As Integer
Dim v_bResizing As Boolean

Dim LBorderWidth As Long
Dim RBorderWidth As Long
Dim BBorderHeight As Long

Dim LBorderSourceX As Long
Dim RBorderSourceX As Long
Dim BBorderSourceX As Long
Dim BLBorderSourceX As Long
Dim BRBorderSourceX As Long

Dim ParentForm As Form

Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private Const SRCCOPY = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub LoadSkin(m_Form As Form)
Dim v_iCenterImgFrequency As Integer
Dim v_iLoop As Integer
Dim v_lhDC As Long
Dim v_lRtn As Long
Dim c As Control
Dim s As String

    Set ParentForm = m_Form
    
    s = Space$(255)
    GetPrivateProfileString "Form", "LBorderWidth", vbNullString, s, Len(s), SkinPath & "\Settings.ini"
    LBorderWidth = Val(s) * Screen.TwipsPerPixelX
     
    s = Space$(255)
    GetPrivateProfileString "Form", "RBorderWidth", vbNullString, s, Len(s), SkinPath & "\Settings.ini"
    RBorderWidth = Val(s) * Screen.TwipsPerPixelX
    
    s = Space$(255)
    GetPrivateProfileString "Form", "BBorderHeight", vbNullString, s, Len(s), SkinPath & "\Settings.ini"
    BBorderHeight = Val(s) * Screen.TwipsPerPixelY
    
    s = Space$(255)
    GetPrivateProfileString "Form", "LBorderSourceX", vbNullString, s, Len(s), SkinPath & "\Settings.ini"
    LBorderSourceX = Val(s)

    s = Space$(255)
    GetPrivateProfileString "Form", "RBorderSourceX", vbNullString, s, Len(s), SkinPath & "\Settings.ini"
    RBorderSourceX = Val(s)

    s = Space$(255)
    GetPrivateProfileString "Form", "BBorderSourceX", vbNullString, s, Len(s), SkinPath & "\Settings.ini"
    BBorderSourceX = Val(s)

    s = Space$(255)
    GetPrivateProfileString "Form", "BLBorderSourceX", vbNullString, s, Len(s), SkinPath & "\Settings.ini"
    BLBorderSourceX = Val(s)

    s = Space$(255)
    GetPrivateProfileString "Form", "BRBorderSourceX", vbNullString, s, Len(s), SkinPath & "\Settings.ini"
    BRBorderSourceX = Val(s)
    
    ParentForm.Height = ParentForm.Height + 145 + BBorderHeight
    ParentForm.Width = ParentForm.Width + ((LBorderWidth + RBorderWidth) / 2) - 45

    For Each c In ParentForm
        Select Case TypeName(c)
            Case "cForm"
                c.Left = 0
                c.Top = 0
            Case "Label"
                c.ForeColor = pSkin.LabelColor
            Case "TextBox"
                c.ForeColor = pSkin.TextColor
                c.BackColor = pSkin.TextBackColor
            Case "CheckBox"
                c.ForeColor = pSkin.LabelColor
                c.BackColor = pSkin.FormBackColor
            Case "OptionButton"
                c.ForeColor = pSkin.LabelColor
                c.BackColor = pSkin.FormBackColor
        End Select
    Next c
    
    With UserControl
        .Width = m_Form.Width
        .Height = m_Form.Height
        m_Form.BackColor = v_oBackColor
        m_Form.Caption = Caption

        .pic_LeftCaption.Picture = LoadPicture(SkinPath & "\img_Caption_Left.bmp")
        .pic_LeftCaption.Top = 0

        .pic_RightCaption.Picture = LoadPicture(SkinPath & "\img_Caption_Right.bmp")
        .pic_RightCaption.Left = .Width - .pic_RightCaption.Width
        
        .pic_CenterCaption.Picture = LoadPicture(SkinPath & "\img_Caption_Center.bmp")
        .pic_CenterCaption.Left = .pic_LeftCaption.Width

        .pic_CenterCaption.Height = 25 * Screen.TwipsPerPixelY

        .lbl_Caption.Width = .pic_CenterCaption.Width
                        
        .img_CloseBtn.Picture = LoadPicture(SkinPath & "\img_Button_Close.gif")
        .img_CloseBtn.Left = .pic_RightCaption.Width - .img_CloseBtn.Width - 75
        .img_CloseBtn.Top = 45
    
        .img_RestoreBtn.Picture = LoadPicture(SkinPath & "\img_Button_Restore.gif")
        .img_RestoreBtn.Left = .pic_RightCaption.Width - .img_RestoreBtn.Width - .img_CloseBtn.Width - 75
        .img_RestoreBtn.Top = 45
    
        .img_MaximizeBtn.Picture = LoadPicture(SkinPath & "\img_Button_Maximize.gif")
        .img_MaximizeBtn.Left = .pic_RightCaption.Width - .img_MaximizeBtn.Width - .img_CloseBtn.Width - 75
        .img_MaximizeBtn.Top = 45
    
        .img_MinimizeBtn.Picture = LoadPicture(SkinPath & "\img_Button_Minimize.gif")
        .img_MinimizeBtn.Left = .pic_RightCaption.Width - .img_MinimizeBtn.Width - .img_MaximizeBtn.Width - .img_CloseBtn.Width - 75
        .img_MinimizeBtn.Top = 45
    
        .pic_Borders.Picture = LoadPicture(SkinPath & "\img_Borders.bmp")

        .lbl_Caption.Top = CaptionTop
        .lbl_Caption.ForeColor = CaptionColor

    End With
    
    'Set the control to the top-left corner
    For Each c In ParentForm
        Select Case TypeName(c)
            Case "cForm"
            Case "CommonDialog"
            Case "Timer"
            Case Else
                c.Top = c.Top + 13
                c.Left = c.Left + 3
        End Select
    Next c
    Set c = Nothing
    
    'Force refresh
    Refresh
    
End Sub

Public Sub Refresh()
Dim v_iCenterImgFrequency As Integer
Dim v_iLoop As Integer
Dim v_lhDC As Long
Dim v_lRtn As Long

    With UserControl
        .Width = ParentForm.Width
        .Height = ParentForm.Height
        ParentForm.BackColor = v_oBackColor

        .pic_LeftCaption.Refresh
        .pic_LeftCaption.Top = 0

        .pic_RightCaption.Refresh
        .pic_RightCaption.Left = .Width - .pic_RightCaption.Width
        
        .pic_CenterCaption.Left = .pic_LeftCaption.Width
        .pic_CenterCaption.Width = .Width - .pic_LeftCaption.Width - .pic_RightCaption.Width
        .pic_CenterCaption.Refresh
        .lbl_Caption.Left = 60
        .lbl_Caption.Width = .pic_CenterCaption.Width
        .lbl_Caption.Refresh
        
        v_iCenterImgFrequency = Abs((.pic_CenterCaption.Width / Screen.TwipsPerPixelX) / 50)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_CenterCaption.hDC, v_iLoop * 50, 0, 50, 25, .pic_CenterCaption.hDC, 0, 0, SRCCOPY)
            Next v_iLoop
        End If
        .lbl_Caption.Width = .pic_CenterCaption.Width

        .img_CloseBtn.Left = .pic_RightCaption.Width - .img_CloseBtn.Width - 5
        .img_CloseBtn.Top = 5

        .img_RestoreBtn.Left = .pic_RightCaption.Width - .img_RestoreBtn.Width - .img_CloseBtn.Width - 5
        .img_RestoreBtn.Top = 5
   
        .img_MaximizeBtn.Left = .pic_RightCaption.Width - .img_MaximizeBtn.Width - .img_CloseBtn.Width - 5
        .img_MaximizeBtn.Top = 5
        
        .img_MinimizeBtn.Left = .pic_RightCaption.Width - .img_MaximizeBtn.Width - .img_CloseBtn.Width - 5
        If .img_MaximizeBtn.Visible = True Or .img_RestoreBtn.Visible = True Then .img_MinimizeBtn.Left = .img_MinimizeBtn.Left - .img_MinimizeBtn.Width
        .img_MinimizeBtn.Top = 5

        .pic_LeftBorder.Top = .pic_LeftCaption.Height
        .pic_LeftBorder.Height = .Height - .pic_LeftCaption.Height
        .pic_LeftBorder.Width = LBorderWidth
        .pic_RightBorder.Refresh
        .pic_RightBorder.Width = RBorderWidth
        .pic_RightBorder.Left = .Width - RBorderWidth
        .pic_RightBorder.Top = .pic_RightCaption.Height
        .pic_RightBorder.Height = ParentForm.Height - .pic_RightCaption.Height
        v_iCenterImgFrequency = Abs(((ParentForm.Height - .pic_LeftCaption.Height) / Screen.TwipsPerPixelY) / 5)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 0 To v_iCenterImgFrequency - 1
                BitBlt .pic_LeftBorder.hDC, 0, v_iLoop * 5, LBorderWidth, 5, .pic_Borders.hDC, LBorderSourceX, 0, SRCCOPY
                BitBlt .pic_RightBorder.hDC, 0, v_iLoop * 5, RBorderWidth, 5, .pic_Borders.hDC, RBorderSourceX, 0, SRCCOPY
            Next v_iLoop
        End If
        .pic_LeftBorder.Refresh
        .pic_RightBorder.Refresh

        .pic_DownBorder.Left = 0
        .pic_DownBorder.Top = ParentForm.Height - BBorderHeight
        .pic_DownBorder.Width = ParentForm.Width
        .pic_DownBorder.Height = BBorderHeight
        v_iCenterImgFrequency = Abs((ParentForm.Width / Screen.TwipsPerPixelX) / 4)

        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 0 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_DownBorder.hDC, v_iLoop * 4, 0, 4, 5, .pic_Borders.hDC, BBorderSourceX, 0, SRCCOPY)
            Next v_iLoop
        End If
        
        BitBlt .pic_DownBorder.hDC, 0, 0, 5, BBorderHeight, .pic_Borders.hDC, BLBorderSourceX, 0, SRCCOPY
        BitBlt .pic_DownBorder.hDC, (ParentForm.Width / Screen.TwipsPerPixelX) - 5, 0, 5, BBorderHeight, .pic_Borders.hDC, BRBorderSourceX, 0, SRCCOPY
        .pic_DownBorder.Refresh
        
        .lbl_Caption.Top = CaptionTop
        .lbl_Caption.ForeColor = CaptionColor
        
    End With
    
End Sub

Public Property Get MaximizeBtn() As Boolean
Attribute MaximizeBtn.VB_ProcData.VB_Invoke_Property = "ppg_SFCustom"

    MaximizeBtn = v_bMaximizeBtn
    
End Property

Public Property Let MaximizeBtn(ByVal m_MaximizeBtn As Boolean)

    v_bMaximizeBtn = m_MaximizeBtn
    PropertyChanged "Maximize"
    
End Property

Public Property Get MinimizeBtn() As Boolean
Attribute MinimizeBtn.VB_ProcData.VB_Invoke_Property = "ppg_SFCustom"

    MinimizeBtn = v_bMinimizeBtn
    
End Property

Public Property Let MinimizeBtn(ByVal m_MinimizeBtn As Boolean)

    v_bMinimizeBtn = m_MinimizeBtn
    PropertyChanged "Minimize"
    
End Property

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "ppg_SFCustom"

    Caption = v_sCaption
    
End Property

Public Property Let Caption(ByVal m_Caption As String)

    v_sCaption = m_Caption
    PropertyChanged "Caption"
    
End Property

Public Property Get SkinPath() As String

    SkinPath = v_sSkinPath
    
End Property

Public Property Let SkinPath(ByVal m_SkinPath As String)

    v_sSkinPath = m_SkinPath
    PropertyChanged "SkinPath"
    
End Property

Public Property Get BackColor() As OLE_COLOR

    BackColor = v_oBackColor
    
End Property

Public Property Let BackColor(ByVal m_BackColor As OLE_COLOR)

    v_oBackColor = m_BackColor
    PropertyChanged "BackColor"
    
End Property

Public Property Get ForeColor() As OLE_COLOR

    ForeColor = v_oForeColor
    
End Property

Public Property Let ForeColor(ByVal m_ForeColor As OLE_COLOR)

    v_oForeColor = m_ForeColor
    PropertyChanged "ForeColor"
    
End Property

Public Property Get CaptionTop() As Integer
Attribute CaptionTop.VB_ProcData.VB_Invoke_Property = "ppg_SFCustom"

    CaptionTop = v_iCaptionTop
    
End Property

Public Property Let CaptionTop(ByVal m_CaptionTop As Integer)

    v_iCaptionTop = m_CaptionTop
    PropertyChanged "CaptionTop"
    
End Property

Public Property Get CaptionColor() As OLE_COLOR

    CaptionColor = v_oCaptionColor
    
End Property

Public Property Let CaptionColor(ByVal m_CaptionColor As OLE_COLOR)

    v_oCaptionColor = m_CaptionColor
    PropertyChanged "CaptionColor"
    
End Property

Public Property Get AllowResizing() As Boolean
Attribute AllowResizing.VB_ProcData.VB_Invoke_Property = "ppg_SFCustom"

    AllowResizing = v_bAllowResizing
    
End Property

Public Property Let AllowResizing(ByVal m_AllowResizing As Boolean)

    v_bAllowResizing = m_AllowResizing
    PropertyChanged "AllowResizing"
    
End Property

Private Sub img_CloseBtn_Click()

    Unload Screen.ActiveForm
    
End Sub

Private Sub img_CloseBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    UserControl.MousePointer = 0
    
End Sub

Private Sub img_MaximizeBtn_Click()

    Screen.ActiveForm.WindowState = 2
    UserControl.img_MaximizeBtn.Visible = False
    UserControl.img_RestoreBtn.Visible = True
    Call Refresh
    
End Sub

Private Sub img_MaximizeBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    UserControl.MousePointer = 0
    
End Sub

Private Sub img_MinimizeBtn_Click()

    Screen.ActiveForm.WindowState = 1
    
End Sub

Private Sub img_MinimizeBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    UserControl.MousePointer = 0
    
End Sub

Private Sub img_RestoreBtn_Click()

    Screen.ActiveForm.WindowState = 0
    UserControl.img_MaximizeBtn.Visible = True
    UserControl.img_RestoreBtn.Visible = False
    Call LoadSkin(ParentForm)
    
End Sub

Private Sub img_RestoreBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    UserControl.MousePointer = 0
    
End Sub

Private Sub lbl_Caption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        v_iMouseX = X
        v_iMouseY = Y
    End If
    
End Sub

Private Sub lbl_Caption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        If ParentForm.WindowState <> 2 Then
            Screen.ActiveForm.Left = Screen.ActiveForm.Left + X - v_iMouseX
            Screen.ActiveForm.Top = Screen.ActiveForm.Top + Y - v_iMouseY
        End If
    End If
    
End Sub

Private Sub lbl_Caption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Screen.ActiveForm.MousePointer = 0
    
End Sub

Private Sub pic_CenterCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        If v_bAllowResizing Then
            v_iMouseX = X
            v_iMouseY = Y
            If Y <= 120 Then v_bResizing = True
        End If
    End If
    
End Sub

Private Sub pic_CenterCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        If Not v_bResizing = False Then
            If ParentForm.WindowState <> 2 Then
                Screen.ActiveForm.Left = Screen.ActiveForm.Left + X - v_iMouseX
                Screen.ActiveForm.Top = Screen.ActiveForm.Top + Y - v_iMouseY
            End If
        End If
    End If
    
    If Y <= 120 Then
        If v_bAllowResizing Then
            UserControl.MousePointer = 7
            If Button = 1 Then
                If v_bResizing Then
                    ParentForm.Top = ParentForm.Top + Y - v_iMouseY
                    ParentForm.Height = ParentForm.Height - Y + v_iMouseY
                    Call Refresh
                End If
            End If
        End If
    Else
        UserControl.MousePointer = 0
    End If
    
End Sub

Private Sub pic_DownBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        If v_bAllowResizing Then
            v_iMouseX = X
            v_iMouseY = Y
            v_bResizing = True
        End If
    End If
    
End Sub

Private Sub pic_DownBorder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If v_bAllowResizing Then
        UserControl.MousePointer = 7
        If Button = 1 Then
            If v_bResizing Then
                ParentForm.Height = ParentForm.Height + Y - v_iMouseY
                Call Refresh
            End If
        End If
    End If
    
End Sub

Private Sub pic_DownBorder_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    v_bResizing = False
    
End Sub

Private Sub pic_LeftBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        If v_bAllowResizing Then
            v_iMouseX = X
            v_iMouseY = Y
            v_bResizing = True
        End If
    End If
    
End Sub

Private Sub pic_LeftBorder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If v_bAllowResizing Then UserControl.MousePointer = 9
    
    If Button = 1 Then
        If v_bResizing Then
            ParentForm.Left = ParentForm.Left + X - v_iMouseX
            ParentForm.Width = ParentForm.Width - X + v_iMouseX
            Call Refresh
        End If
    End If
    
End Sub

Private Sub pic_LeftBorder_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    v_bResizing = False
    
End Sub

Private Sub pic_RightBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        If v_bAllowResizing Then
            v_iMouseX = X
            v_iMouseY = Y
            v_bResizing = True
        End If
    End If
    
End Sub

Private Sub pic_RightBorder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If v_bAllowResizing Then UserControl.MousePointer = 9
    
    If Button = 1 Then
        If v_bResizing Then
            ParentForm.Width = ParentForm.Width + X - v_iMouseX
            Call Refresh
        End If
    End If
    
End Sub

Private Sub pic_RightBorder_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    v_bResizing = False
    
End Sub

Private Sub UserControl_Click()

    RaiseEvent Click
    
End Sub

Private Sub UserControl_Initialize()

    'Set the start positions
    pic_LeftCaption.Left = 0
    pic_LeftCaption.Top = 0
    pic_CenterCaption.Left = 0
    pic_CenterCaption.Top = 0
    
End Sub

Private Sub UserControl_InitProperties()

    'Default settings
    v_bMaximizeBtn = DefMaximizeBtn
    v_bMinimizeBtn = DefMinimizeBtn
    v_sCaption = DefCaption
    v_sSkinPath = App.Path & "\FormSkins\" & Skin_GetCurrent
    v_oBackColor = DefBackColor
    v_oForeColor = DefForeColor
    v_oCaptionColor = DefCaptionColor
    v_bAllowResizing = DefAllowResizing
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseMove(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseUp(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    v_bMaximizeBtn = PropBag.ReadProperty("MaximizeBtn", DefMaximizeBtn)
    
    UserControl.img_MaximizeBtn.Visible = v_bMaximizeBtn
    UserControl.img_RestoreBtn.Visible = v_bMaximizeBtn
    
    If v_bMaximizeBtn Then
        UserControl.img_MinimizeBtn.Left = UserControl.pic_RightCaption.Width - UserControl.img_MinimizeBtn.Width - UserControl.img_CloseBtn.Width - UserControl.img_MaximizeBtn.Width - 75
    Else
        UserControl.img_MinimizeBtn.Left = UserControl.pic_RightCaption.Width - UserControl.img_MinimizeBtn.Width - UserControl.img_CloseBtn.Width - 75
    End If

    v_bMinimizeBtn = PropBag.ReadProperty("MinimizeBtn", DefMinimizeBtn)
    UserControl.img_MinimizeBtn.Visible = v_bMinimizeBtn
    
    v_sCaption = PropBag.ReadProperty("Caption", DefCaption)
    UserControl.lbl_Caption.Caption = v_sCaption
    
    v_sSkinPath = PropBag.ReadProperty("SkinPath", App.Path & "\FormSkins\" & Skin_GetCurrent)
    v_oBackColor = PropBag.ReadProperty("BackColor", DefBackColor)
    
    v_oForeColor = PropBag.ReadProperty("ForeColor", DefForeColor)
    UserControl.lbl_Caption.ForeColor = v_oForeColor
    
    v_iCaptionTop = PropBag.ReadProperty("CaptionTop", DefCaptionTop)
    UserControl.lbl_Caption.Top = v_iCaptionTop

    v_oCaptionColor = PropBag.ReadProperty("CaptionColor", DefCaptionColor)
    UserControl.lbl_Caption.ForeColor = v_oCaptionColor

    v_bAllowResizing = PropBag.ReadProperty("AllowResizing", DefAllowResizing)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("MaximizeBtn", v_bMaximizeBtn, DefMaximizeBtn)
    Call PropBag.WriteProperty("MinimizeBtn", v_bMinimizeBtn, DefMinimizeBtn)
    Call PropBag.WriteProperty("Caption", v_sCaption, DefCaption)
    Call PropBag.WriteProperty("SkinPath", v_sSkinPath, App.Path & "\FormSkins\" & Skin_GetCurrent)
    Call PropBag.WriteProperty("BackColor", v_oBackColor, DefBackColor)
    Call PropBag.WriteProperty("ForeColor", v_oForeColor, DefForeColor)
    Call PropBag.WriteProperty("CaptionTop", v_iCaptionTop, DefCaptionTop)
    Call PropBag.WriteProperty("CaptionColor", v_oCaptionColor, DefCaptionColor)
    Call PropBag.WriteProperty("AllowResizing", v_bAllowResizing, DefAllowResizing)
 
End Sub
