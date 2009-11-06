Attribute VB_Name = "GUISkin"
Option Explicit

Type TSkin
    SkinName As String
    FormBackColor As Long
    CaptionTop As Integer
    CaptionColor As Long
    SButtonForeColor As Long
    LabelColor As Long
    TextColor As Long
    TextBackColor As Long
End Type
Public pSkin As TSkin

Public GUISkinPath As String

Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Sub Skin_InitStructure(ByVal SkinName As String)
Dim RetStr As String
Dim FilePath As String

    FilePath = App.Path & "\FormSkins\" & SkinName & "\Settings.ini"
    
    RetStr = Space$(255)
    getprivateprofilestring "General", "SkinName", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.SkinName = Trim$(RetStr)
        
    RetStr = Space$(255)
    getprivateprofilestring "Form", "BackColor", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.FormBackColor = Val(Trim$(RetStr))
    
    RetStr = Space$(255)
    getprivateprofilestring "Form", "CaptionTop", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.CaptionTop = Val(Trim$(RetStr))
    
    RetStr = Space$(255)
    getprivateprofilestring "Form", "CaptionColor", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.CaptionColor = Val(Trim$(RetStr))

    RetStr = Space$(255)
    getprivateprofilestring "Button", "ForeColor", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.SButtonForeColor = Val(Trim$(RetStr))
    
    RetStr = Space$(255)
    getprivateprofilestring "General", "TextColor", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.TextColor = Val(Trim$(RetStr))

    RetStr = Space$(255)
    getprivateprofilestring "General", "LabelColor", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.LabelColor = Val(Trim$(RetStr))
    
    RetStr = Space$(255)
    getprivateprofilestring "General", "TextColor", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.TextColor = Val(Trim$(RetStr))
    
    RetStr = Space$(255)
    getprivateprofilestring "General", "TextBackColor", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.TextBackColor = Val(Trim$(RetStr))
    
End Sub

Private Sub Skin_Change(Targetform As Form)
Dim c As Control

    Skin_InitStructure Skin_GetCurrent
    
    For Each c In Targetform
    
        Select Case TypeName(c)
        
            Case "cForm"
                c.SkinPath = App.Path & "\FormSkins\" & Skin_GetCurrent
                c.BackColor = pSkin.FormBackColor
                c.CaptionTop = pSkin.CaptionTop
                c.CaptionColor = pSkin.CaptionColor
                Call c.LoadSkin(Targetform)
                
            Case "cButton"
                c.SkinPath = App.Path & "\FormSkins\" & Skin_GetCurrent
                c.ForeColor = pSkin.SButtonForeColor
                c.LoadSkin

        End Select
        
    Next c
    
    Set c = Nothing
    
End Sub

Public Sub Skin_Set(Targetform As Form)

    Skin_InitStructure Skin_GetCurrent
    Skin_Change Targetform
    
End Sub

Public Sub Skin_SetForm(Targetform As Form)
Dim c As Control
    
    Targetform.BackColor = pSkin.FormBackColor
    For Each c In Targetform
        Select Case UCase$(TypeName(c))
            Case "CFORM"
                c.Left = 0
                c.Top = 0
            Case "PICTUREBOX"
                c.BackColor = pSkin.FormBackColor
            Case "LABEL"
                c.ForeColor = pSkin.LabelColor
            Case "TEXTBOX"
                c.ForeColor = pSkin.TextColor
                c.BackColor = pSkin.TextBackColor
            Case "LISTBOX"
                c.ForeColor = pSkin.TextColor
                c.BackColor = pSkin.TextBackColor
            Case "CHECKBOX"
                c.ForeColor = pSkin.LabelColor
                c.BackColor = pSkin.FormBackColor
            Case "OPTIONBUTTON"
                c.ForeColor = pSkin.LabelColor
                c.BackColor = pSkin.FormBackColor
        End Select
    Next c

End Sub

Public Function Skin_GetCurrent() As String
Dim s As String

    s = Space$(255)
    getprivateprofilestring "INIT", "CurrentSkin", vbNullString, s, Len(s), App.Path & "\FormSkins\CurrentSkin.ini"
    Skin_GetCurrent = Trim$(s)
    Skin_GetCurrent = Left$(Skin_GetCurrent, Len(Skin_GetCurrent) - 1)

End Function
