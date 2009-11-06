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

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Sub Skin_InitStructure(ByVal SkinName As String)
Dim RetStr As String
Dim FilePath As String

    FilePath = App.Path & "\FormSkins\" & SkinName & "\Settings.ini"
    
    RetStr = Space$(255)
    GetPrivateProfileString "General", "SkinName", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.SkinName = Trim$(RetStr)
        
    RetStr = Space$(255)
    GetPrivateProfileString "Form", "BackColor", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.FormBackColor = Val(Trim$(RetStr))
    
    RetStr = Space$(255)
    GetPrivateProfileString "Form", "CaptionTop", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.CaptionTop = Val(Trim$(RetStr))
    
    RetStr = Space$(255)
    GetPrivateProfileString "Form", "CaptionColor", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.CaptionColor = Val(Trim$(RetStr))

    RetStr = Space$(255)
    GetPrivateProfileString "Button", "ForeColor", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.SButtonForeColor = Val(Trim$(RetStr))
    
    RetStr = Space$(255)
    GetPrivateProfileString "General", "TextColor", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.TextColor = Val(Trim$(RetStr))

    RetStr = Space$(255)
    GetPrivateProfileString "General", "LabelColor", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.LabelColor = Val(Trim$(RetStr))
    
    RetStr = Space$(255)
    GetPrivateProfileString "General", "TextColor", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.TextColor = Val(Trim$(RetStr))
    
    RetStr = Space$(255)
    GetPrivateProfileString "General", "TextBackColor", vbNullString, RetStr, Len(RetStr), FilePath
    pSkin.TextBackColor = Val(Trim$(RetStr))
    
End Sub

Private Sub Skin_Change(TargetForm As Form)
Dim c As Control

    Skin_InitStructure Skin_GetCurrent
    
    For Each c In TargetForm
    
        Select Case TypeName(c)
        
            Case "cForm"
                c.SkinPath = App.Path & "\FormSkins\" & Skin_GetCurrent
                c.BackColor = pSkin.FormBackColor
                c.CaptionTop = pSkin.CaptionTop
                c.CaptionColor = pSkin.CaptionColor
                Call c.LoadSkin(TargetForm)
                
            Case "cButton"
                c.SkinPath = App.Path & "\FormSkins\" & Skin_GetCurrent
                c.ForeColor = pSkin.SButtonForeColor
                c.LoadSkin

        End Select
        
    Next c
    
    Set c = Nothing
    
End Sub

Public Sub Skin_Set(TargetForm As Form)

    Skin_InitStructure Skin_GetCurrent
    Skin_Change TargetForm
    
End Sub

Public Function Skin_GetCurrent() As String
Dim s As String

    s = Space$(255)
    GetPrivateProfileString "INIT", "CurrentSkin", vbNullString, s, Len(s), App.Path & "\FormSkins\CurrentSkin.ini"
    Skin_GetCurrent = Trim$(s)
    Skin_GetCurrent = Left$(Skin_GetCurrent, Len(Skin_GetCurrent) - 1)

End Function
