Attribute VB_Name = "Manifest"
Option Explicit

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Const ICC_USEREX_CLASSES = &H200

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Public Sub InitManifest()
Dim iccex As tagInitCommonControlsEx

    'This routine will make sure the application hooks to ComCtrl32.DLL
    'This must be called on the very first line of sub Main (whichever is called first)

    On Error Resume Next

    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    
    InitCommonControlsEx iccex

    On Error GoTo 0
   
End Sub

