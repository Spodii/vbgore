Attribute VB_Name = "Manifest"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Sub InitManifest()

    'This routine will make sure the application hooks to ComCtrl32.DLL
    'This must be called on the very first line of sub Main

    InitCommonControls
   
End Sub

