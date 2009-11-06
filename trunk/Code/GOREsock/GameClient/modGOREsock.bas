Attribute VB_Name = "modGOREsock"
Option Explicit

Public GOREsockClient As GOREsockClient

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Let WindowProc = GOREsockClient.WndProc(hwnd, uMsg, wParam, lParam)
    
End Function
