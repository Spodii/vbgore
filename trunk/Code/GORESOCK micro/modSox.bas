Attribute VB_Name = "modSocket"
Option Explicit

Public SocketControl As Socket

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Let WindowProc = SocketControl.WndProc(hWnd, uMsg, wParam, lParam)

End Function
