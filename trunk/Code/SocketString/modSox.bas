Attribute VB_Name = "modSox"
Option Explicit

Public SoxControl As Sox ' Our Public reference to Sox, This will allow us to call Sox commands from anywhere in the project like Sox.SendData instead of frmMain.Sox.SendData

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Let WindowProc = SoxControl.WndProc(hWnd, uMsg, wParam, lParam)

End Function

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 18:20)  Decl: 3  Code: 9  Total: 12 Lines
':) CommentOnly: 0 (0%)  Commented: 1 (8.3%)  Empty: 3 (25%)  Max Logic Depth: 1
