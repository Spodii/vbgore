Attribute VB_Name = "SystemTray"
Option Explicit

'System tray variables
Private Const NIF_MESSAGE As Long = &H1 'Message
Private Const NIF_ICON As Long = &H2    'Icon
Private Const NIF_TIP As Long = &H4     'TooTipText
Private Const NIM_ADD As Long = &H0     'Add to tray
Private Const NIM_MODIFY As Long = &H1  'Modify
Private Const NIM_DELETE As Long = &H2  'Delete From Tray
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Enum TrayRetunEventEnum
    MouseMove = &H200       'On Mousemove
    LeftUp = &H202          'Left Button Mouse Up
    LeftDown = &H201        'Left Button MouseDown
    LeftDbClick = &H203     'Left Button Double Click
    RightUp = &H205         'Right Button Up
    RightDown = &H204       'Right Button Down
    RightDbClick = &H206    'Right Button Double Click
    MiddleUp = &H208        'Middle Button Up
    MiddleDown = &H207      'Middle Button Down
    MiddleDbClick = &H209   'Middle Button Double Click
End Enum
#If False Then
Private MouseMove, LeftUp, LeftDown, LeftDbClick, RightUp, RightDown, RightDbClick, MiddleUp, MiddleDown, MiddleDbClick
#End If
Public Enum ModifyItemEnum
    ToolTip = 1             'Modify ToolTip
    Icon = 2                'Modify Icon
End Enum
#If False Then
Private ToolTip, Icon
#End If
Private TrayIcon As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Sub TrayAdd(TargetForm As Form, ByVal ToolTip As String, ReturnCallEvent As TrayRetunEventEnum)

'Add to the tray

    With TrayIcon
        .cbSize = Len(TrayIcon)
        .hwnd = TargetForm.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = ReturnCallEvent
        .hIcon = TargetForm.Icon
        .szTip = ToolTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, TrayIcon

End Sub

Public Sub TrayDelete()

'Remove from the tray

    Shell_NotifyIcon NIM_DELETE, TrayIcon

End Sub

Public Sub TrayModify(Item As ModifyItemEnum, vNewValue As Variant)

'Modify the tray

    Select Case Item
    Case ToolTip
        TrayIcon.szTip = vNewValue & vbNullChar
    Case Icon
        TrayIcon.hIcon = vNewValue.Handle
    End Select
    Shell_NotifyIcon NIM_MODIFY, TrayIcon

End Sub
