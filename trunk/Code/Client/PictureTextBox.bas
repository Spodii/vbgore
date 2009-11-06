Attribute VB_Name = "PictureTextBox"
Option Explicit

'Notice - Text boxes must be multiline for this to work!
'I know this isn't the best way to go about doing this, but it isn't
'used for very long nor is it used in any other projects, so no point in wasting time
'making it very versitile

'Holds the returns from SetWindowLong
Private frmNewPrev As Long
Private frmConnectPrev As Long

'APIs we will be using
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDC As Long, ByVal DX As Long, ByVal DY As Long, ByVal DWidth As Long, ByVal DHeight As Long, ByVal ShDC As Long, ByVal SX As Long, ByVal SY As Long, ByVal vbSrCopy As Long) As Long
Private Declare Function SetBkMode Lib "GDI32" (ByVal hDC As Long, ByVal hMode As Long) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hDC As Long) As Long

Public Sub SetPictureTextboxes(ByVal hwnd As Long)

    'Set the form to subclass and the textbox heights
    Select Case hwnd
    
    Case frmConnect.hwnd
        frmConnectPrev = SetWindowLong(hwnd, -4, AddressOf frmConnectProc)
        With frmConnect
            .NameTxt.Height = Int(.NameTxt.Height \ .TextHeight("_")) * .TextHeight("_")
            .PasswordTxt.Height = Int(.PasswordTxt.Height \ .TextHeight("_")) * .TextHeight("_")
        End With
        
    Case frmNew.hwnd
        frmNewPrev = SetWindowLong(hwnd, -4, AddressOf frmNewProc)
        With frmNew
            .NameTxt.Height = Int(.NameTxt.Height \ .TextHeight("_")) * .TextHeight("_")
            .PasswordTxt.Height = Int(.PasswordTxt.Height \ .TextHeight("_")) * .TextHeight("_")
        End With
        
    End Select
    
End Sub

Public Sub FreePictureTextboxes(ByVal hwnd As Long)
    
    'Free the form
    Select Case hwnd
        
    Case frmConnect.hwnd
        SetWindowLong hwnd, -4, frmConnectPrev
        
    Case frmNew.hwnd
        SetWindowLong hwnd, -4, frmNewPrev
    
    End Select
    
End Sub

Private Function frmNewProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    'Check for a message we want
    If uMsg = &H133 Then
    
        'Make sure our form is visible
        If frmNew.Visible Then
        
            'Look for the hWnds we want and handle accordingly
            Select Case WindowFromDC(wParam)
            
            Case frmNew.PasswordTxt.hwnd
                With frmNew.PasswordTxt
                    SetBkMode wParam, 1
                    BitBlt wParam, 0, 0, .Width, .Height, frmNew.hDC, .Left, .Top, vbSrcCopy
                End With
    
            Case frmNew.NameTxt.hwnd
                With frmNew.NameTxt
                    SetBkMode wParam, 1
                    BitBlt wParam, 0, 0, .Width, .Height, frmNew.hDC, .Left, .Top, vbSrcCopy
                End With
        
            Case frmNew.ClassCmb.hwnd
                With frmNew.ClassCmb
                    SetBkMode wParam, 1
                    BitBlt wParam, 0, 0, .Width, .Height, frmNew.hDC, .Left, .Top, vbSrcCopy
                End With
                
            Case frmNew.BodyCmb.hwnd
                With frmNew.BodyCmb
                    SetBkMode wParam, 1
                    BitBlt wParam, 0, 0, .Width, .Height, frmNew.hDC, .Left, .Top, vbSrcCopy
                End With
                
            Case frmNew.HeadCmb.hwnd
                With frmNew.HeadCmb
                    SetBkMode wParam, 1
                    BitBlt wParam, 0, 0, .Width, .Height, frmNew.hDC, .Left, .Top, vbSrcCopy
                End With
                    
            End Select
            
        End If
        
    End If
    
    'Send the message to the form
    frmNewProc = CallWindowProc(frmNewPrev, hwnd, uMsg, wParam, lParam)
    
End Function

Private Function frmConnectProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    'Check for a message we want
    If uMsg = &H133 Then
    
        'Make sure our form is visible
        If frmConnect.Visible Then
        
            'Look for the hWnds we want and handle accordingly
            Select Case WindowFromDC(wParam)
            
            Case frmConnect.PasswordTxt.hwnd
                With frmConnect.PasswordTxt
                    SetBkMode wParam, 1
                    BitBlt wParam, 0, 0, .Width, .Height, frmConnect.hDC, .Left, .Top, vbSrcCopy
                End With
    
            Case frmConnect.NameTxt.hwnd
                With frmConnect.NameTxt
                    SetBkMode wParam, 1
                    BitBlt wParam, 0, 0, .Width, .Height, frmConnect.hDC, .Left, .Top, vbSrcCopy
                End With
    
            End Select
            
        End If
        
    End If
    
    'Send the message to the form
    frmConnectProc = CallWindowProc(frmConnectPrev, hwnd, uMsg, wParam, lParam)
    
End Function

