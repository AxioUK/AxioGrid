Attribute VB_Name = "mouseScroll"
Option Explicit

' declaraciones del api
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    wParam As Any, lParam As Any) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Constantes
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const GWL_WNDPROC = (-4)
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_VSCROLL As Integer = &H115

Dim PrevProc As Long

Public Sub HookScroll(Obj As Object)
    PrevProc = SetWindowLong(Obj.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHookScroll(Obj As Object)
    SetWindowLong Obj.hwnd, GWL_WNDPROC, PrevProc
End Sub

' Procedimiento qie intercepta los mensajes de windows, en este caso para _
  interceptar el uso del Scroll del mouse
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function WindowProc(ByVal hwnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

    WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
    
    If uMsg = WM_MOUSEWHEEL Then
        If wParam < 0 Then
            ' envia mediante SendMessage el comando para mover el Scroll hacia abajo
            SendMessage hwnd, WM_VSCROLL, ByVal 1, ByVal 0
        Else
            ' Mueve el scroll hacia arriba
            SendMessage hwnd, WM_VSCROLL, ByVal 0, ByVal 0
        End If
    End If
End Function

