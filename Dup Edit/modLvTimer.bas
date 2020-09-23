Attribute VB_Name = "modLvTimer"
Option Explicit
' REQUIRED: copy & paste these few lines in any module of your project
' This is used by every lvButtons control as a replacement of the Timer control
' By doing it this way, each button control does NOT need an individual timer control
' The timer function is primarily used to determine when the mouse enters/leaves
' the button's physical region on the screen. See compiling information notes below.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSource As Any, _
                                                                     ByVal ByteLen As Long)
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, _
                                                                ByVal lpString As String) As Long

Public Function lv_TimerCallBack(ByVal hwnd As Long, _
                                 ByVal Message As Long, _
                                 ByVal wParam As Long, _
                                 ByVal lParam As Long) As Long
                                 
  Dim tgtButton As lvButtons_H

    ' when timer was intialized, the button control's hWnd
    ' had property set to the handle of the control itself
    ' and the timer ID was also set as a window property
    CopyMemory tgtButton, GetProp(hwnd, "lv_ClassID"), &H4
    Call tgtButton.TimerUpdate(GetProp(hwnd, "lv_TimerID"))  ' fire the button's event
    CopyMemory tgtButton, 0&, &H4                                    ' erase this instance

End Function

':) Roja's VB Code Fixer V1.0.97 (6/20/03 2:45:35 AM) 10 + 18 = 28 Lines Thanks Ulli for inspiration and lots of code.

