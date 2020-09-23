Attribute VB_Name = "Module1"
Option Explicit

'APIs needed for subclassing
Declare Function FindWindowEx& Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String)
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'This API is for getting the child of a parent
Declare Function GetParent& Lib "user32" (ByVal hWnd As Long)

Global ListViewHeader_hWnd As Long
Global lngOldProc As Long

Global Const GWL_WNDPROC = (-4)
Global Const WM_RBUTTONUP = &H205


Public Function SubClass(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'uMsg is the event that occurs
    'here we check if the event was a release on the right mousebutton:
    If uMsg = WM_RBUTTONUP Then
        'Now we can do desired actions ie. a popup like in windows ;)
        Form1.PopupMenu Form1.mnuPOP 'PS... if you remove the Form1.mnuPOP, and run the project...
                                     'apperantly, VB crashed or just exits :) (try it)
    Else
        'If the event was not a right mouse button release, then go on as usual
        'If we dont resume this, the control will not function normal, and only react on OUR classing
        SubClass = CallWindowProc(lngOldProc, hWnd, uMsg, wParam, lParam)
        Exit Function
    End If
End Function

Sub EnumListVewChild(Parent As Long)
Dim hWnd As Long

If Parent = 0 Then Exit Sub

 hWnd = FindWindowEx(Parent, hWnd, vbNullString, vbNullString)

 If hWnd <> 0 Then
    If GetParent(hWnd) = Parent Then
     ListViewHeader_hWnd = hWnd
    End If
 End If

End Sub
