Attribute VB_Name = "modApi"
Type POINTAPI
    X As Long
    Y As Long
End Type

Public Const SW_SHOWNOACTIVATE = 4

Public old_HWND                         As Long
Public new_HWND                         As Long
Public FColor                           As Long

Declare Function GetActiveWindow Lib "User32" () As Long
Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Sub PopDownAll()
Dim ctl As Control
Dim ctl_hWnd As Long
Dim txt As String
    ' Find the control with the given hWnd.
    On Error Resume Next
    For Each ctl In Parent.Controls
            If ctl.Tag = "AutoButton" Then
             ctl.PopDown
            End If
    Next ctl

End Sub

