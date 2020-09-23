Attribute VB_Name = "modNoFocusRect"
'API Declarations
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Consts
Private Const GWL_WNDPROC = (-4)
Private Const WM_ACTIVATE = &H6
Private Const WM_SETFOCUS = &H7

'Vars
Public StandardButtonProc As Long

Public Sub NoFocusRect(Button As Object, vValue As Boolean)
    If vValue = True Then 'Focus rect on
        'Save the adress of the standard button procedure
        StandardButtonProc = GetWindowLong(Button.hWnd, GWL_WNDPROC)
        'Subclass the button to control its Windows Messages
        SetWindowLong Button.hWnd, GWL_WNDPROC, AddressOf ButtonProc
    Else 'Focus rect off
        'Remove the subclassing from the button
        SetWindowLong Button.hWnd, GWL_WNDPROC, StandardButtonProc
    End If
End Sub

Public Function ButtonProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'The procedure that gets all windows messages for the subclassed
    'button
    On Error Resume Next
    Select Case uMsg&
        'The button is going to get the focus
        Case WM_SETFOCUS
        'Exit the procedure -> The message doesn´t reach the button
        Exit Function
    End Select
    'Call the standard Button Procedure
    ButtonProc = CallWindowProc(StandardButtonProc, hWnd&, uMsg&, wParam&, lParam&)
End Function
