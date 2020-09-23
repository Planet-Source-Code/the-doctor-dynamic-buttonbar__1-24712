VERSION 5.00
Begin VB.UserControl LargeButton3D 
   BackColor       =   &H80000010&
   CanGetFocus     =   0   'False
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   ScaleHeight     =   2925
   ScaleWidth      =   4005
   Begin VB.Timer tmrLeft 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2880
      Top             =   720
   End
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   720
   End
   Begin VB.Line LineDown 
      BorderColor     =   &H80000015&
      Visible         =   0   'False
      X1              =   600
      X2              =   0
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line LineRight 
      BorderColor     =   &H80000015&
      Visible         =   0   'False
      X1              =   600
      X2              =   600
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line LineUp 
      BorderColor     =   &H80000014&
      Visible         =   0   'False
      X1              =   0
      X2              =   600
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line LineLeft 
      BorderColor     =   &H80000014&
      Visible         =   0   'False
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Image imgPIC 
      Height          =   495
      Left            =   2400
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgEPIC 
      Height          =   495
      Left            =   2400
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgDPIC 
      Height          =   495
      Left            =   2400
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "LargeButton3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Local variables for the properties
Private booIsDown                   As Boolean
Private booStayUp                   As Boolean
Private booIsPopUp                  As Boolean

'Variables
Public booIsClicked                 As Boolean

'Objects
Public ctl                          As Control

'Events
Event Click()

'***************************************************************************
'Properties
'***************************************************************************

Public Property Get IsDown() As Boolean
    IsDown = booIsDown
End Property

Public Property Let IsDown(booValue As Boolean)
    booIsDown = booValue
    PropertyChanged "IsDown"
End Property

Public Property Get IsPopup() As Boolean
    IsPopup = booIsPopUp
End Property

Public Property Let IsPopup(booValue As Boolean)
    booIsPopUp = booValue
    PropertyChanged "IsPopup"
End Property

Public Property Get StayUp() As Boolean
    StayUp = booStayUp
End Property

Public Property Let StayUp(ByVal booValue As Boolean)
    'If Ambient.UserMode Then Err.Raise 393
    booStayUp = booValue
    PropertyChanged "StayUp"
    tmrCheck.Enabled = Not booValue
    
    If booValue Then
     PopUp
    Else
     PopDown
    End If
End Property

Public Property Get Stretch() As Boolean
    Stretch = imgPIC.Stretch
End Property

Public Property Let Stretch(ByVal booStretch As Boolean)
    imgPIC.Stretch() = booStretch
    PropertyChanged "Stretch"
    
    UserControl_Resize
    Refresh
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal booValue As Boolean)
    UserControl.Enabled() = booValue
    PropertyChanged "Enabled"
    
    If UserControl.Enabled = False Then
        imgPIC.Picture = imgDPIC.Picture
    Else
        imgPIC.Picture = imgEPIC.Picture
    End If
    
    UserControl_Resize
    Refresh
End Property

Public Property Get Picture() As Picture
    Set Picture = imgEPIC.Picture
End Property

Public Property Set Picture(ByVal picValue As Picture)
    Set imgEPIC.Picture = picValue
    PropertyChanged "Picture"
    If Enabled Then
        imgPIC.Picture = imgEPIC.Picture
    Else
        imgPIC.Picture = imgDPIC.Picture
    End If

    UserControl_Resize
    Refresh
End Property
'
Public Property Get DisabledPicture() As Picture
    Set DisabledPicture = imgDPIC.Picture
End Property

Public Property Set DisabledPicture(ByVal picValue As Picture)
    Set imgDPIC.Picture = picValue
    PropertyChanged "DisabledPicture"
    
    UserControl_Resize
    Refresh
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'***************************************************************************
'Procedures
'***************************************************************************
Public Sub Refresh()
    UserControl.Refresh
End Sub

Public Sub PopPress()
    IsDown = False
    
    'Doe de acties met de lijnen
    With LineDown
        .Visible = True
        .X1 = 0
        .X2 = 600
        .Y1 = 0
        .Y2 = 0
    End With
    With LineUp
        .Visible = True
        .X1 = 600
        .X2 = 0
        .Y1 = 600
        .Y2 = 600
    End With
    With LineLeft
        .Visible = True
        .X1 = 600
        .X2 = 600
        .Y1 = 0
        .Y2 = 600
    End With
    With LineRight
        .Visible = True
        .X1 = 0
        .X2 = 0
        .Y1 = 0
        .Y2 = 600
    End With
'    imgPIC.Left = imgPIC.Left + 20
'    imgPIC.Top = imgPIC.Top + 20
End Sub

Public Sub PopDown()
    If StayUp = True Then Exit Sub
    If IsDown Then Exit Sub
    IsDown = True
    
'    If imgPIC.Left <> UserControl.ScaleWidth / 2 - imgPIC.Width / 2 Then _
'     imgPIC.Left = UserControl.ScaleWidth / 2 - imgPIC.Width / 2
'    If imgPIC.Top <> UserControl.ScaleHeight / 2 - imgPIC.Height / 2 Then _
'     imgPIC.Top = UserControl.ScaleHeight / 2 - imgPIC.Height / 2
     
    'Doe de acties met de lijnen
    With LineDown
        .Visible = False
        .X1 = 0
        .X2 = 600
        .Y1 = 0
        .Y2 = 0
    End With
    With LineUp
        .Visible = False
        .X1 = 600
        .X2 = 0
        .Y1 = 600
        .Y2 = 600
    End With
    With LineLeft
        .Visible = False
        .X1 = 600
        .X2 = 600
        .Y1 = 0
        .Y2 = 600
    End With
    With LineRight
        .Visible = False
        .X1 = 0
        .X2 = 0
        .Y1 = 0
        .Y2 = 600
    End With
End Sub

Public Sub PopUp()
    UserControl_Resize
    IsDown = False

    'Doe de acties met de lijnen
    With LineDown
        .Visible = True
        .X1 = 600
        .X2 = 0
        .Y1 = 600
        .Y2 = 600
    End With
    With LineUp
        .Visible = True
        .X1 = 0
        .X2 = 600
        .Y1 = 0
        .Y2 = 0
    End With
    With LineLeft
        .Visible = True
        .X1 = 0
        .X2 = 0
        .Y1 = 0
        .Y2 = 600
    End With
    With LineRight
        .Visible = True
        .X1 = 600
        .X2 = 600
        .Y1 = 0
        .Y2 = 600
    End With
End Sub

Public Sub CheckButton(hwd As Long, doit As String)
    If doit <> "vas123" Then
        MsgBox "CheckButton is an event that only the control can use. Please remove it from the project code.", vbCritical, "Error..."
        Exit Sub
    End If

    If GetActiveWindow() <> GetParentHwnd Then
        PopDown
        Exit Sub
    End If
    
    If hwnd = hwd Then
        If StayUp = True Then
            tmrCheck.Enabled = True
            Exit Sub
        End If
'        For Each ctl In UserControl.ParentControls
'            If ctl.hwnd = hwnd Then
'                ctl.ZOrder
'            End If
'        Next ctl
        PopUp
        If booIsClicked = True Then PopPress
        tmrLeft.Enabled = True
        tmrCheck.Enabled = False
        Exit Sub
    Else
        PopDown
    End If
    tmrCheck.Enabled = True
End Sub

'***************************************************************************
'Control events
'***************************************************************************
Private Sub imgPIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then PopPress
    booIsClicked = True
End Sub

Private Sub imgPIC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim curHwnd As Long
    Dim pt As POINTAPI
    
    GetCursorPos pt
    
    curHwnd = WindowFromPoint(pt.X, pt.Y)
    
    If curHwnd <> hwnd Then
        booIsClicked = False
        PopDown
        CheckButton hwnd, "vas123"
        Exit Sub
    Else
        booIsClicked = False
        PopUp
        If Button = vbLeftButton Then RaiseEvent Click
    End If
End Sub

Private Function GetParentHwnd() As Long
    GetParentHwnd = UserControl.Parent.ParentHwnd
End Function

Private Sub tmrCheck_Timer()
On Error Resume Next
    Dim pt As POINTAPI

    If GetActiveWindow() <> GetParentHwnd Then Exit Sub
    GetCursorPos pt
    new_HWND = WindowFromPoint(pt.X, pt.Y)
    old_HWND = new_HWND
    tmrCheck.Enabled = False
    If new_HWND <> Me.hwnd Then PopDown
    CheckButton new_HWND, "vas123"
End Sub

Private Sub tmrLeft_Timer()
On Error Resume Next
    Dim pt As POINTAPI

    GetCursorPos pt
    new_HWND = WindowFromPoint(pt.X, pt.Y)
    If new_HWND <> hwnd Then
        If booIsClicked Then PopUp Else PopDown
        IsPopup = False
        IsDown = True
        tmrLeft.Enabled = False
        tmrCheck.Enabled = True
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then PopPress
    booIsClicked = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim curHwnd As Long
    Dim pt As POINTAPI
    
    GetCursorPos pt
    curHwnd = WindowFromPoint(pt.X, pt.Y)
    If curHwnd <> hwnd Then
        booIsClicked = False
        PopDown
        CheckButton hwnd, "vas123"
        Exit Sub
    Else
        booIsClicked = False
        PopUp
        If Button = vbLeftButton Then RaiseEvent Click
    End If
End Sub

Private Sub UserControl_Resize()

    UserControl.Width = 610
    UserControl.Height = 610

    imgPIC.Left = 480
    imgPIC.Top = 120
    
    'Doe de acties met de lijnen
    With LineDown
        .X1 = 600
        .X2 = 0
        .Y1 = 600
        .Y2 = 600
    End With
    With LineUp
        .X1 = 0
        .X2 = 600
        .Y1 = 0
        .Y2 = 0
    End With
    With LineLeft
        .X1 = 0
        .X2 = 0
        .Y1 = 0
        .Y2 = 600
    End With
    With LineRight
        .X1 = 600
        .X2 = 600
        .Y1 = 0
        .Y2 = 600
    End With
    
    LineDown.Visible = True
    LineUp.Visible = True
    LineLeft.Visible = True
    LineRight.Visible = True
    
    If imgPIC.Stretch = True Then
        imgPIC.Top = 0
        imgPIC.Left = 0
        imgPIC.Height = UserControl.ScaleHeight
        imgPIC.Width = ScaleWidth
    Else
        If imgPIC.Left <> UserControl.ScaleWidth / 2 - imgPIC.Width / 2 Then imgPIC.Left = UserControl.ScaleWidth / 2 - imgPIC.Width / 2
        If imgPIC.Top <> UserControl.ScaleHeight / 2 - imgPIC.Height / 2 Then imgPIC.Top = UserControl.ScaleHeight / 2 - imgPIC.Height / 2
    End If

End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

