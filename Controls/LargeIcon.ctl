VERSION 5.00
Begin VB.UserControl LargeIcon 
   BackColor       =   &H80000010&
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1530
   ControlContainer=   -1  'True
   ScaleHeight     =   1185
   ScaleWidth      =   1530
   Begin prjButtonBar.LargeButton3D LargeButton3D1 
      Height          =   615
      Left            =   480
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "Testcaption"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "LargeIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Local variables for the properties
Private booHide                     As Boolean

'Events
Event Click()

'***************************************************************************
'Properties
'***************************************************************************
Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal strValue As String)
    lblCaption.Caption = strValue
    PropertyChanged "Caption"
End Property

Public Property Get ParentHwnd() As Long
    'Retrieves the Hwnd from the parent
    ParentHwnd = UserControl.Parent.ParentHwnd
End Property

Public Property Let IconsStayUp(booValue As Boolean)
    LargeButton3D1.StayUp = booValue
End Property

Public Property Get IconsStayUp() As Boolean
    IconsStayUp = LargeButton3D1.StayUp
End Property

Public Property Get Hide() As Boolean
    Hide = booHide
End Property

Public Property Let Hide(booValue As Boolean)
    booHide = strValue
    PropertyChanged "Hide"
End Property

Public Property Get Icon() As Picture
    Set Icon = LargeButton3D1.Picture
End Property

Public Property Set Icon(icoValue As Picture)
    Set LargeButton3D1.Picture = icoValue
End Property

'***************************************************************************
'Procedures
'***************************************************************************
'***************************************************************************
'Control events
'***************************************************************************
Private Sub LargeButton3D1_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 1215
    
    lblCaption.Left = 0
    lblCaption.Width = UserControl.Width
    LargeButton3D1.Left = (UserControl.Width - LargeButton3D1.Width) / 2
End Sub

