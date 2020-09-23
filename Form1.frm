VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8685
   Begin VB.CommandButton Command5 
      Caption         =   "Remove icon"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add icon"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Icons stay on top"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Editable yes/no"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set captions-property"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "1. Change caption with right-mouse click on the groups"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MDIForm1.ButtonBar1.Caption(0) = "This"
    MDIForm1.ButtonBar1.Caption(1) = "property"
    MDIForm1.ButtonBar1.Caption(2) = "sets"
    MDIForm1.ButtonBar1.Caption(3) = "the"
    MDIForm1.ButtonBar1.Caption(4) = "captions"
    MDIForm1.ButtonBar1.Caption(5) = "for the groups"
End Sub

Private Sub Command2_Click()
    MDIForm1.ButtonBar1.Editable = Not MDIForm1.ButtonBar1.Editable
    If MDIForm1.ButtonBar1.Editable Then
        Command2.Caption = "&Editable"
    Else
        Command2.Caption = "Not &Editable"
    End If
End Sub

Private Sub Command3_Click()
    MDIForm1.ButtonBar1.IconsStayUp = Not MDIForm1.ButtonBar1.IconsStayUp
    If MDIForm1.ButtonBar1.IconsStayUp Then
        Command3.Caption = "&Icons StayUp"
    Else
        Command3.Caption = "&Icons do not StayUp"
    End If
End Sub

Private Sub Command4_Click()
    'In this version... you must make sure the localindex is ascending within a group
    MDIForm1.ButtonBar1.AddButton 1, 1, "Group 1 - One", MDIForm1.ImageList1.ListImages(1).Picture
    MDIForm1.ButtonBar1.AddButton 1, 2, "Group 1 - Two", MDIForm1.ImageList1.ListImages(2).Picture
    MDIForm1.ButtonBar1.AddButton 4, 1, "Group 4 - One", MDIForm1.ImageList1.ListImages(3).Picture
    MDIForm1.ButtonBar1.AddButton 5, 1, "Group 5 - One", MDIForm1.ImageList1.ListImages(4).Picture
End Sub

Private Sub Command5_Click()
    MDIForm1.ButtonBar1.DeleteButton 1, 2
End Sub
