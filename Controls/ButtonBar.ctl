VERSION 5.00
Begin VB.UserControl ButtonBar 
   Alignable       =   -1  'True
   BackColor       =   &H80000000&
   ClientHeight    =   8415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9015
   DefaultCancel   =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   9015
   Begin VB.PictureBox picMain 
      BackColor       =   &H80000010&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin prjButtonBar.LargeIcon LargeIcon1 
         Height          =   1215
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   2143
      End
   End
End
Attribute VB_Name = "ButtonBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private cmdGroup                    As CommandButton
Attribute cmdGroup.VB_VarHelpID = -1
Private WithEvents cmdGroup0        As CommandButton
Attribute cmdGroup0.VB_VarHelpID = -1
Private WithEvents cmdGroup1        As CommandButton
Attribute cmdGroup1.VB_VarHelpID = -1
Private WithEvents cmdGroup2        As CommandButton
Attribute cmdGroup2.VB_VarHelpID = -1
Private WithEvents cmdGroup3        As CommandButton
Attribute cmdGroup3.VB_VarHelpID = -1
Private WithEvents cmdGroup4        As CommandButton
Attribute cmdGroup4.VB_VarHelpID = -1
Private WithEvents cmdGroup5        As CommandButton
Attribute cmdGroup5.VB_VarHelpID = -1
Private WithEvents cmdGroup6        As CommandButton
Attribute cmdGroup6.VB_VarHelpID = -1
Private WithEvents cmdGroup7        As CommandButton
Attribute cmdGroup7.VB_VarHelpID = -1
Private WithEvents cmdGroup8        As CommandButton
Attribute cmdGroup8.VB_VarHelpID = -1
Private WithEvents cmdGroup9        As CommandButton
Attribute cmdGroup9.VB_VarHelpID = -1

Private ctl                         As Object

'Local variables for the properties
Private intGroups                   As Integer
Private intCurrentGroup             As Integer
Private booEditable                 As Boolean
Private booIconsStayUp              As Boolean

'Variables
Private booChangeCanceled           As Boolean
Private intCounter                  As Integer

'Objects
Private WithEvents txtCaption       As TextBox
Attribute txtCaption.VB_VarHelpID = -1
Private objChangingCommand          As CommandButton
Private colIcons                    As New clsIcons

'Events
Event Click(Group As Integer, IconIndex As Integer)

'***************************************************************************
'Properties
'***************************************************************************
Public Property Let Caption(intIndex, strCaption As String)
    'Sets the caption for a group
    If intIndex <= Groups Then
        Select Case intIndex
            Case 0: cmdGroup0.Caption = strCaption
            Case 1: cmdGroup1.Caption = strCaption
            Case 2: cmdGroup2.Caption = strCaption
            Case 3: cmdGroup3.Caption = strCaption
            Case 4: cmdGroup4.Caption = strCaption
            Case 5: cmdGroup5.Caption = strCaption
            Case 6: cmdGroup6.Caption = strCaption
            Case 7: cmdGroup7.Caption = strCaption
            Case 8: cmdGroup8.Caption = strCaption
            Case 9: cmdGroup9.Caption = strCaption
        End Select
        PropertyChanged "caption"
    End If
End Property

Public Property Let Editable(booValue As Boolean)
    'Set to true when you can change the buttonbar
    booEditable = booValue
    PropertyChanged "Editable"
End Property

Public Property Get Editable() As Boolean
Attribute Editable.VB_ProcData.VB_Invoke_Property = ";Behavior"
    'Retrieves true/false if you can edit the buttonbar
    Editable = booEditable
End Property

Public Property Let Groups(intValue As Integer)
    'Set the number of groups
    If intValue > 10 Then intValue = 10
    intGroups = intValue
    SetGroups
    PropertyChanged "Groups"
End Property

Public Property Get Groups() As Integer
Attribute Groups.VB_Description = "Defines the number of groups"
Attribute Groups.VB_ProcData.VB_Invoke_Property = ";Appearance"
    'Retrieves the number of groups
    Groups = intGroups
End Property

Public Property Let CurrentGroup(intValue As Integer)
    'Set the number of groups
    intCurrentGroup = intValue
End Property

Public Property Get CurrentGroup() As Integer
    'Retrieves the number of groups
    CurrentGroup = intCurrentGroup
End Property

Public Property Let IconsStayUp(booValue As Boolean)
    booIconsStayUp = booValue
    
    For intCounter = 1 To LargeIcon1.Count - 1
        LargeIcon1(intCounter).IconsStayUp = booValue
    Next
End Property

Public Property Get IconsStayUp() As Boolean
Attribute IconsStayUp.VB_ProcData.VB_Invoke_Property = ";Behavior"
    IconsStayUp = booIconsStayUp
End Property

Public Property Get ParentHwnd() As Long
    'Retrieves the Hwnd from the parent
    ParentHwnd = Parent.hwnd
End Property

'***************************************************************************
'Methodes
'***************************************************************************
Public Sub Refresh()
    Dim objclsIcon As clsIcon
    
    For intCounter = 1 To LargeIcon1.Count - 1
        Unload LargeIcon1(intCounter)
    Next
    
    intCounter = 1
    For Each objclsIcon In colIcons
        If objclsIcon.ParentIndex = CurrentGroup Then
            'Load control
            Load LargeIcon1(intCounter)
            'Position of the icon
            Set LargeIcon1(intCounter).Icon = objclsIcon.Icon
            LargeIcon1(intCounter).Width = UserControl.Width
            Select Case objclsIcon.ParentIndex
                Case 0: LargeIcon1(intCounter).Top = cmdGroup0.Top + (objclsIcon.LocalIndex * LargeIcon1(intCounter).Height - 800)
                Case 1: LargeIcon1(intCounter).Top = cmdGroup1.Top + (objclsIcon.LocalIndex * LargeIcon1(intCounter).Height - 800)
                Case 2: LargeIcon1(intCounter).Top = cmdGroup2.Top + (objclsIcon.LocalIndex * LargeIcon1(intCounter).Height - 800)
                Case 3: LargeIcon1(intCounter).Top = cmdGroup3.Top + (objclsIcon.LocalIndex * LargeIcon1(intCounter).Height - 800)
                Case 4: LargeIcon1(intCounter).Top = cmdGroup4.Top + (objclsIcon.LocalIndex * LargeIcon1(intCounter).Height - 800)
                Case 5: LargeIcon1(intCounter).Top = cmdGroup5.Top + (objclsIcon.LocalIndex * LargeIcon1(intCounter).Height - 800)
                Case 6: LargeIcon1(intCounter).Top = cmdGroup6.Top + (objclsIcon.LocalIndex * LargeIcon1(intCounter).Height - 800)
                Case 7: LargeIcon1(intCounter).Top = cmdGroup7.Top + (objclsIcon.LocalIndex * LargeIcon1(intCounter).Height - 800)
                Case 8: LargeIcon1(intCounter).Top = cmdGroup8.Top + (objclsIcon.LocalIndex * LargeIcon1(intCounter).Height - 800)
                Case 9: LargeIcon1(intCounter).Top = cmdGroup9.Top + (objclsIcon.LocalIndex * LargeIcon1(intCounter).Height - 800)
            End Select
            'Check whether the icon must stay on top or not
            LargeIcon1(intCounter).IconsStayUp = IconsStayUp
            'Set caption of icon
            LargeIcon1(intCounter).Caption = objclsIcon.IconCaption
            LargeIcon1(intCounter).Visible = True
            'intcounter++
            intCounter = intCounter + 1
        End If
    Next
    
End Sub

Public Sub AddButton(Group As Integer, Index As Integer, Caption As String, Image As Picture)
    'Method that will add a button to a certain group
    colIcons.Add Group, Index, Caption, Image
    Refresh
End Sub

Public Sub DeleteButton(Group As Integer, Index As Integer)
    'Method that will delete a button from a certain group
    colIcons.Remove Group, Index
    Refresh

End Sub

'***************************************************************************
'Private Procedures
'***************************************************************************
Private Function SetGroups()
    Dim intI As Integer
    
    'Remove all buttons first
    For Each ctl In Controls
        If Left(ctl.Name, 8) = "cmdGroup" Then
            Controls.Remove ctl.Name
        End If
    Next
    
    'Build all buttons again...
    For intI = 0 To Groups
        'Make sure there are always two digits for the number when you name the button
        Set cmdGroup = Controls.Add("VB.CommandButton", "cmdGroup" & CStr(intI), picMain)
        With cmdGroup
           .Visible = True
           .Height = lngHeight
           .Width = .Parent.Width - 60
           .Caption = "Group " & CStr(intI)
           .Top = 0 + (intI * lngHeight)
           .Left = 0
           'Disable the focus on the commandbutton...
           'Put this line in comment until you compile ... otherwise you can't use stop
           NoFocusRect cmdGroup, True
        End With
        Select Case intI
            Case 0: Set cmdGroup0 = cmdGroup
            Case 1: Set cmdGroup1 = cmdGroup
            Case 2: Set cmdGroup2 = cmdGroup
            Case 3: Set cmdGroup3 = cmdGroup
            Case 4: Set cmdGroup4 = cmdGroup
            Case 5: Set cmdGroup5 = cmdGroup
            Case 6: Set cmdGroup6 = cmdGroup
            Case 7: Set cmdGroup7 = cmdGroup
            Case 8: Set cmdGroup8 = cmdGroup
            Case 9: Set cmdGroup9 = cmdGroup
        End Select
    Next
    CurrentGroup = 1
    
End Function

Private Sub ButtonClick(intIndex As Integer, Optional booRaiseEvent As Boolean = True)
    For Each ctl In Controls
        If Left(ctl.Name, 8) = "cmdGroup" Then
            If CInt(Right(ctl.Name, 1)) <= intIndex Then
                'Move button up
                ctl.Top = 0 + (CInt(Right(ctl.Name, 1)) * lngHeight)
            ElseIf CInt(Right(ctl.Name, 1)) > intIndex Then
                'Move button down
                ctl.Top = picMain.Height - ((Groups - CInt(Right(ctl.Name, 1)) + 1) * lngHeight) - 60
            End If
        End If
    Next
    CurrentGroup = intIndex
    Refresh

    'Trigger event
    If booRaiseEvent Then RaiseEvent Click(intCurrentGroup, intIndex)
End Sub

Private Sub ChangeCaption(cmdButton As CommandButton)
    If Editable Then
        'Remove textbox
        On Error Resume Next
        Controls.Remove txtCaption
        'Set the global variable
        Set objChangingCommand = cmdButton
        'Create a txtBox to change the caption of the cmdButton
        Set txtCaption = Controls.Add("VB.TextBox", "txtCaption", picMain)
        With txtCaption
           .BackColor = vbBlack
           .ForeColor = vbWhite
           .Alignment = vbCenter
           .Visible = True
           .Height = 285
           .Width = cmdButton.Width - 30
           .Left = 15
           .Top = cmdButton.Top + 15
           .Text = cmdButton.Caption
           .SelStart = Len(cmdButton.Caption)
           .SelLength = 0
           .ZOrder 0
        End With
        txtCaption.SetFocus
    End If
End Sub

'***************************************************************************
'Control events
'***************************************************************************

Private Sub LargeIcon1_Click(Index As Integer)
    RaiseEvent Click(CurrentGroup, Index)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Groups = PropBag.ReadProperty("Groups", 0)
    Editable = PropBag.ReadProperty("Editable", False)
    IconsStayUp = PropBag.ReadProperty("IconsStayUp", True)
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Groups", intGroups, 0)
    Call PropBag.WriteProperty("Editable", booEditable, False)
    Call PropBag.WriteProperty("IconsStayUp", booIconsStayUp, True)
    
End Sub

Private Sub UserControl_Resize()
    With picMain
        .Left = 0
        .Top = 0
        .Height = UserControl.Height
        .Width = UserControl.Width
    End With
End Sub

'***************************************************************************
'Textbox Events
'***************************************************************************
Private Sub txtCaption_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
        If Trim(txtCaption) = "" Then
        Else
            txtCaption_Validate False
            'Set the variable to nothing
            Set objChangingCommand = Nothing
            'Remove textbox
            Controls.Remove txtCaption
        End If
    ElseIf KeyCode = vbKeyEscape Then
        'Remove textbox
        booChangeCanceled = True
        Controls.Remove txtCaption
        Exit Sub
    End If
    booChangeCanceled = False
End Sub

Private Sub txtCaption_LostFocus()
    Controls.Remove txtCaption
End Sub

Private Sub txtCaption_Validate(Cancel As Boolean)
    On Error GoTo Hell
    If booChangeCanceled Then
    Else
        objChangingCommand.Caption = Trim(txtCaption)
    End If
    booChangeCanceled = False
Hell:
End Sub

'***************************************************************************
'Button events
'***************************************************************************
Private Sub cmdGroup0_Click()
    'Move the buttons
    ButtonClick 0, False
End Sub
Private Sub cmdGroup1_Click()
    'Move the buttons
    ButtonClick 1, False
End Sub
Private Sub cmdGroup2_Click()
    'Move the buttons
    ButtonClick 2, False
End Sub
Private Sub cmdGroup3_Click()
    'Move the buttons
    ButtonClick 3, False
End Sub
Private Sub cmdGroup4_Click()
    'Move the buttons
    ButtonClick 4, False
End Sub
Private Sub cmdGroup5_Click()
    'Move the buttons
    ButtonClick 5, False
End Sub
Private Sub cmdGroup6_Click()
    'Move the buttons
    ButtonClick 6, False
End Sub
Private Sub cmdGroup7_Click()
    'Move the buttons
    ButtonClick 7, False
End Sub
Private Sub cmdGroup8_Click()
    'Move the buttons
    ButtonClick 8, False
End Sub
Private Sub cmdGroup9_Click()
    'Move the buttons
    ButtonClick 9, False
End Sub
Private Sub cmdGroup0_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then ChangeCaption cmdGroup0
End Sub
Private Sub cmdGroup1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then ChangeCaption cmdGroup1
End Sub
Private Sub cmdGroup2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then ChangeCaption cmdGroup2
End Sub
Private Sub cmdGroup3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then ChangeCaption cmdGroup3
End Sub
Private Sub cmdGroup4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then ChangeCaption cmdGroup4
End Sub
Private Sub cmdGroup5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then ChangeCaption cmdGroup5
End Sub
Private Sub cmdGroup6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then ChangeCaption cmdGroup6
End Sub
Private Sub cmdGroup7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then ChangeCaption cmdGroup7
End Sub
Private Sub cmdGroup8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then ChangeCaption cmdGroup8
End Sub
Private Sub cmdGroup9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then ChangeCaption cmdGroup9
End Sub
