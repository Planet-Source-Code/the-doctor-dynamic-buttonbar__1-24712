VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'local variable(s) to hold property value(s)
Private mvarParentIndex As Integer 'local copy
Private mvarLocalIndex As Integer 'local copy
Private mvarIconCaption As String
Private mvarIcon As Picture 'local copy

Public Property Let ParentIndex(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Target = 5
    mvarParentIndex = vData
End Property

Public Property Get ParentIndex() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Target
    ParentIndex = mvarParentIndex
End Property

Public Property Let LocalIndex(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Source = 5
    mvarLocalIndex = vData
End Property

Public Property Get LocalIndex() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Source
    LocalIndex = mvarLocalIndex
End Property

Public Property Let IconCaption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Source = 5
    mvarIconCaption = vData
End Property

Public Property Get IconCaption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Source
    IconCaption = mvarIconCaption
End Property

Public Property Let Icon(ByVal vData As Picture)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Source = 5
    Set mvarIcon = vData
End Property

Public Property Get Icon() As Picture
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Source
    Set Icon = mvarIcon
End Property
