VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(ParentIndex As Integer, LocalIndex As Integer, IconCaption As String, imgIcon As Picture) As clsIcon
    'create a new object
    Dim objNewMember As clsIcon
    Set objNewMember = New clsIcon

    'set the properties passed into the method
    objNewMember.ParentIndex = ParentIndex
    objNewMember.LocalIndex = LocalIndex
    objNewMember.IconCaption = IconCaption
    objNewMember.Icon = imgIcon

    mCol.Add objNewMember
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsIcon
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(ParentIndex As Integer, LocalIndex As Integer)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    Dim objclsIcon As clsIcon
    Dim colAfterDelete As New Collection
    
    For Each objclsIcon In mCol
        If objclsIcon.ParentIndex = ParentIndex And _
            objclsIcon.LocalIndex = LocalIndex Then
            'Don't put this class into the new collection
        Else
            'create a new object
            Dim objNewMember As clsIcon
            Set objNewMember = New clsIcon
        
            'set the properties passed into the method
            objNewMember.ParentIndex = objclsIcon.ParentIndex
            objNewMember.LocalIndex = objclsIcon.LocalIndex
            objNewMember.IconCaption = objclsIcon.IconCaption
            objNewMember.Icon = objclsIcon.Icon
        
            colAfterDelete.Add objNewMember
            
            Set objNewMember = Nothing
        End If
    Next
    Set mCol = Nothing
    Set mCol = colAfterDelete
    Set colAfterDelete = Nothing

End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

