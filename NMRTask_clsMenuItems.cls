VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsMenuItem"
Attribute VB_Ext_KEY = "Member0" ,"clsMenuItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection

Public Function Add(Caption As String, Style As Long, Icon As Long, ByVal sKey As String, ByVal vTag As Variant) As clsMenuItem
On Error GoTo ErrorHandler
    'create a new object
    If Len(sKey) > 0 Then
        Dim objNewMember As clsMenuItem
        Set objNewMember = New clsMenuItem
    
        'set the properties passed into the method
        objNewMember.Caption = Caption
        objNewMember.Style = Style
        objNewMember.Icon = Icon
        objNewMember.Key = sKey
        objNewMember.Tag = vTag
        mCol.Add objNewMember, sKey
    
        'return the object created
        Set Add = objNewMember
        Set objNewMember = Nothing
    End If

Done:
    Exit Function
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    Select Case Err.Number
        Case 457
            Set Add = Nothing
            GoTo Done
        Case Else
            If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_Add" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "sKey=" & sKey)
            GoTo Done
    End Select
End Function

Public Property Get Item(Key As String) As clsMenuItem
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(Key)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(Key As String)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    mCol.Remove Key
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
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

Public Sub Clear()
    FreeMenus
    Set mCol = New Collection
End Sub
