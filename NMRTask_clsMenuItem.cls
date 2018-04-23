VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarCaption As String 'local copy
Private mvarStyle As Long 'local copy
Private mvarIcon As Long 'local copy
Private mvarKey As String
Private mvarTag As Variant

'local variable(s) to hold property value(s)
Private mvarSubItems As clsMenuItems 'local copy

Public Property Set SubItems(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SubItems = Form1
    Set mvarSubItems = vData
End Property

Public Property Get SubItems() As clsMenuItems
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubItems
    Set SubItems = mvarSubItems
End Property

Public Property Let Style(ByVal vData As Long)
    mvarStyle = vData
End Property

Public Property Get Style() As Long
    Style = mvarStyle
End Property

Public Property Let Icon(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Icon = 5
    mvarIcon = vData
End Property

Public Property Get Icon() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Icon
    Icon = mvarIcon
End Property

Public Property Let Caption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5
    mvarCaption = vData
End Property

Public Property Get Caption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property

Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property

Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Let Tag(ByVal vData As Variant)
    mvarTag = vData
End Property

Public Property Get Tag() As Variant
    Tag = mvarTag
End Property

Private Sub Class_Initialize()
    Set mvarSubItems = New clsMenuItems
End Sub

Private Sub Class_Terminate()
    Set mvarSubItems = Nothing
End Sub
