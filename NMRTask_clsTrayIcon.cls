VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text

Private m_Icon As Integer
Private m_Key As String
Private m_ToolTip As String

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(ByVal vNewValue As String)
    m_Key = vNewValue
End Property

Public Property Get Icon() As Integer
    Icon = m_Icon
End Property

Public Property Let Icon(ByVal vNewValue As Integer)
    m_Icon = vNewValue
End Property

Public Property Get ToolTip() As String
    ToolTip = m_ToolTip
End Property

Public Property Let ToolTip(ByVal vNewValue As String)
    m_ToolTip = vNewValue
End Property

