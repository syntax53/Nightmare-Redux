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
Attribute VB_Exposed = False
Attribute VB_Description = "This stores all the information about each button."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text

' an object of this class represents a window
' icon displayed in application task bar

' the concept of 'touch' is for iterating through
' a collection of objects of the clsIcon type
' untouched items will be then marked (for delete ?)

Private m_strTitle As String ' holds the caption from the window
Private m_hWnd As Long ' the windows handle
Private m_nState As FormWindowStateConstants  ' minimized/maximozed/normal
Private m_hIcon As Long ' handle to the icon
Private m_Selected_Backcolour As Long ' Holds a colour
Private m_Flash As Boolean 'Tells us wether or not the Button is flashing
Private m_Unselected_Backcolour As Long ' Hold a Colour
Private m_FlashColour As Long 'Holds the flash colour
Private m_FlashOn As Boolean 'Used as a Switch
Private m_FlashCount As Integer 'the Count of flashes.
Private m_FlashTimes As Integer 'The Amount of time that the window will flash
Private m_Change_Font_Colour As Boolean 'Change the color??
Private m_FontColour As Long 'Hold the Font Colour

' m_ref_Icons refers to icon to be presented

Private m_bTouch As Boolean
Private m_bNew As Boolean ' is the icon new?
Private m_bChanged As Boolean ' has the window changed?

Public Property Get FontColour() As Long
    FontColour = m_FontColour
End Property

Public Property Let FontColour(ByVal Colour As Long)
    m_FontColour = Colour
End Property

Public Property Get FlashColour() As Long
    FlashColour = m_FlashColour
End Property

Public Property Let FlashColour(ByVal Colour As Long)
    m_FlashColour = Colour
End Property

Public Property Get Change_Font_Colour() As Boolean
    Change_Font_Colour = m_Change_Font_Colour
End Property


Public Property Let Change_Font_Colour(ByVal Colour As Boolean)
    m_Change_Font_Colour = Colour
    m_bChanged = True
    
End Property

Public Property Get SetFlashCount() As Integer
    SetFlashCount = m_FlashTimes
End Property

Public Property Let SetFlashCount(ByVal FlashTimes As Integer)
    m_FlashTimes = FlashTimes
End Property

Public Property Get IsNew() As Boolean
    IsNew = m_bNew
End Property

Public Property Get Title() As String
    Title = m_strTitle
End Property

Public Property Let Title(ByVal strNewValue As String)
    m_strTitle = strNewValue
    m_bChanged = True
End Property

Public Property Get Selected_BackColour() As Long
    Selected_BackColour = m_Selected_Backcolour
End Property

Public Property Let Selected_BackColour(ByVal Colour As Long)
    m_Selected_Backcolour = Colour
    m_bChanged = True
End Property

Public Property Get Unselected_BackColour() As Long
    Unselected_BackColour = m_Unselected_Backcolour
End Property

Public Property Let Unselected_BackColour(ByVal Colour As Long)
    m_Unselected_Backcolour = Colour
    m_bChanged = True
End Property

Public Property Get IsFlashing() As Boolean
    IsFlashing = m_Flash
End Property

Public Property Let IsFlashing(ByVal Flashing As Boolean)
    m_Flash = Flashing
    m_bChanged = True
End Property

Public Property Get FlashOn() As Boolean
    FlashOn = m_FlashOn
End Property

Public Property Let FlashOn(ByVal FlashOn As Boolean)
    m_FlashOn = FlashOn
End Property


Public Property Get FlashCount() As Integer
    FlashCount = m_FlashCount
End Property

Public Property Let FlashCount(ByVal Count As Integer)
    m_FlashCount = Count
End Property

Public Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

Public Property Let hWnd(ByVal hNewValue As Long)
    m_hWnd = hNewValue
End Property

Public Property Get State() As FormWindowStateConstants
    State = m_nState
End Property

Public Property Let State(ByVal newStateValue As FormWindowStateConstants)
    m_nState = newStateValue
End Property

Friend Sub ClearTouch()
    m_bNew = False
    m_bTouch = False
    m_bChanged = False
End Sub

Friend Sub Touch()
    m_bTouch = True
End Sub

Friend Function IsTaught() As Boolean
    IsTaught = m_bTouch
End Function

Private Sub Class_Initialize()
    m_bNew = True
    m_bTouch = True
    m_bChanged = True
    m_Selected_Backcolour = vbButtonFace
    m_Unselected_Backcolour = vbButtonFace
    m_FlashTimes = 5
    
    m_Change_Font_Colour = False
    m_FontColour = vbBlack

End Sub

Public Property Get IsChanged() As Boolean
    IsChanged = m_bChanged
End Property

Public Property Let IconPtr(ByVal hIcon As Long)
    m_hIcon = hIcon
End Property

Public Property Get IconPtr() As Long
On Error GoTo ErrorHandler
    'IconPtr = m_hIcon
    
    ' this gets the icon's handle and creates a local copy
    ' so we can use it
    Dim hIcon As Long
    Dim hLocalIcon As Long
    If hWnd <> 0 Then
        
        hIcon = GetWndIcon(m_hWnd, False)
        hLocalIcon = CopyIcon(hIcon)
        IconPtr = hLocalIcon
        
        If hLocalIcon <> 0 And m_hIcon <> 0 Then
            ' we had an icon already and are now replacing it.
            ' so we have to destroy the old one
            DestroyIcon m_hIcon
        End If
        
        ' now set the new one
        m_hIcon = hLocalIcon
    Else
        IconPtr = 0
        m_hIcon = 0
    End If
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class clsIcon_IconPtr" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Private Sub Class_Terminate()
    ' destroy our local copy of the icon ...
    If m_hIcon <> 0 Then
        DestroyIcon m_hIcon
    End If
End Sub


