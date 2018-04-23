VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCOwnMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_ID As Long '// ID
Private m_hMenuID As Long '// The menu entry's ID
Private m_hMenu As Long '// The menu entry's ID
Private m_sMessage As String '// The menu entry's text
'Private m_objPicture As Object '// The menu entry's picture object
Private m_Icon As Integer ' icon to draw
Private m_hImageList As Long ' image list to draw from
Private m_lpDrawStruct As DRAWITEMSTRUCT '// The menu entry's current drawing information
Public hwndOwner As Long     '// the window which owns this object
Private m_Key As String
Private m_Style As Integer 'local copy
Private m_MenuBarText As String
Private m_Tag As Variant
Private m_HasSubMenu As Boolean
Private m_Disabled As Boolean

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long

Private Const SRCCOPY = &HCC0020

Private Const PS_SOLID = 0

Private Const COLOR_MENU = 4
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14

Private Const ODS_SELECTED = &H1

Private Const NEWTRANSPARENT = 1

Private Const DRAWWIDTH = 1
Private Const PicWidth = 20
Private Const TEXTBUFFER = 5

Public Property Let Caption(sMessage As String)
    m_sMessage = sMessage
End Property

Public Property Get Caption() As String
    Caption = m_sMessage
End Property

Public Sub InitMenu(hMenuID As Long, sMessage As String, iPicture As Integer, ByVal hImageList As Long, ByVal ID As Long, ByVal sKey As String, ByVal iStyle As Integer, ByVal hMenu As Long, ByVal sMenuBarText As String, ByVal vTag As Variant, ByVal bHasSubMenu As Boolean)
    '// Set object state
    m_hMenuID = hMenuID
    m_sMessage = sMessage
    m_Icon = iPicture
    m_hImageList = hImageList
    m_ID = ID
    m_Key = sKey
    m_Style = iStyle
    m_hMenu = hMenu
    m_MenuBarText = sMenuBarText
    m_Tag = vTag
    m_HasSubMenu = bHasSubMenu
    m_Disabled = False
End Sub

Public Sub InitStruct(hdc As Long, nAction As Long, nID As Long, nState As Long, nLeft As Long, nTop As Long, nBottom As Long, nRight As Long)
'// VB doesn't allow us to pass user defined structures to classes
'// therefore we have to go about it in a roundabout fashion
'// this leaves *you* as the programmer with room for future improvement
m_lpDrawStruct.hdc = hdc
m_lpDrawStruct.itemAction = nAction
m_lpDrawStruct.itemID = nID
m_lpDrawStruct.itemState = nState
m_lpDrawStruct.rcItem.Left = nLeft
m_lpDrawStruct.rcItem.Top = nTop
m_lpDrawStruct.rcItem.Bottom = nBottom
m_lpDrawStruct.rcItem.Right = nRight
End Sub

Public Property Get MenuID() As Long
MenuID = m_hMenuID
End Property

Private Sub PrintClear(crColor As Long, ByVal bMain As Boolean, ByVal hWnd As Long)
'// Set DC background mode to clear

SetBkMode m_lpDrawStruct.hdc, NEWTRANSPARENT

'// Get old type color
Dim crOldType As Long
Dim lOldFont As Long
Dim lNewFont As Long
crOldType = GetTextColor(m_lpDrawStruct.hdc)

'// Set new type color
If m_Disabled Then
    crColor = TranslateColor(vbHighlightText)
End If

SetTextColor m_lpDrawStruct.hdc, crColor

If bMain Then
    lNewFont = CreateMyFont(10, 0, hWnd, "Arial")
    lOldFont = SelectObject(m_lpDrawStruct.hdc, lNewFont)
    If m_Disabled Then
        SetTextColor m_lpDrawStruct.hdc, crColor
        TextOut m_lpDrawStruct.hdc, PicWidth + TEXTBUFFER + IIf(bMain, LEFTWIDTH, 0) + 1, m_lpDrawStruct.rcItem.Top + 7, m_sMessage, Len(m_sMessage)
        crColor = TranslateColor(vb3DShadow)
        SetTextColor m_lpDrawStruct.hdc, crColor
    End If
    TextOut m_lpDrawStruct.hdc, PicWidth + TEXTBUFFER + IIf(bMain, LEFTWIDTH, 0), m_lpDrawStruct.rcItem.Top + 6, m_sMessage, Len(m_sMessage)
    DeleteObject SelectObject(m_lpDrawStruct.hdc, lOldFont)
    DeleteObject lNewFont
Else
    '// Print text
    If m_Disabled Then
        crColor = TranslateColor(vbHighlightText)
        SetTextColor m_lpDrawStruct.hdc, crColor
        TextOut m_lpDrawStruct.hdc, PicWidth + TEXTBUFFER + IIf(bMain, LEFTWIDTH, 0) + 1, m_lpDrawStruct.rcItem.Top + 3, m_sMessage, Len(m_sMessage)
        crColor = TranslateColor(vb3DShadow)
        SetTextColor m_lpDrawStruct.hdc, crColor
    End If
    TextOut m_lpDrawStruct.hdc, PicWidth + TEXTBUFFER + IIf(bMain, LEFTWIDTH, 0), m_lpDrawStruct.rcItem.Top + 2, m_sMessage, Len(m_sMessage)
End If


'// Reset old color
SetTextColor m_lpDrawStruct.hdc, crOldType
End Sub

Public Sub DrawMenu(ByVal bMain As Boolean, ByVal MenuBarColor As OLE_COLOR, ByVal hWnd As Long, ByVal mID As Long, ByVal hMenu As Long, ByVal sMenuBarText As String, ByVal MenuBarTextColor As OLE_COLOR, ByVal bDisabled As Boolean)
'// Create a temporary copy of our member DRAWITEMSTRUCT
Dim lpDrawInfo As DRAWITEMSTRUCT
lpDrawInfo = m_lpDrawStruct
Static bDone As Boolean

'// Create a brushes or get colors for specific menu attributes
'// These attributes are selected by the user in the Windows Display settings dialog
'// This ensures that menu customization will affect our menus too
Dim hSelectedItem As Long, crSelected As Long
Dim hMenuColor As Long
Dim hMenuBar As Long
'hSelectedItem = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
'crSelected = GetSysColor(COLOR_HIGHLIGHTTEXT)
'hMenuColor = CreateSolidBrush(GetSysColor(COLOR_MENU))
hSelectedItem = CreateSolidBrush(gMenuHighlight)
crSelected = gMenuHighlightText
hMenuColor = CreateSolidBrush(gMenuBackColor)
        
'// Draw a highlight in the selection color if this element is selected
'// If this element is not selected we must clean up our previosu drawing

Dim DrawRect As RECT
Dim oPoint As POINTAPI
Dim lPen As Long
Dim lPen2 As Long
Dim lPen3 As Long
Dim lOldFont As Long
Dim lNewFont As Long
Dim oTest As RECT
Dim lOldBack As Long
Dim lNewBack As Long
Dim lNewDC As Long
Dim hIcon As Long

m_Disabled = bDisabled
DrawRect = m_lpDrawStruct.rcItem
DrawRect.Left = ((PicWidth + TEXTBUFFER) - 4) + IIf(bMain, LEFTWIDTH, 0)
DrawRect.Right = DrawRect.Right + IIf(bMain, LEFTWIDTH, 0)

SetBkColor lpDrawInfo.hdc, gMenuBackColor

If m_Style = 0 Then
    '// Print this menu entry's caption
    If lpDrawInfo.itemState = ODS_SELECTED And Not m_Disabled Then
        FillRect lpDrawInfo.hdc, DrawRect, hSelectedItem
    Else
        oTest = lpDrawInfo.rcItem
        oTest.Left = oTest.Left + IIf(bMain, LEFTWIDTH, 0)
        FillRect lpDrawInfo.hdc, oTest, hMenuColor
    End If
    PrintClear IIf(lpDrawInfo.itemState = ODS_SELECTED, crSelected, gMenuForeColor), bMain, hWnd
    '// Draw the bitmap for this menu entry
    
    SetBkMode m_lpDrawStruct.hdc, NEWTRANSPARENT
    
    If m_hImageList <> 0 Then
'        hIcon = ImageList_GetIcon(m_hImageList, m_Icon, 0)
'
'        If bMain Then
'            DrawIconEx lpDrawInfo.hdc, lpDrawInfo.rcItem.Left + DRAWWIDTH + 1 + IIf(bMain, LEFTWIDTH, 0), lpDrawInfo.rcItem.Top + DRAWWIDTH + 7, hIcon, 16, 16, 0, 0, DI_NORMAL
'        Else
'            DrawIconEx lpDrawInfo.hdc, lpDrawInfo.rcItem.Left + DRAWWIDTH + 1 + IIf(bMain, LEFTWIDTH, 0), lpDrawInfo.rcItem.Top + DRAWWIDTH + 1, hIcon, 16, 16, 0, 0, DI_NORMAL
'        End If
'
'        DestroyIcon hIcon

        If bMain Then
            ImageList_Draw m_hImageList, m_Icon, lpDrawInfo.hdc, lpDrawInfo.rcItem.Left + DRAWWIDTH + 1 + IIf(bMain, LEFTWIDTH, 0), lpDrawInfo.rcItem.Top + DRAWWIDTH + 7, ILD_TRANSPARENT
        Else
            ImageList_Draw m_hImageList, m_Icon, lpDrawInfo.hdc, lpDrawInfo.rcItem.Left + DRAWWIDTH + 1 + IIf(bMain, LEFTWIDTH, 0), lpDrawInfo.rcItem.Top + DRAWWIDTH + 1, ILD_TRANSPARENT
        End If
    
'        If bMain Then
'            ImageList_Draw m_hImageList, m_Icon, lpDrawInfo.hdc, lpDrawInfo.rcItem.Left + DRAWWIDTH + 1 + IIf(bMain, LEFTWIDTH, 0), lpDrawInfo.rcItem.Top + DRAWWIDTH + 7, 16, 16, CLR_NONE, CLR_DEFAULT, ILD_TRANSPARENT
'        Else
'            ImageList_Draw m_hImageList, m_Icon, lpDrawInfo.hdc, lpDrawInfo.rcItem.Left + DRAWWIDTH + 1 + IIf(bMain, LEFTWIDTH, 0), lpDrawInfo.rcItem.Top + DRAWWIDTH + 1, 16, 16, CLR_NONE, CLR_DEFAULT, ILD_TRANSPARENT
'        End If
    End If

    '// If our menu is selected we need to draw a 3D box around the picture
    'If lpDrawInfo.itemState = ODS_SELECTED Then Draw3D bMain
ElseIf m_Style = 1 Then
    ' separator bar
    oTest = lpDrawInfo.rcItem
    oTest.Left = oTest.Left + IIf(bMain, LEFTWIDTH, 0)
    FillRect lpDrawInfo.hdc, oTest, hMenuColor
    
    lPen = CreatePen(PS_SOLID, DRAWWIDTH, TranslateColor(vb3DShadow))
    lPen2 = CreatePen(PS_SOLID, DRAWWIDTH, TranslateColor(vb3DHighlight))
    lPen3 = CreatePen(PS_SOLID, DRAWWIDTH, RGB(0, 0, 0))
    MoveToEx lpDrawInfo.hdc, 3 + IIf(bMain, LEFTWIDTH, 0), lpDrawInfo.rcItem.Top + (lpDrawInfo.rcItem.Bottom - lpDrawInfo.rcItem.Top) / 2, oPoint
    DeleteObject SelectObject(lpDrawInfo.hdc, lPen)
    LineTo lpDrawInfo.hdc, lpDrawInfo.rcItem.Right - 3, lpDrawInfo.rcItem.Top + (lpDrawInfo.rcItem.Bottom - lpDrawInfo.rcItem.Top) / 2
    MoveToEx lpDrawInfo.hdc, 4 + IIf(bMain, LEFTWIDTH, 0), lpDrawInfo.rcItem.Top + ((lpDrawInfo.rcItem.Bottom - lpDrawInfo.rcItem.Top) / 2) + 1, oPoint
    DeleteObject SelectObject(lpDrawInfo.hdc, lPen2)
    LineTo lpDrawInfo.hdc, lpDrawInfo.rcItem.Right - 2, lpDrawInfo.rcItem.Top + ((lpDrawInfo.rcItem.Bottom - lpDrawInfo.rcItem.Top) / 2) + 1
    DeleteObject SelectObject(lpDrawInfo.hdc, lPen3)
    DeleteObject lPen
    DeleteObject lPen2
    DeleteObject lPen3
End If
'
    If bMain Then
        ' all of this draws the bar down the side, and the text on it.
        SetBkMode m_lpDrawStruct.hdc, NEWTRANSPARENT
        hMenuBar = CreateSolidBrush(TranslateColor(MenuBarColor))
        DrawRect.Left = lpDrawInfo.rcItem.Left
        DrawRect.Right = LEFTWIDTH - 1
        DrawRect.Top = 0
        DrawRect.Bottom = gBottom
        FillRect lpDrawInfo.hdc, DrawRect, hMenuBar
        DeleteObject hMenuBar
        If Len(MenuBarText) > 0 Then
            lPen = GetTextColor(lpDrawInfo.hdc)
            SetTextColor lpDrawInfo.hdc, TranslateColor(MenuBarTextColor)
            lNewFont = CreateMyFont(14, 90, hWnd, "Times New Roman")
            lOldFont = SelectObject(lpDrawInfo.hdc, lNewFont)
            TextOut lpDrawInfo.hdc, DrawRect.Left - 1, gBottom - 5, sMenuBarText, Len(sMenuBarText)
            SetTextColor lpDrawInfo.hdc, lPen
            DeleteObject SelectObject(lpDrawInfo.hdc, lOldFont)
            DeleteObject lNewFont
        End If
    End If

'// Delete used GDI objects
DeleteObject hSelectedItem
DeleteObject hMenuColor

End Sub

Private Sub Draw3D(ByVal bMain As Boolean)
'// Create a drawing space in
Dim rctPicture As RECT
Dim oPoint As POINTAPI
rctPicture.Top = m_lpDrawStruct.rcItem.Top
rctPicture.Left = m_lpDrawStruct.rcItem.Left + IIf(bMain, LEFTWIDTH, 0)
rctPicture.Right = PicWidth + IIf(bMain, LEFTWIDTH, 0)
rctPicture.Bottom = m_lpDrawStruct.rcItem.Bottom

'// Create pens for drawing the box border
Dim hpBlack As Long, hpWhite As Long

hpWhite = CreatePen(PS_SOLID, DRAWWIDTH, RGB(255, 255, 255)) '// White half of box
hpBlack = CreatePen(PS_SOLID, DRAWWIDTH, RGB(70, 70, 70))    '// Dark Grey half of box

'// Draw upper left corner of box
DeleteObject SelectObject(m_lpDrawStruct.hdc, hpWhite)

MoveToEx m_lpDrawStruct.hdc, rctPicture.Left + 1, (rctPicture.Bottom - 1), oPoint
LineTo m_lpDrawStruct.hdc, rctPicture.Left + 1, rctPicture.Top + 1
LineTo m_lpDrawStruct.hdc, (rctPicture.Right - 1), rctPicture.Top + 1

'// Draw lower right corner of box
DeleteObject SelectObject(m_lpDrawStruct.hdc, hpBlack)

LineTo m_lpDrawStruct.hdc, (rctPicture.Right - 1), rctPicture.Bottom - 1
LineTo m_lpDrawStruct.hdc, rctPicture.Left + 1, rctPicture.Bottom - 1

'// Clean up GDI objects
DeleteObject hpWhite
DeleteObject hpBlack
End Sub

Public Property Get mID() As Long
    mID = m_ID
End Property

Public Property Let mID(ByVal vNewValue As Long)
    m_ID = vNewValue
End Property

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(ByVal sNewKey As String)
    m_Key = sNewKey
End Property

Public Property Get Tag() As Variant
    Tag = m_Tag
End Property

Public Property Let Tag(ByVal sNewTag As Variant)
    m_Tag = sNewTag
End Property

Public Property Let Style(ByVal vData As Long)
    m_Style = vData
End Property

Public Property Get Style() As Long
    Style = m_Style
End Property

Public Property Let hMenu(ByVal vData As Long)
    m_hMenu = vData
End Property

Public Property Get hMenu() As Long
    hMenu = m_hMenu
End Property

Public Property Get MenuBarText() As String
    MenuBarText = m_MenuBarText
End Property

Public Property Let MenuBarText(ByVal sNewMenuBarText As String)
    m_MenuBarText = sNewMenuBarText
End Property

Public Property Let HasSubMenu(ByVal vData As Boolean)
    m_HasSubMenu = vData
End Property

Public Property Get HasSubMenu() As Boolean
    HasSubMenu = m_HasSubMenu
End Property

Public Property Let Disabled(ByVal vData As Boolean)
    m_Disabled = vData
End Property

Public Property Get Disabled() As Boolean
    Disabled = m_Disabled
End Property

