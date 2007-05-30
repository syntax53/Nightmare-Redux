Attribute VB_Name = "modDeclares"
Option Explicit
Option Compare Text

Public ghWnd As Long
Public ghWnd_Menu As Long
Public gMenuHighlight As Long
Public gMenuHighlightText As Long
Public gMenuBackColor As Long
Public gMenuForeColor As Long

Public Const SM_CXEDGE = 45
Public Const SM_CYEDGE = 46
Public Const SM_CYCAPTION = 4
'
Public Const BDR_INNER = &HC
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
Public Const BF_BOTTOM = &H8
Public Const BF_DIAGONAL = &H10
Public Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Public Const BF_LEFT = &H1
Public Const BF_MIDDLE = &H800    ' Fill in the middle.
Public Const BF_MONO = &H8000     ' For monochrome borders.
Public Const BF_RIGHT = &H4
Public Const BF_SOFT = &H1000     ' Use for softer buttons.
Public Const BF_TOP = &H2
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CHARSTREAM = 4          '  Character-stream, PLP
Public Const DT_CENTER = &H1
Public Const DT_DISPFILE = 6            '  Display-file
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_INTERNAL = &H1000
Public Const DT_LEFT = &H0
Public Const DT_METAFILE = 5            '  Metafile, VDM
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_PLOTTER = 0             '  Vector plotter
Public Const DT_RASCAMERA = 3           '  Raster camera
Public Const DT_RASDISPLAY = 1          '  Raster display
Public Const DT_RASPRINTER = 2          '  Raster printer
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
Public Const DT_PATH_ELLIPSIS = &H4000
Public Const DT_END_ELLIPSIS = &H8000
Public Const DT_WORD_ELLIPSIS = &H40000

Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9

Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDITILE = &H226
Public Const WM_SIZE = &H5
Public Const WM_KILLFOCUS = &H8
Public Const WM_SETFOCUS = &H7

Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDIREFRESHMENU = &H234
Public Const WM_MDISETMENU = &H230

Public Const WM_CLOSE = &H10

Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function CopyRect Lib "user32.dll" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Declare Function OleTranslateColor Lib "Olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hpal As Long, pcolorref As Long) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Declare Function GetSysColor Lib "user32" (ByVal nIndex As ColConst) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Enum ColConst
    COLOR_ACTIVEBORDER = 10
    COLOR_ACTIVECAPTION = 2
    COLOR_ADJ_MAX = 100
    COLOR_ADJ_MIN = -100
    COLOR_APPWORKSPACE = 12
    COLOR_BACKGROUND = 1
    COLOR_BTNFACE = 15
    COLOR_BTNHIGHLIGHT = 20
    COLOR_BTNSHADOW = 16
    COLOR_BTNTEXT = 18
    COLOR_CAPTIONTEXT = 9
    COLOR_GRAYTEXT = 17
    COLOR_HIGHLIGHT = 13
    COLOR_HIGHLIGHTTEXT = 14
    COLOR_INACTIVEBORDER = 11
    COLOR_INACTIVECAPTION = 3
    COLOR_INACTIVECAPTIONTEXT = 19
    COLOR_MENU = 4
    COLOR_MENUTEXT = 7
    COLOR_SCROLLBAR = 0
    COLOR_WINDOW = 5
    COLOR_WINDOWFRAME = 6
    COLOR_WINDOWTEXT = 8
End Enum

Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Public Const CLR_INVALID = &HFFFF

Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private objTB As ctlTaskBar
Private pfWndProc As Long
    
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetClassLong Lib "user32.dll" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SendMessageTimeout Lib "user32" Alias _
        "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal Msg As _
        Long, ByVal wParam As Long, ByVal lParam As Long, ByVal _
        fuFlags As Long, ByVal uTimeout As Long, lpdwResult As _
        Long) As Long
Public Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long

Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWndParent As Long, _
    ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Public Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_CHILD = &H40000000
Public Const GWL_STYLE = (-16)
Public Const GCL_HICON = (-14)
Public Const GCL_HICONSM = (-34)
Public Const WM_GETICON = &H7F
Public Const WM_SETICON = &H80
Public Const ICON_SMALL = 0
Public Const ICON_BIG = 1
Public Const DI_IMAGE = &H2
Public Const DI_COMPAT = &H4
Public Const DI_DEFAULTSIZE = &H8
Public Const DI_MASK = &H1
Public Const DI_NORMAL = &H3
Public Const WM_ACTIVATE = &H6
Public Const SIZE_MINIMIZED = 1


Public Declare Function TrackPopupMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal _
    uFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal _
    hWnd As Long, ByVal prcRect As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Public Const TPM_LEFTALIGN = &H0
Public Const TPM_TOPALIGN = &H0
Public Const TPM_NONOTIFY = &H80
Public Const TPM_RETURNCMD = &H100
Public Const TPM_LEFTBUTTON = &H0
Public Const TPM_VERTICAL = &H40&
Public Const TPM_RECURSE = &H1&
Public Const TPM_HORNEGANIMATION = &H800&
Public Const TPM_HORPOSANIMATION = &H400&
Public Const TPM_NOANIMATION = &H4000&
Public Const TPM_VERNEGANIMATION = &H2000&
Public Const TPM_VERPOSANIMATION = &H1000&

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_SYSCOMMAND = &H112
Public Const SC_SIZE = &HF000
Public Const SC_MOVE = &HF010
Public Const SC_CLOSE = &HF060
Public Const SC_MINIMIZE = &HF020
Public Const SC_MAXIMIZE = &HF030

Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ImageList_GetImageCount Lib "COMCTL32.DLL" (ByVal himl As Long) As Long
Public Declare Function ImageList_Draw Lib "COMCTL32.DLL" (ByVal himl As Long, ByVal I As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
Public Declare Function ImageList_DrawEx Lib "COMCTL32.DLL" (ByVal himl As Long, ByVal I As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
Public Declare Function ImageList_GetIcon Lib "COMCTL32.DLL" (ByVal himl As Long, ByVal I As Long, ByVal diFlags As Long) As Long


Public Const ILD_MASK = &H10
Public Const ILD_TRANSPARENT = &H1
Public Const ILD_SELECTED = &H4
Public Const ILD_FOCUS = &H4
Public Const ILD_NORMAL = &H0
Public Const ILD_BLEND = &H1
Public Const ILD_BLEND25 = &H2
Public Const ILD_BLEND50 = &H4
Public Const ILD_IMAGE = &H20
Public Const ILD_OVERLAYMASK = &HF00

Public Const CLR_NONE = &HFFFFFFFF
Public Const CLR_DEFAULT = &HFF000000

Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

' menus

Public Const MF_CHECKED = &H8&
Public Const MF_APPEND = &H100&
Public Const MF_DISABLED = &H2&
Public Const MF_GRAYED = &H1&
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const MF_POPUP = &H10&
Public Const MFS_ENABLED = &H0
Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MIIM_DATA As Long = &H20

Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
'Public Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
'///////////////////////////////////////////////////
'// m_omList() is a dynamic array of COwnMenu
'// objects which represent individual menu entries
'///////////////////////////////////////////////////
Private m_omList() As clsCOwnMenu
Public m_nOMCount As Long
Public m_bListInitialized As Boolean
Public mID As Long

Public Enum enmItemStyle
    Normal = 0
    Separator = 1
End Enum

'//////////////////////////////////////////////////////
'/// m_lPrevProc is the address of the procedure
'/// previously associated with the subclassed window
'//////////////////////////////////////////////////////
Private m_lPrevProc As Long

'////////////////////////////////////////////////////////////////
'//// Windows API functions
'////////////////////////////////////////////////////////////////
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetDCBrushColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function SetDCBrushColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal colorref As Long) As Long

'////////////////////////////////////////////////////////////////
'//// Windows API Constants
'////////////////////////////////////////////////////////////////
Public Const MF_OWNERDRAW = &H100&
Public Const MF_BYPOSITION = &H400&
Private Const GWL_WNDPROC = (-4)
Private Const WM_DRAWITEM = &H2B
Private Const WM_MEASUREITEM = &H2C
Private Const WM_COMMAND = &H111

'////////////////////////////////////////////////////////////////
'//// Structures used for Windows API functions
'////////////////////////////////////////////////////////////////

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Type MEASUREITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemWidth As Long
        itemHeight As Long
        ItemData As Long
End Type

Public Type DRAWITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemAction As Long
        itemState As Long
        hwndItem As Long
        hdc As Long
        rcItem As RECT
        ItemData As Long
End Type

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long

'// text measurement functions/structures
Public Type SIZE
    cx As Long
    cy As Long
End Type

Public Enum MENUCODE
    ByCommand = 0
    ByPosition = 1
End Enum

Public Const LEFTWIDTH = 25
Public gBottom As Long

Public Function SetMenuData(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As MENUCODE, ByVal wData As Long) As Long
    Dim minfo As MENUITEMINFO
    minfo.cbSize = Len(minfo)
    minfo.fMask = &H20
    minfo.dwItemData = wData
    SetMenuData = SetMenuItemInfo(hMenu, nPosition, -1 * wFlags, minfo)
End Function

Public Function GetMenuData(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As MENUCODE) As Long
    Dim minfo As MENUITEMINFO
    minfo.cbSize = Len(minfo)
    minfo.fMask = &H20
    GetMenuItemInfo hMenu, nPosition, -1 * wFlags, minfo
    GetMenuData = minfo.dwItemData
End Function


'//// FreeMenus - Frees the memory allocated on the heap
'////             for our COwnMenu objects
Public Sub FreeMenus()
Dim nIndex As Long
On Error Resume Next
If m_nOMCount > 0 Then
    For nIndex = 0 To m_nOMCount
        Set m_omList(nIndex) = Nothing
        DestroyMenu m_omList(nIndex).hMenu
    Next nIndex
End If
m_nOMCount = 0
ReDim m_omList(0)
m_bListInitialized = False

End Sub

'// Thiw procedure will tell Windows how big our items are.
Private Sub MeasureItem(ByRef mnu As clsCOwnMenu, ByRef lpMeasureInfo As MEASUREITEMSTRUCT, ByVal bMain As Boolean, ByVal bHasSubMenu As Boolean)
Dim hDrawDC As Long
Dim IMAGE_WIDTH As Long
Dim MENU_HEIGHT As Long
If bMain Then
    IMAGE_WIDTH = 24
    MENU_HEIGHT = 30
Else
    IMAGE_WIDTH = 16
    MENU_HEIGHT = 20
End If
Const MENU_SEP_HEIGHT = 10
hDrawDC = GetDC(mnu.hwndOwner)

Dim lpSize As POINTAPI
GetTextExtentPoint32 hDrawDC, mnu.Caption, Len(mnu.Caption), lpSize

Select Case mnu.Style
    Case 0
        lpMeasureInfo.itemHeight = MENU_HEIGHT
    Case 1
        lpMeasureInfo.itemHeight = MENU_SEP_HEIGHT
End Select

lpMeasureInfo.itemWidth = lpSize.x + IMAGE_WIDTH + IIf(bMain, LEFTWIDTH + 10, 10) + IIf(bHasSubMenu, 15, 0)
If lpMeasureInfo.itemWidth < objTB.MenuButtonWidth Then
    lpMeasureInfo.itemWidth = objTB.MenuButtonWidth - (IMAGE_WIDTH * 1.55) + IIf(bMain, LEFTWIDTH, 0) + IIf(bHasSubMenu, 15, 0)
End If
ReleaseDC mnu.hwndOwner, hDrawDC
End Sub

Public Sub MakeOwnerDraw(hMenu As Long, nIndex As Long, nID As Long)
'// Modify the menu's attributes
ModifyMenu hMenu, nIndex, MF_OWNERDRAW Or MF_BYPOSITION, nID, vbNullString
End Sub

Public Sub RegisterMenu(hMenu As Long, nPosition As Long, hwndOwner As Long, sMessage As String, iPicture As Integer, hImageList As Long, ByVal ID As Long, ByVal Key As String, ByVal iStyle As Integer, ByVal sMenuBarText As String, ByVal vTag As Variant, ByVal bHasSubMenu As Boolean)
'// Set this menu entry up as an owner drawn menu
    Dim lID As Long
    lID = GetMenuItemID(hMenu, nPosition)
    MakeOwnerDraw hMenu, nPosition, lID
    
    ' set the itemdata on the menu item, so we can catch it
    ' in the draw/measure events
    SetMenuData hMenu, nPosition, ByPosition, ID

'// Create a new owner drawn menu object on the heap
If (m_bListInitialized = False) Then
    ReDim m_omList(0)
    Set m_omList(0) = New clsCOwnMenu
    lID = GetMenuItemID(hMenu, nPosition)
    m_omList(0).InitMenu lID, sMessage, iPicture, hImageList, ID, Key, iStyle, hMenu, sMenuBarText, vTag, bHasSubMenu
    
    m_bListInitialized = True
Else
    m_nOMCount = m_nOMCount + 1
    
    ReDim Preserve m_omList(m_nOMCount)
    Set m_omList(m_nOMCount) = New clsCOwnMenu
    lID = GetMenuItemID(hMenu, nPosition)
    m_omList(m_nOMCount).hwndOwner = hwndOwner
    m_omList(m_nOMCount).InitMenu lID, sMessage, iPicture, hImageList, ID, Key, iStyle, hMenu, sMenuBarText, vTag, bHasSubMenu
End If
End Sub
' end menus


Public Function ShowSystemMenu(ByVal hWnd As Long)
On Error GoTo ErrorHandler

    Dim curpos As POINTAPI  ' holds the current mouse coordinates
    Dim retval As Long        ' generic return value
    Dim lMenu As Long
    Dim lSys As Long
    Dim lCount As Long
    
    
    ' get a copy of the system menu from the window.
    lSys = GetSystemMenu(hWnd, 0)
    
'    lCount = GetMenuItemCount(lSys)
    
'    If lCount = 9 Then
'        AppendMenu lSys, MF_SEPARATOR, 0&, vbNullString
'    End If
    
    retval = GetCursorPos(curpos)
    
    ' raise the menu at the current cursor position
    lMenu = TrackPopupMenu(lSys, TPM_RETURNCMD Or TPM_LEFTBUTTON Or TPM_TOPALIGN Or TPM_LEFTALIGN, curpos.x, curpos.y, 0, hWnd, 0)
    
    'handle menu clicks
    Select Case lMenu
        Case 61456 ' move
            DefWindowProc hWnd, WM_SYSCOMMAND, SC_MOVE, 0
            
        Case 61536 ' close
            PostMessage hWnd, WM_CLOSE, 0&, 0&
            
        Case 61504 ' next mdi child
            SendMessage ghWnd, WM_MDINEXT, hWnd, 0&
            
        Case 61440
            ' Size
            DefWindowProc hWnd, WM_SYSCOMMAND, SC_SIZE, 0
            
        Case 61472
            ' Minimize
            ShowWindow hWnd, 2
            
        Case 61488
            ' Maximize
            SendMessage ghWnd, WM_MDIMAXIMIZE, hWnd, 0&
            
        Case 61728
            ' restore
            SendMessage ghWnd, WM_MDIRESTORE, hWnd, 0&
        
        Case Else
            Debug.Print "Menu Item: " & lMenu
    End Select
Done:
    Exit Function
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class Declares_ShowSystemMenu" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "hWnd=" & hWnd)
    GoTo Done
End Function

Public Function GetWndIcon(hWndIcon As Long, bLarge As Boolean) As Long
    'Attempts to grab the icon for a window
    Dim lRet As Long
    On Error Resume Next
    ' First off, attempt WM_GETICON, use SendMesageTimeout so we don't
    '  hang on windows that aren't responding
    SendMessageTimeout hWndIcon, WM_GETICON, IIf(bLarge, ICON_BIG, _
                       ICON_SMALL), 0, 0, 1000, lRet
                
    If lRet = 0 Then
        ' If WM_GETICON didn't return anything, try using
        '  GetClassLong to get the icon for the window's class
        lRet = GetClassLong(hWndIcon, IIf(bLarge, GCL_HICON, _
                   GCL_HICONSM))
    End If
    
    If lRet = 0 Then
        
        'GetWndIcon = LoadIcon(0&, )
    End If
    
    GetWndIcon = lRet
    
End Function

Public Sub SubClassParentWnd(ByRef obj As ctlTaskBar, ByVal hWnd As Long)
    ' purpose of the function is to substitutue
    ' WndProc of MDIClient window
    
    Dim hWnd2 As Long
    ghWnd = FindWindowEx(GetParent(obj.hWnd), 0, "MDIClient", vbNullString)
    If GetParent(obj.hWnd) <> 0 Then
        hWnd2 = GetParent(obj.hWnd)
        Set objTB = obj
        pfWndProc = SetWindowLong(ghWnd, GWL_WNDPROC, AddressOf MDI_ParentWndProc)
        
        m_lPrevProc = GetWindowLong(hWnd, GWL_WNDPROC)
        ghWnd_Menu = hWnd
        SetWindowLong hWnd, GWL_WNDPROC, AddressOf IconProc
    End If
End Sub

Public Sub UnSubClassParentWnd(ByRef obj As ctlTaskBar)
    ' to revert the previous state
    
    If GetParent(obj.hWnd) <> 0 Then
        SetWindowLong ghWnd, GWL_WNDPROC, pfWndProc
        SetWindowLong ghWnd_Menu, GWL_WNDPROC, m_lPrevProc
        
        Set objTB = Nothing
    End If
End Sub

Function MDI_ParentWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    Dim lWinStyle As Long
    Static TheCount As Integer
    
    ' since message is handled, we can notify
    ' an object something interesting happened
    On Error Resume Next
    lRet = CallWindowProc(pfWndProc, hWnd, Msg, wParam, lParam)
    'Debug.Print lRet
    If GetParent(objTB.hWnd) <> 0 Then
    'Debug.Print Msg
        Select Case Msg
            Case WM_MDIGETACTIVE
                objTB.OnRefresh lRet
                objTB.RaiseChildActivate lRet
                TheCount = TheCount + 1
                Debug.Print "Get Active", TheCount
            
            Case WM_KILLFOCUS
                objTB.OnRefresh wParam
            
            Case WM_MDIRESTORE
                objTB.RaiseChildRestore wParam
            
            Case WM_MDIMAXIMIZE
                objTB.RaiseChildMaximize wParam
            
            Case WM_MDICREATE
                objTB.RaiseChildCreate lRet
            
            Case WM_MDIDESTROY
                objTB.OnRefresh
                objTB.RaiseChildDestroy wParam
                
        End Select
    End If

    MDI_ParentWndProc = lRet
End Function

Public Function IconProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim nRegisteredIndex As Long
'// Used to iterate through all registered menu objects
'// We must make sure that the menu object array has been initialized
'// if it has not then we have no business processing any messages
On Error Resume Next

Dim bDisabled As Boolean
If m_bListInitialized = False Then
    IconProc = CallWindowProc(m_lPrevProc, hWnd, uMsg, wParam, lParam)
    Exit Function
End If

'// The familiar window message select case block
Select Case uMsg
    Case WM_DRAWITEM
        '// The following code will copy a structure pointed to by lParam
        '// into our lpDrawInfo structure
        If wParam = 0 Then
            'wparam = 0 means its a menu
            ' we only want to do this for the menu's
            Dim lpDrawInfo As DRAWITEMSTRUCT
            CopyMem lpDrawInfo, ByVal lParam, Len(lpDrawInfo)
            
            ' fix the bottom, so we can draw the MenuBar
            If lpDrawInfo.rcItem.Bottom > gBottom Then gBottom = lpDrawInfo.rcItem.Bottom
            
            '// We must draw an owner drawn menu
            '// loop through all currently created menu objects
            '// and see if we have correctly received this message
            'If Not objTB Is Nothing And Not m_omList Is Nothing Then
                bDisabled = False
                For nRegisteredIndex = 0 To m_nOMCount
                    If (m_omList(nRegisteredIndex).mID) = lpDrawInfo.ItemData Then
                        '// We have found our registered menu
                        '// Let's tell the menu object to draw itself
                        m_omList(nRegisteredIndex).InitStruct lpDrawInfo.hdc, lpDrawInfo.itemAction, lpDrawInfo.itemID, lpDrawInfo.itemState, lpDrawInfo.rcItem.Left, lpDrawInfo.rcItem.Top, lpDrawInfo.rcItem.Bottom, lpDrawInfo.rcItem.Right
                        objTB.RaiseMenuItemDrawDisabled bDisabled, m_omList(nRegisteredIndex).Key, m_omList(nRegisteredIndex).Tag
                        m_omList(nRegisteredIndex).DrawMenu (m_omList(nRegisteredIndex).mID > 9000), objTB.MenuBarColor, objTB.hWnd, m_omList(nRegisteredIndex).mID, m_omList(nRegisteredIndex).hMenu, objTB.MenuBarText, objTB.MenuBarTextColor, bDisabled
                        Exit For
                    End If
                Next nRegisteredIndex
            'End If
            'hBrush = CreateSolidBrush(TranslateColor(vbRed))
            'FillRect lpDrawInfo.hdc, lpDrawInfo.rcItem, hBrush
        End If
    
    Case WM_MEASUREITEM
        Dim lpMeasureInfo As MEASUREITEMSTRUCT
        '// Get the MEASUREITEM struct from the pointer
        If wParam = 0 Then
            'wparam = 0 means its a menu
            'If Not m_omList Is Nothing Then
                CopyMem lpMeasureInfo, ByVal lParam, Len(lpMeasureInfo)
                For nRegisteredIndex = 0 To m_nOMCount
                    If (m_omList(nRegisteredIndex).mID) = lpMeasureInfo.ItemData Then
                        '// We have found our registered menu
                        MeasureItem m_omList(nRegisteredIndex), lpMeasureInfo, (m_omList(nRegisteredIndex).mID > 9000), m_omList(nRegisteredIndex).HasSubMenu
                        Exit For
                    End If
                Next nRegisteredIndex
                CopyMem ByVal lParam, lpMeasureInfo, Len(lpMeasureInfo)
            'End If
        End If
    
    Case WM_COMMAND
        ' handle the menu item click
        If Not objTB Is Nothing Then
            For nRegisteredIndex = 0 To m_nOMCount
                If (m_omList(nRegisteredIndex).mID) = wParam Then
                    '// We have found our registered menu
                    If m_omList(nRegisteredIndex).Disabled = False And m_omList(nRegisteredIndex).Style = 0 Then
                        objTB.RaiseMenuItemClick m_omList(nRegisteredIndex).Key, m_omList(nRegisteredIndex).Tag
                    End If
                    Exit For
                End If
            Next nRegisteredIndex
        End If
    Case Else
        
        '// Call previous WndProc
        IconProc = CallWindowProc(m_lPrevProc, hWnd, uMsg, wParam, lParam)
End Select
End Function


Public Function TranslateColor(ByVal clrColor As OLE_COLOR, _
    Optional hPalette As Long = 0) As Long
    On Error Resume Next
    If OleTranslateColor(clrColor, hPalette, TranslateColor) Then
      TranslateColor = CLR_INVALID
    End If

End Function

Public Function WindowText(ByVal hWnd As Long) As String
On Error GoTo ErrorHandler
    ' this function returns the caption for the window
    ' specified by hWnd
    Dim sCaption As String
    Dim nRet As Long
    
    sCaption = Space$(256)
    nRet = GetWindowText(hWnd, sCaption, Len(sCaption))
    If nRet Then
        sCaption = Left$(sCaption, nRet)
    End If
    
    WindowText = sCaption
Done:
    Exit Function
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class Declares_WindowText" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "hWnd=" & hWnd)
    GoTo Done
End Function
