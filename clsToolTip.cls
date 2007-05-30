VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Initialization of New ClassNames
Private Const ICC_BAR_CLASSES = &H4      'toolbar, statusbar, trackbar, tooltips
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Private Type tagINITCOMMONCONTROLSEX
   dwSize As Long   ' size of this structure
   dwICC As Long    ' flags indicating which classes to be initialized.
End Type

' ToolTip Styles
Private Const TTS_ALWAYSTIP = &H1
Private Const TTS_NOPREFIX = &H2
Private Const TTS_BALLOON = &H40 ' comctl32.dll v5.8 require

Private Const CW_USEDEFAULT = &H80000000

'Private Const WS_POPUP = &H80000000

Private Const WM_USER = &H400

' ToolTip Messages
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTM_ADDTOOL = (WM_USER + 4)
Private Const TTM_DELTOOL = (WM_USER + 5)
Private Const TTM_NEWTOOLRECT = (WM_USER + 6)
Private Const TTM_GETTOOLINFO = (WM_USER + 8)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3

Private Const TTF_IDISHWND = &H1
Private Const TTF_CENTERTIP = &H2
Private Const TTF_SUBCLASS = &H10

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TOOLINFO
    cbSize      As Long
    uFlags      As Long
    hWnd        As Long
    uId         As Long
    cRect       As RECT
    hinst       As Long
    lpszText    As String
End Type

Private Const LF_FACESIZE = 32
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
'        lfFaceName As String
        lfFaceName(LF_FACESIZE) As Byte
End Type

Public Enum TTStyle
    ttStyleStandard = 1
    ttStyleBalloon = 2
End Enum

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long


Private Const WM_SETFONT = &H30

Private Const GWL_STYLE = (-16)
Private hTT As Long
'local variable(s) to hold property value(s)
Private mvarTipWidth As Long 'local copy
Private mvarDelayTime As Long 'local copy
Private mvarVisibleTime As Long 'local copy
Private mvarBkColor As Long 'local copy
Private mvarTxtColor As Long 'local copy
Private mvarStyle As TTStyle 'local copy
Private mvarFont As StdFont 'local copy

Public Property Set Font(ByVal vData As StdFont)
  Dim lf As LOGFONT, hFont As Long, lHeight As Long
  lHeight = -MulDiv(vData.Size, GetDeviceCaps(GetDC(hTT), 90&), 72)
  lf.lfCharSet = vData.Charset
  lf.lfItalic = Abs(vData.Italic)
  lf.lfStrikeOut = Abs(vData.Strikethrough)
  lf.lfUnderline = Abs(vData.Underline)
  lf.lfWeight = vData.Weight
  Dim tmpArr() As Byte
  tmpArr = StrConv(vData.Name & Chr$(0), vbFromUnicode)
  Dim i As Integer, lArr As Long
  lArr = UBound(tmpArr)
  For i = 0 To lArr
      lf.lfFaceName(i) = tmpArr(i)
  Next i
  hFont = CreateFontIndirect(lf)
  SendMessageLong hTT, WM_SETFONT, hFont, 1&
  Set mvarFont = vData
End Property

Public Property Get Font() As StdFont
   Set Font = mvarFont
End Property

Public Property Let txtColor(ByVal vData As Long)
Attribute txtColor.VB_Description = "Set/Get ToolTip Text Color"
  mvarTxtColor = vData
  SendMessageLong hTT, TTM_SETTIPTEXTCOLOR, vData, 0&
End Property

Public Property Get txtColor() As Long
  txtColor = mvarTxtColor
End Property

Public Property Let BkColor(ByVal vData As Long)
Attribute BkColor.VB_Description = "Set/Get ToolTip Back Color"
  mvarBkColor = vData
  SendMessageLong hTT, TTM_SETTIPBKCOLOR, vData, 0&
End Property

Public Property Get BkColor() As Long
  BkColor = mvarBkColor
End Property

Public Property Let VisibleTime(ByVal vData As Long)
Attribute VisibleTime.VB_Description = "Set/Get ToolTip Visible Time"
   mvarVisibleTime = vData
   SendMessageLong hTT, TTM_SETDELAYTIME, TTDT_AUTOPOP, vData
End Property

Public Property Get VisibleTime() As Long
  VisibleTime = mvarVisibleTime
End Property

Public Property Let DelayTime(ByVal vData As Long)
Attribute DelayTime.VB_Description = "Set/Get ToolTip Delay Time"
    mvarDelayTime = vData
    SendMessageLong hTT, TTM_SETDELAYTIME, TTDT_INITIAL, vData
End Property

Public Property Get DelayTime() As Long
    DelayTime = mvarDelayTime
End Property

Public Property Let TipWidth(ByVal vData As Long)
Attribute TipWidth.VB_Description = "Set/Get ToolTip Maximum width"
   mvarTipWidth = vData
   SendMessageLong hTT, TTM_SETMAXTIPWIDTH, 0, vData
End Property

Public Property Get TipWidth() As Long
   TipWidth = mvarTipWidth
End Property

Public Property Let Style(ByVal vData As TTStyle)
   Dim lStyle As Long
   mvarStyle = vData
   If hTT Then
      lStyle = GetWindowLong(hTT, GWL_STYLE)
      If vData = ttStyleBalloon Then lStyle = lStyle Or TTS_BALLOON
      If vData = ttStyleStandard And (lStyle And ttStyleBalloon) Then lStyle = lStyle Xor TTS_BALLOON
      SetWindowLong hTT, GWL_STYLE, lStyle
   End If
End Property

Public Property Get Style() As TTStyle
Attribute Style.VB_Description = "Set/Get ToolTip Style (Standard/Balloon)"
   Style = mvarStyle
End Property

Private Sub InitComctl32(dwFlags As Long)
   Dim icc As tagINITCOMMONCONTROLSEX
   On Error GoTo Err_OldVersion
   icc.dwSize = Len(icc)
   icc.dwICC = dwFlags
   InitCommonControlsEx icc
   On Error GoTo 0
   Exit Sub
Err_OldVersion:
   InitCommonControls
End Sub

Public Sub SetToolTipObj(objHwnd As Long, sTipText As String, Optional bCenter As Boolean = False)
Attribute SetToolTipObj.VB_Description = "Assign ToolTip with sTipText to object (objHwnd)"
    Dim TI As TOOLINFO
    With TI
        .hWnd = objHwnd
        .uFlags = TTF_IDISHWND Or TTF_SUBCLASS
        If bCenter Then
            .uFlags = .uFlags Or TTF_CENTERTIP
        End If
        .uId = objHwnd
        .lpszText = sTipText
        .cbSize = Len(TI)
    End With
    SendMessage hTT, TTM_ADDTOOL, 0, TI
End Sub

Public Sub SetToolTipItem(objHwnd As Long, ByVal nItem As Long, lft As Long, tp As Long, rght As Long, btm As Long, sTipText As String, Optional bCenter As Boolean = False)
Attribute SetToolTipItem.VB_Description = "Assign ToolTip with stipText to specific item of Object/control with enum items (listbox, listview, treeview etc)"
   Dim TI As TOOLINFO, rc As RECT
   rc.Bottom = btm
   rc.Left = lft
   rc.Right = rght
   rc.Top = tp
   With TI
        .cRect = rc
        .hWnd = objHwnd
        .uFlags = TTF_SUBCLASS
        If bCenter Then
            .uFlags = .uFlags Or TTF_CENTERTIP
        End If
        .uId = nItem
        .lpszText = sTipText
        .cbSize = Len(TI)
    End With
    SendMessage hTT, TTM_ADDTOOL, 0, TI
End Sub

Public Sub DelToolTip(objHwnd As Long, Optional ByVal nItem As Long = -1)
Attribute DelToolTip.VB_Description = "Remove ToolTip from object or item"
   Dim TI As TOOLINFO
   TI.hWnd = objHwnd
   TI.cbSize = Len(TI)
   If nItem < 0 Then TI.uId = objHwnd Else TI.uId = nItem
   SendMessage hTT, TTM_DELTOOL, 0, TI
End Sub

Public Sub AjustItemRect(objHwnd As Long, nItem As Long, lft As Long, tp As Long, rght As Long, btm As Long)
Attribute AjustItemRect.VB_Description = "Change rectangular where ToolTip appear for specific item"
    Dim TI As TOOLINFO, rc As RECT
    With TI
        .hWnd = objHwnd
        .uId = nItem
        .cbSize = Len(TI)
    End With
    SendMessage hTT, TTM_GETTOOLINFO, 0&, TI
    rc.Bottom = btm
    rc.Left = lft
    rc.Right = rght
    rc.Top = tp
    TI.cRect = rc
    SendMessage hTT, TTM_NEWTOOLRECT, 0&, TI
End Sub

Private Sub Class_Initialize()
   InitComctl32 ICC_BAR_CLASSES
   hTT = CreateWindowEx(8, "tooltips_class32", 0&, TTS_NOPREFIX Or TTS_ALWAYSTIP, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 0&, 0&, App.hInstance, 0&)
   Style = ttStyleStandard
   TipWidth = 300
   BkColor = &HEEFFFF
   txtColor = vbBlack
   DelayTime = 500
   VisibleTime = 2000
End Sub

Private Sub Class_Terminate()
  If hTT Then DestroyWindow (hTT)
  If Not mvarFont Is Nothing Then DeleteObject ObjPtr(mvarFont)
End Sub

