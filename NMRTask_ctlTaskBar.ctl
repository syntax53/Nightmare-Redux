VERSION 5.00
Begin VB.UserControl ctlTaskBar 
   Alignable       =   -1  'True
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   ToolboxBitmap   =   "NMRTask_ctlTaskBar.ctx":0000
   Begin VB.Timer tmrRefresh 
      Interval        =   250
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer FlashTimer 
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrMouse 
      Interval        =   100
      Left            =   480
      Top             =   0
   End
End
Attribute VB_Name = "ctlTaskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

' default offsets and widths
Private Const DEFAULT_ITEM_WIDTH As Single = 250
Private Const FIRST_OFFSET As Single = 1
Private Const STANDARD_OFFSET As Single = 3
Private Const ICON_WIDTH As Single = 18

'For Button height
Private m_Task_Height As Integer

'Used to stop multiable building of the
'Private Tmr_Tick As Long
'Private Bol_Checking As Boolean
Private Default_Font_Colour As Long

'
' for drawing (button selection)
Private m_nIndexBeingSelected As Integer
Private m_bInsetSelected As Boolean
Private m_LastMouseOver As Integer
Private m_iLast As Integer
Private m_AvailableHeight As Long
Private m_AvailableWidth As Long

'Becouse I have no idea how the buld process works and I could not stop it moving the f$%king buttons
Private Bol_Refresh As Boolean

'elements linked with icons collection
'updated instantly on every change
Private m_maxCount As Integer
Public m_colIcons As Collection
Attribute m_colIcons.VB_VarMemberFlags = "440"
Public m_colTrayIcons As Collection
Attribute m_colTrayIcons.VB_VarMemberFlags = "440"
Private m_refActive As clsIcon

' for drawing
Private m_cxBorder As Long
Private m_cyBorder As Long
Private m_NoDraw As Boolean
Private m_ClickedMain As Boolean

'and sizing
Private m_nOptimalHeight As Long
Private m_nAlign As AlignConstants
Private m_ActualHeight As Long
Private m_ActualWidth As Long

' tool tips
Private m_strOriginalTooltip As String
Private m_bTooltip As Boolean

' properties
Private m_ForeColor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_RaisedBackColor As OLE_COLOR
Private m_SunkenBackColor As OLE_COLOR
Private m_SelectingBackColor As OLE_COLOR
Public Enum enmStyles
    Default = 0
    CoolBar = 1
End Enum
Private m_Style As enmStyles
Private m_CoolBarSeparator As Boolean
Private m_ShowActive As Boolean
Private m_ShowTray As Boolean
Private m_ShowMenu As Boolean
Private m_MenuCaption As String
Private m_AutoHide As Boolean
Private m_AutoHideWait As Integer
Private m_AutoHideAnimate As Boolean
Private m_AutoHideAnimateFrames As Integer
Private m_hImageList As Long
Private m_MenuButtonIcon As Long
Private m_MenuButtonWidth As Long
Private m_MenuBarColor As OLE_COLOR
Private m_MenuBarTextColor As OLE_COLOR
Private m_MenuBarText As String
Private m_MenuHighlightColor As OLE_COLOR
Private m_MenuHighlightTextColor As OLE_COLOR
Private m_MenuBackColor As OLE_COLOR
Private m_MenuForeColor As OLE_COLOR

' property defaults
Private Const m_def_ForeColor = vbButtonText
Private Const m_def_BackColor = vbButtonFace
Private Const m_def_RaisedBackColor = vbButtonFace
Private Const m_def_SunkenBackColor = vb3DHighlight
Private Const m_def_SelectingBackColor = vbButtonFace
Private Const m_def_Style = enmStyles.Default
Private Const m_def_CoolBarSeparator = False
Private Const m_def_ShowActive = True
Private Const m_def_ShowTray = False
Private Const m_def_ShowMenu = False
Private Const m_def_MenuCaption = "Start"
Private Const m_def_AutoHide = False
Private Const m_def_AutoHideWait = 1200
Private Const m_def_AutoHideAnimate = False
Private Const m_def_AutoHideAnimateFrames = 50
Private Const m_def_MenuButtonIcon = -1
Private Const m_def_MenuButtonWidth = 80
Private Const m_def_MenuBarColor = vbActiveTitleBar
Private Const m_def_MenuBarTextColor = vbActiveTitleBarText
Private Const m_def_MenuBarText = ""
Private Const m_def_MenuHighlightColor = vbHighlight
Private Const m_def_MenuHighlightTextColor = vbHighlightText
Private Const m_def_MenuBackColor = vbMenuBar
Private Const m_def_MenuForeColor = vbMenuText

' Events
Public Event ChildMinimize(ByVal hWnd As Long, ByVal Caption As String)
Attribute ChildMinimize.VB_Description = "This event is non-functional at the moment."
Public Event ChildMaximize(ByVal hWnd As Long, ByVal Caption As String)
Attribute ChildMaximize.VB_Description = "This event triggers when an MDI Child form is maximized within the MDI Client we are watching."
Public Event ChildRestore(ByVal hWnd As Long, ByVal Caption As String)
Attribute ChildRestore.VB_Description = "This event triggers when an MDI Child form Restores within the MDI Client window we are watching."
Public Event ChildActivate(ByVal hWnd As Long, ByVal Caption As String)
Attribute ChildActivate.VB_Description = "This event fires when a MDI Child form activates within the MDI Client form that we are watching."
Public Event ChildCreate(ByVal hWnd As Long, ByVal Caption As String)
Attribute ChildCreate.VB_Description = "This event triggers when an MDI Child is created within the MDI Client that we are watching."
Public Event ChildDestroy(ByVal hWnd As Long)
Attribute ChildDestroy.VB_Description = "This event triggers when an MDI Child form is closed/unloaded from within the MDI Client we are watching. "
Public Event AutoHide()
Public Event AutoHideShow()
Public Event TrayIconClick(ByVal Button As Integer, ByVal Index As Integer, ByVal Key As String, ByVal ToolTip As String)
Attribute TrayIconClick.VB_Description = "This event is triggered when an item in the tray area is clicked."
Public Event MenuItemClick(ByVal Key As String, ByVal vTag As Variant)
Public Event MenuItemDrawDisabled(ByRef Disabled As Boolean, ByVal Key As String, ByVal vTag As Variant)

Private m_Menu As Long
Public MainMenu As clsMenuItems
Attribute MainMenu.VB_VarMemberFlags = "400"
Private m_MenuItemData As Long
Private m_Main As Long

Private m_FontStyle As StdFont

Public Property Get ButtonFont() As StdFont
'On Error Resume Next
    Set ButtonFont = m_FontStyle
    
End Property
Public Property Set ButtonFont(ByVal Value As StdFont)
    Set m_FontStyle = Value
    Call PropertyChanged("ButtonFont")

End Property

' all of these Raise* sub's are here so the module can raise
' events on the taskbar.
Public Sub RaiseMenuItemDrawDisabled(ByRef Disabled As Boolean, ByVal Key As String, ByVal vTag As Variant)
On Error GoTo ErrorHandler
    
    RaiseEvent MenuItemDrawDisabled(Disabled, Key, vTag)

Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_RaiseMenuItemDrawDisabled" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "Disabled=" & Disabled)
    GoTo Done
End Sub

Public Sub RaiseMenuItemClick(ByVal Key As String, ByVal vTag As Variant)
Attribute RaiseMenuItemClick.VB_MemberFlags = "40"
On Error GoTo ErrorHandler
    If Len(Key) > 0 Then
        RaiseEvent MenuItemClick(Key, vTag)
    End If

Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_RaiseMenuItemClick" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "Key=" & Key)
    GoTo Done
End Sub

Public Sub RaiseChildCreate(ByVal hWnd As Long)
Attribute RaiseChildCreate.VB_MemberFlags = "40"
On Error GoTo ErrorHandler
    Dim sCaption As String
    sCaption = WindowText(hWnd)
    
    RaiseEvent ChildCreate(hWnd, sCaption)

Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_RaiseChildCreate" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "hWnd=" & hWnd)
    GoTo Done
End Sub

Public Sub RaiseChildDestroy(ByVal hWnd As Long)
Attribute RaiseChildDestroy.VB_MemberFlags = "40"
On Error GoTo ErrorHandler
    RaiseEvent ChildDestroy(hWnd)
Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_RaiseChildDestroy" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "hWnd=" & hWnd)
    GoTo Done
End Sub

Public Sub RaiseChildMinimize(ByVal hWnd As Long)
Attribute RaiseChildMinimize.VB_MemberFlags = "40"
On Error GoTo ErrorHandler
    Dim sCaption As String
    sCaption = WindowText(hWnd)
    
    RaiseEvent ChildMinimize(hWnd, sCaption)

Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_RaiseChildMinimize" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "hWnd=" & hWnd)
    GoTo Done
End Sub

Public Sub RaiseChildMaximize(ByVal hWnd As Long)
Attribute RaiseChildMaximize.VB_MemberFlags = "40"
On Error GoTo ErrorHandler
    Dim sCaption As String
    sCaption = WindowText(hWnd)

    RaiseEvent ChildMaximize(hWnd, sCaption)
    
Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_RaiseChildMaximize" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "hWnd=" & hWnd)
    GoTo Done
End Sub

Public Sub RaiseChildRestore(ByVal hWnd As Long)
Attribute RaiseChildRestore.VB_MemberFlags = "40"
On Error GoTo ErrorHandler
    Dim sCaption As String
    sCaption = WindowText(hWnd)
    
    RaiseEvent ChildRestore(hWnd, sCaption)
    
Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_RaiseChildRestore" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "hWnd=" & hWnd)
    GoTo Done
End Sub

Public Sub RaiseChildActivate(ByVal hWnd As Long)
Attribute RaiseChildActivate.VB_MemberFlags = "40"
On Error GoTo ErrorHandler
    Dim sCaption As String
    
    sCaption = WindowText(hWnd)
    RaiseEvent ChildActivate(hWnd, sCaption)
    
Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_RaiseChildActivate" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "hWnd=" & hWnd)
    GoTo Done
End Sub

Friend Property Get hWnd() As Long
On Error GoTo ErrorHandler
    ' this call is simple, it returns the usercontrols hWnd,
    ' so we can use it in GetParent() API calls in the module
    hWnd = UserControl.hWnd
    
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_hWnd" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Friend Sub OnRefresh(Optional ByVal hWndActive As Long = 0)
On Error GoTo ErrorHandler

' we are about to refresh our taskbar
' called only from substitutedWndProc for MDI window
If m_refActive Is Nothing Then
    UpdateIconsCollection hWndActive
    MapIconCollection

ElseIf hWndActive <> m_refActive.hWnd Then
    UpdateIconsCollection hWndActive
    MapIconCollection

ElseIf hWndActive = m_refActive.hWnd Then
    '// added to fix problem when minimized more then 2 windows
    UpdateIconsCollection hWndActive
    MapIconCollection
    If m_refActive.State = vbMinimized Then
        Call ShowWindow(m_refActive.hWnd, SW_HIDE)
    End If
    '// end minimize fix
    
    If Not m_refActive.Title = WindowText(hWndActive) Then
        m_refActive.Title = WindowText(hWndActive)
        PaintOne hWndActive
    End If
End If

Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_OnRefresh" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "hWndActive=" & hWndActive)
    GoTo Done
End Sub

Private Sub FlashTimer_Timer()
On Error GoTo ErrorHandler
    ' painting whole area
    Dim I As Integer
    Dim rcItem As RECT
    Dim Icn As clsIcon
    Dim lEdgeParam As Long
    Dim hBrush As Long  ' receives handle to the blue hatched brush to use
    Dim r As RECT  ' rectangular area to fill
    Dim lRet As Long  ' return value
    Dim nDiff As Single
    Dim nIconTop As Single
    Dim rcIcon As RECT
    Dim lpDrawTextParams As DRAWTEXTPARAMS
    Dim nTextH As Single
    Dim oTest As POINTAPI
    Dim oTrayPoint As POINTAPI
    Dim vEdge As Variant
    Dim oTrayIcon As clsTrayIcon
    Dim bRet As Boolean
    Dim j As Integer
    Dim Font_Color As Long
        
    
    'Stop the Timer
    FlashTimer.Enabled = False
    
    If m_colIcons Is Nothing Or m_NoDraw = True Then
        ' no buttons, clear the control
        UserControl.Cls
        Exit Sub
    End If
    
    I = 0
    
    ' set the colors
    UserControl.BackColor = m_BackColor
    UserControl.ForeColor = m_ForeColor
    
    lpDrawTextParams.iLeftMargin = 1
    lpDrawTextParams.iRightMargin = 1
    lpDrawTextParams.iTabLength = 2
    lpDrawTextParams.cbSize = 20
    
    I = 0
    For Each Icn In m_colIcons
        I = I + 1
        If Icn.IsFlashing Then
            If Icn.FlashOn Then
                Icn.FlashOn = False
                If Icn.FlashCount >= Icn.SetFlashCount Then
                    Icn.FlashOn = False
                    Icn.IsFlashing = False
                    Icn.FlashCount = 0
                End If
            Else
                Icn.FlashOn = True
                Icn.FlashCount = Icn.FlashCount + 1
            End If
            
        Else
            GoTo NextItem
        End If

        If Icn Is m_refActive Then
            Icn.FlashOn = False
            Icn.IsFlashing = False
            Icn.FlashCount = 0
            'GoTo NextItem
        End If
        
        ' three states of a push button
        If Icn Is m_refActive And m_ShowActive Then
            vEdge = EDGE_SUNKEN
            UserControl.FontBold = True
            hBrush = CreateSolidBrush(TranslateColor(Icn.Selected_BackColour))
        
        ' NOT being pushed at the moment
        ElseIf Not I = m_nIndexBeingSelected Then
            vEdge = EDGE_RAISED
            If Icn.IsFlashing And Icn.FlashOn Then
                hBrush = CreateSolidBrush(TranslateColor(Icn.FlashColour))
            Else
                hBrush = CreateSolidBrush(TranslateColor(Icn.Unselected_BackColour))
            End If
            UserControl.FontBold = False
        
        ' being pushed at the moment
        Else
            vEdge = EDGE_SUNKEN
            UserControl.FontBold = False
            hBrush = CreateSolidBrush(TranslateColor(m_SelectingBackColor))
       
        End If
                                
        If ItemRect(I, rcItem) Then
            ' draw the edge.
            If vEdge <> 0 Then
                DrawEdge UserControl.hdc, rcItem, vEdge, BF_RECT
            End If
        
            ' we selected the right hBrush up above, now use it to
            ' draw the back color.
            lRet = CopyRect(r, rcItem)
            r.Bottom = r.Bottom - 2
            r.Top = r.Top + 1
            r.Left = r.Left + 1
            r.Right = r.Right - 2
            lRet = FillRect(UserControl.hdc, r, hBrush)  ' fill the rectangle using the brush
            lRet = DeleteObject(hBrush) ' clean up
                
            ' fix the rect to fit inside the new border thats
            ' been drawn
            rcItem.Left = rcItem.Left + m_cxBorder + 1
            rcItem.Top = rcItem.Top + m_cyBorder
            rcItem.Right = rcItem.Right - m_cxBorder - 1
            rcItem.Bottom = rcItem.Bottom - m_cyBorder - 1
            
            ' used to calculate the position to draw the icon
            nDiff = rcItem.Bottom - rcItem.Top
            
            ' draw the icon
            ' calculate the position to draw the icon
            nIconTop = rcItem.Top + (nDiff - ICON_WIDTH) \ 2
            If Icn.IconPtr <> 0 Then
                DrawIconEx UserControl.hdc, rcItem.Left, nIconTop + 2, Icn.IconPtr, 16, 16, 0, 0, DI_NORMAL
            Else
                ' no icon was returned, so we cant draw anything.
            End If
            
            ' drawing a text with default font for a control
            If Icn.IconPtr <> 0 Then
                ' has an icon. so add space for it.
                rcItem.Left = rcItem.Left + ICON_WIDTH + 2
            Else
                ' no icon, so draw it over to the left.
                rcItem.Left = rcItem.Left + 2
            End If
            
            lpDrawTextParams.iLeftMargin = 1
            lpDrawTextParams.iRightMargin = 1
            lpDrawTextParams.iTabLength = 2
            lpDrawTextParams.cbSize = 20
            
            ' calculate all the dimensions for the rect to
            ' draw the text in.
            GetTextExtentPoint32 UserControl.hdc, Icn.Title, Len(Icn.Title), oTest
            nTextH = oTest.y
            If nTextH < nDiff Then
                nDiff = (nDiff - nTextH) \ 2
                rcItem.Bottom = rcItem.Bottom - nDiff
                rcItem.Top = rcItem.Top + nDiff
            End If
            
            'Set the Font Style.
            Set UserControl.Font = m_FontStyle
            
            If Icn.Change_Font_Colour Then
                UserControl.ForeColor = Icn.FontColour
               
            Else
                UserControl.ForeColor = Default_Font_Colour
               
            End If
                
            ' draw the text
            DrawTextEx UserControl.hdc, Icn.Title, Len(Icn.Title), rcItem, _
            DT_LEFT Or DT_VCENTER Or DT_WORD_ELLIPSIS, lpDrawTextParams
            
            
            'Set the default colour back to normal
            If Icn.Change_Font_Colour Then
               UserControl.ForeColor = Default_Font_Colour
                 
            End If
        
        End If
        
        j = I + 1
        
        ' now if it is CoolBar style, we draw the separator
        ' line inbetween buttons
        If m_colIcons.Count > 1 And j < m_colIcons.Count Then
            If m_Style = CoolBar And m_CoolBarSeparator = True Then
                If m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
                    ' draw 2 lines to give it the 3d look.
                    UserControl.Line (r.Right + 3, r.Top)-(r.Right + 3, r.Bottom), vb3DShadow
                    UserControl.Line (r.Right + 4, r.Top + 1)-(r.Right + 4, r.Bottom + 1), TranslateColor(m_SunkenBackColor)
                ElseIf m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
                    UserControl.Line (r.Left, r.Bottom + 3)-(r.Right, r.Bottom + 3), vb3DShadow
                    UserControl.Line (r.Left, r.Bottom + 4)-(r.Right, r.Bottom + 4), TranslateColor(m_SunkenBackColor)
                End If
            End If
        End If
NextItem:
    Next
    
    ' draw the tray
    If m_ShowTray And (Not m_colTrayIcons Is Nothing) Then
        TrayRect rcItem
        DrawEdge UserControl.hdc, rcItem, BDR_SUNKENOUTER, BF_RECT
        I = -1
        For Each oTrayIcon In m_colTrayIcons
            I = I + 1
            If TrayIconPoint(I, oTrayPoint) Then
                If hImageList <> 0 Then
                    ImageList_DrawEx hImageList, oTrayIcon.Icon, UserControl.hdc, oTrayPoint.x, oTrayPoint.y, 16, 16, CLR_NONE, CLR_DEFAULT, ILD_TRANSPARENT
                End If
            End If
        Next
        Set oTrayIcon = Nothing
    End If
    
Done:
    DeleteObject hBrush
    
    'Start the Timer
    FlashTimer.Enabled = True
    
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_UserControl_Paint" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Sub

'
Private Sub tmrMouse_Timer()
On Error GoTo ErrorHandler
    ' this timer is used to take away the focus rect, for bars
    ' that are coolbar style, when the mouse leaves the
    ' window for the usercontrol.
'    On Error Resume Next
    Dim oPoint As POINTAPI
    Dim lRet As Long
    Dim hWnd As Long
    Dim I As Integer
    Dim lWait As Long
    Dim lTemp As Long
    Dim iCount As Integer
    Static bTest As Boolean
    Static iLast As Integer
    Static lstart As Long
    Static lNow As Long
    Static bFirstDone As Boolean
    
    If Not Ambient.UserMode() Then Exit Sub
    
    ' get the X and Y of the mouse's current position
    lRet = GetCursorPos(oPoint)
    hWnd = WindowFromPoint(oPoint.x, oPoint.y)
    ' get the handle of the window underneath that X and Y
    
    ' if its the first time through we have to
    ' set these up with defaults.
    If lstart = 0 And lNow = 0 Then
        lstart = GetTickCount()
        lNow = lstart
    End If
    ' if the last time we came through this sub
    ' the hWnd was NOT the same as the usercontrol.hWnd (btest = false)
    ' AND we havent re-drawn the last mouse'd over
    ' element (ilast = 1) then redraw it, to clear away
    ' the mouseover border
    If bTest = False And iLast = 1 Then
        iLast = 0
        If m_Style = CoolBar Then
            InvalidateElement m_LastMouseOver
        End If
        m_iLast = -1
    End If
    If hWnd = UserControl.hWnd Then
        'If UserControl.Ambient.UserMode() Then MsgBox "Showing"
        bTest = True
        iLast = 1
        ' handle AutoHide = True
        If m_AutoHide Then
            ' bring it back to size
            If m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
                If UserControl.Extender.Height <> m_ActualHeight Then
                    UserControl.Extender.Height = m_ActualHeight
                    RaiseEvent AutoHideShow
                End If
            ElseIf m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
                If UserControl.Extender.Width <> m_ActualWidth Then
                    UserControl.Extender.Width = m_ActualWidth
                    RaiseEvent AutoHideShow
                End If
            End If
            ' allow the paint to happen now
            m_NoDraw = False
        End If
    Else
        'If UserControl.Ambient.UserMode() Then MsgBox "Hiding"
        ' the gettickcount's are used to time how long they have been off the
        ' bar, so we can hold off on making the bar hide, for a couple
        ' seconds.
        If bTest = True Then
            ' get the starting time
            lstart = GetTickCount()
            bTest = False
        Else
            ' now every loop through, while we are off, get the new time
            lNow = GetTickCount()
        End If
    
        ' handle AutoHide = True
        ' if lNow is more than m_AutoHideWait AFTER lStart then hide
        If Not m_ClickedMain And m_AutoHide And (((lNow - lstart) > m_AutoHideWait) Or bFirstDone = False) Then
            ' we shrink it down, because they moved the mouse
            ' outside of the usercontrol
            If m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
                If m_AutoHideAnimate And UserControl.Extender.Height <> 75 And m_AutoHideAnimateFrames > 1 Then
                    ' all of this code handles the animation
                    lTemp = UserControl.Extender.Height \ m_AutoHideAnimateFrames
                    iCount = m_AutoHideAnimateFrames + 1
                    Do Until iCount = 1
                        iCount = iCount - 1
                        lWait = GetTickCount()
                        If (lTemp * iCount) < 75 Then
                            Exit Do
                        End If
                        UserControl.Extender.Height = lTemp * iCount
                        Do Until (GetTickCount() - lWait) > 3
                            DoEvents
                        Loop
                    Loop
                End If
                UserControl.Extender.Height = 75
                RaiseEvent AutoHide
            ElseIf m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
                If m_AutoHideAnimate And UserControl.Extender.Width <> 75 And m_AutoHideAnimateFrames > 1 Then
                    lTemp = UserControl.Extender.Width \ m_AutoHideAnimateFrames
                    iCount = m_AutoHideAnimateFrames + 1
                    Do Until iCount = 1
                        iCount = iCount - 1
                        lWait = GetTickCount()
                        If (lTemp * iCount) < 75 Then
                            Exit Do
                        End If
                        UserControl.Extender.Width = lTemp * iCount
                        Do Until (GetTickCount() - lWait) > 3
                            DoEvents
                        Loop
                    Loop
                End If
                UserControl.Extender.Width = 75
                RaiseEvent AutoHide
            End If
            ' make sure we arent painting, and clear the usercontrol
            m_NoDraw = True
            ' bFirstDone is used to hide the bar initially if autohide
            ' is turned on
            bFirstDone = True
            UserControl.Cls
        End If
        
    End If
Done:
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case 398
            Err.Clear
            GoTo Done
        Case Else
            Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
            If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_tmrMouse_Timer" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
            GoTo Done
    End Select
End Sub

Private Sub tmrRefresh_Timer()
Dim Icn As clsIcon

'timer to refresh window titles until keyboard actions are captured

If Not m_colIcons Is Nothing Then
    If Not m_NoDraw Then
        For Each Icn In m_colIcons
            If Not Icn.Title = WindowText(Icn.hWnd) Then
                Icn.Title = WindowText(Icn.hWnd)
                Call PaintOne(Icn.hWnd)
            End If
        Next Icn
    End If
End If

Set Icn = Nothing
End Sub

Private Sub UserControl_Hide()
    ' if the usercontrol is going to hide, then we dont want to
    ' do any of the work we do. so stop it all
    If Ambient.UserMode() Then
        On Error Resume Next
        UnSubClassParentWnd Me
        ClearCollection
        UserControl.Parent.Arrange vbArrangeIcons
    End If
    
End Sub

Private Sub UserControl_Initialize()
    Set MainMenu = New clsMenuItems
    Bol_Refresh = True
    Default_Font_Colour = UserControl.ForeColor
End Sub

Private Sub UserControl_InitProperties()
    If GetParent(Me.hWnd) = 0 Then
        ' no parent.
        Err.Raise 20000, "TaskBar", "TaskBar control may be placed on MDI froms only"
    Else
        ' have a parent. setup default property values
        If Ambient.UserMode() Then
            ghWnd = GetParent(hWnd)
        End If
        m_Task_Height = 22 'GetSystemMetrics(SM_CYCAPTION)
        
        m_Style = m_def_Style
        m_ForeColor = m_def_ForeColor
        m_BackColor = m_def_BackColor
        m_RaisedBackColor = m_def_RaisedBackColor
        m_SunkenBackColor = m_def_SunkenBackColor
        m_SelectingBackColor = m_def_SelectingBackColor
        m_CoolBarSeparator = m_def_CoolBarSeparator
        m_ShowActive = m_def_ShowActive
        m_ShowTray = m_def_ShowTray
        m_ShowMenu = m_def_ShowMenu
        m_MenuCaption = m_def_MenuCaption
        m_AutoHide = m_def_AutoHide
        m_AutoHideWait = m_def_AutoHideWait
        m_AutoHideAnimate = m_def_AutoHideAnimate
        m_AutoHideAnimateFrames = m_def_AutoHideAnimateFrames
        m_MenuButtonIcon = m_def_MenuButtonIcon
        m_MenuButtonWidth = m_def_MenuButtonWidth
        m_MenuBarColor = m_def_MenuBarColor
        m_MenuBarTextColor = m_def_MenuBarTextColor
        m_MenuBarText = m_def_MenuBarText
        m_MenuHighlightColor = m_def_MenuHighlightColor
        gMenuHighlight = TranslateColor(m_MenuHighlightColor)
        m_MenuHighlightTextColor = m_def_MenuHighlightTextColor
        gMenuHighlightText = TranslateColor(m_MenuHighlightTextColor)
        m_MenuBackColor = m_def_MenuBackColor
        gMenuBackColor = TranslateColor(m_MenuBackColor)
        m_MenuForeColor = m_def_MenuForeColor
        gMenuForeColor = TranslateColor(m_MenuForeColor)
        m_NoDraw = False
        m_ClickedMain = False
    End If
End Sub

Private Sub UserControl_LostFocus()
    ' if the usercontrol loses focus, then we
    ' want to make sure that nothing has the mouseover
    ' effect still drawn.
    Dim I As Integer
    If m_colIcons Is Nothing Then Exit Sub
    For I = 0 To m_colIcons.Count
        InvalidateElement I
    Next I
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ErrorHandler
    ' mouse down, instant click (button behaviour later)
    If Not Ambient.UserMode() Then Exit Sub
    Dim oRect As RECT
    If Button = vbLeftButton Then
        ' detect which element and mark it
        ' set for refreshing (invalidate)
        ' set capture and wait for release
        m_nIndexBeingSelected = ElementFromPoint(x, y)
        m_bInsetSelected = (m_nIndexBeingSelected > 0)
        If (m_nIndexBeingSelected > 0) Then
            ' the set capture call just makes sure
            ' that we receive all the mouse events for our
            ' process (even if it is over another object)
            ' until we call the ReleaseCapture event.
            ' this way we can trap for the mouse up and
            ' mousemove after a mousedown, even if they
            ' happen over another part of the application.
            SetCapture UserControl.hWnd
            InvalidateElement m_nIndexBeingSelected
            m_ClickedMain = False
        Else
            If m_ShowMenu = True Then
                If IsPointInMainMenu(x, y) Then
                    If MainMenuRect(oRect) Then
                        InvalidateRect UserControl.hWnd, oRect, False
                        SetCapture UserControl.hWnd
                        DrawEdge UserControl.hdc, oRect, EDGE_SUNKEN, BF_RECT
                        m_ClickedMain = True
                    Else
                        m_ClickedMain = False
                    End If
                Else
                    m_ClickedMain = False
                End If
            End If
        End If
        
        
        
    End If
Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_UserControl_MouseDown" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "Button=" & Button, "Shift=" & Shift, "x=" & x, "y=" & y)
    GoTo Done
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ErrorHandler
    Dim Icn As clsIcon
    Dim I As Integer
    Dim bNewStatus As Boolean
    Dim nElPointed As Integer
    Dim rc As RECT
    Dim rcRect As RECT
    Dim bDispElement As Boolean
    Dim bDispTrayIcon As Boolean
    Dim sTrayIconTip As String
    Dim oTest As POINTAPI
    Dim iIcon As Integer
    If Not Ambient.UserMode() Then Exit Sub
    
    
    If m_nIndexBeingSelected > 0 Then
        ' moving while item pressed can change state
        ' of a button originaly pressed
        ' is the mouse still over the pressed element?
        bNewStatus = IsPointInElement(x, y, m_nIndexBeingSelected)
        ' if the pressed element was one of the task bar buttons
        ' and the mouse is not over that element anymore, then
        ' take away the mouseover effect, or put the raised
        ' edge back so it looks like a butotn again.
        If m_bInsetSelected <> bNewStatus Then
            m_bInsetSelected = bNewStatus
            InvalidateElement m_nIndexBeingSelected
        End If
    ElseIf Button = 0 Then
        ' handle the tray
        If IsPtInTray(x, y) Then
            iIcon = TrayIconFromPoint(x, y)
            If (iIcon <> -1 And m_colTrayIcons.Count > 0) Then
                If m_colTrayIcons.Item(iIcon + 1).ToolTip <> vbNullString Then
                    bDispTrayIcon = True
                    sTrayIconTip = m_colTrayIcons.Item(iIcon + 1).ToolTip
                Else
                    bDispTrayIcon = False
                End If
            Else
                bDispTrayIcon = False
            End If
        End If
        ' handle elements
        nElPointed = ElementFromPoint(x, y)
        If nElPointed > 0 Then
            bDispElement = False
            ' if we are moving around, we want to change
            ' tooltip text (original one is in use)
            '
            ' the rule is that if we enter a button where text can't
            ' fit into, we change ToolTipText property
            ' we also set m_bTooltip flag on just to rember
            ' to restore original contns later
            If ItemRect(nElPointed, rc) Then
                If rc.Left + m_cxBorder < x And rc.Right - m_cxBorder > x And _
                    rc.Top + m_cyBorder <= y And rc.Bottom - m_cyBorder > y Then
                    ' test for the text width.
                    GetTextExtentPoint32 UserControl.hdc, m_colIcons(nElPointed).Title, Len(m_colIcons(nElPointed).Title), oTest
                    bDispElement = oTest.x > rc.Right - rc.Left - ICON_WIDTH - 6
                    ' if its using the coolbar style
                    ' it needs to have the mouseover effect.
                    If m_Style = CoolBar Then
                        ' we use the iLast so we dont keep re-drawing
                        ' it helps with the flicker.
                        If m_iLast <> nElPointed Then
                            For I = 0 To m_colIcons.Count
                                If nElPointed = I Then
                                    DrawEdge UserControl.hdc, rc, EDGE_RAISED, BF_RECT
                                    ' this is used for clearing
                                    ' the edge from a button, when
                                    ' we leave the usercontrol
                                    ' with the mouse (mouseout)
                                    m_LastMouseOver = nElPointed
                                    m_iLast = nElPointed
                                Else
                                    InvalidateElement I
                                End If
                            Next I
                        End If
                    End If
                End If
            End If
            
        ElseIf m_bTooltip And Not bDispTrayIcon Then
            UserControl.Extender.ToolTipText = m_strOriginalTooltip
            m_bTooltip = False
        End If
        ' just setting the right tooltip
        If bDispElement Then
            If UserControl.Extender.ToolTipText <> m_colIcons(nElPointed).Title Then
                UserControl.Extender.ToolTipText = m_colIcons(nElPointed).Title
            End If
            m_bTooltip = True
        ElseIf bDispTrayIcon Then
            If UserControl.Extender.ToolTipText <> sTrayIconTip Then
                UserControl.Extender.ToolTipText = sTrayIconTip
            End If
            m_bTooltip = True
        ElseIf IsPointInMainMenu(x, y) Then
            If UserControl.Extender.ToolTipText <> m_MenuCaption Then
                UserControl.Extender.ToolTipText = m_MenuCaption
            End If
            m_bTooltip = True
        ElseIf m_bTooltip Then
            
            UserControl.Extender.ToolTipText = m_strOriginalTooltip
            m_bTooltip = False
        End If
    
    End If
Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_UserControl_MouseMove" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "Button=" & Button, "Shift=" & Shift, "x=" & x, "y=" & y)
    GoTo Done
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error GoTo ErrorHandler
On Error Resume Next
    Dim lLen As Long  ' length of the string retrieved
    Dim I As Integer
    Dim nNewActive As Integer
    Dim iIcon As Integer
    Dim oRect As RECT
    If Not Ambient.UserMode() Then Exit Sub
    
    'if in element we want - ok
    If m_nIndexBeingSelected > 0 And Button = vbLeftButton Then
        ' A user has released the mouse button while
        ' trying to push a button
        ' so we release the mouse capture.
        ReleaseCapture
        If IsPointInElement(x, y, m_nIndexBeingSelected) Then
            ' still inside a button from mouse down handler
            ' so change selection
            For I = 0 To m_colIcons.Count
                InvalidateElement I
            Next I
            'Restore the Curently active window first.
            If m_refActive.State = vbMaximized Then
                'SendMessage ghWnd, WM_MDIRESTORE, m_refActive.hWnd, 0&
                ' since the button has been pushed, activate the window.
                ActivateWindow m_nIndexBeingSelected
                'SendMessage ghWnd, WM_MDIMAXIMIZE, m_refActive.hWnd, 0&
            
            Else
                ' since the button has been pushed, activate the window.
                ActivateWindow m_nIndexBeingSelected
            End If
        End If
        If m_bInsetSelected Then InvalidateElement m_nIndexBeingSelected
        m_nIndexBeingSelected = 0
        m_bInsetSelected = False
    ElseIf m_ShowMenu And vbLeftButton And m_ClickedMain Then
        ReleaseCapture
        'If IsPointInMainMenu(x, y) Then
            If MainMenuRect(oRect) Then
                ShowMainMenu
                InvalidateRect UserControl.hWnd, oRect, False
                DrawEdge UserControl.hdc, oRect, EDGE_RAISED, BF_RECT
            End If
        'End If
        m_ClickedMain = False
    ElseIf Button = vbRightButton Then
        ' raise the menu for the child, if needed
        On Error Resume Next
        ' get the element of the right mouse click.
        nNewActive = ElementFromPoint(x, y)
        If 0 < nNewActive Then
            ' show the menu
            ShowSystemMenu m_colIcons(nNewActive).hWnd
        ElseIf IsPtInTray(x, y) Then
            ' maybe its in the tray, see and raise the event
            iIcon = TrayIconFromPoint(x, y)
            If iIcon <> -1 Then
                RaiseEvent TrayIconClick(Button, iIcon, m_colTrayIcons.Item(iIcon + 1).Key, m_colTrayIcons.Item(iIcon + 1).ToolTip)
            End If
        End If
    ElseIf Button = vbLeftButton And IsPtInTray(x, y) Then
        iIcon = TrayIconFromPoint(x, y)
        If iIcon <> -1 Then
            RaiseEvent TrayIconClick(Button, iIcon, m_colTrayIcons.Item(iIcon + 1).Key, m_colTrayIcons.Item(iIcon + 1).ToolTip)
        End If
    End If
Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_UserControl_MouseUp" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "Button=" & Button, "Shift=" & Shift, "x=" & x, "y=" & y)
    GoTo Done
End Sub

Private Sub UserControl_Paint()
On Error GoTo ErrorHandler
    ' painting whole area
    
    Dim I As Integer
    Dim rcItem As RECT
    Dim Icn As clsIcon
    Dim lEdgeParam As Long
    Dim hBrush As Long  ' receives handle to the blue hatched brush to use
    Dim r As RECT  ' rectangular area to fill
    Dim lRet As Long  ' return value
    Dim nDiff As Single
    Dim nIconTop As Single
    Dim rcIcon As RECT
    Dim lpDrawTextParams As DRAWTEXTPARAMS
    Dim nTextH As Single
    Dim oTest As POINTAPI
    Dim oTrayPoint As POINTAPI
    Dim vEdge As Variant
    Dim oTrayIcon As clsTrayIcon
    Dim bRet As Boolean
    
    If Not Ambient.UserMode() Then
        Exit Sub
    End If
    
    'Stop the Flash timer
    FlashTimer.Enabled = False
    
    If m_colIcons Is Nothing Or m_NoDraw = True Then
        ' no buttons, clear the control
        UserControl.Cls
        Exit Sub
    End If
    
    I = 0
    
    ' set the colors
    UserControl.BackColor = m_BackColor
    UserControl.ForeColor = m_ForeColor
    
    lpDrawTextParams.iLeftMargin = 1
    lpDrawTextParams.iRightMargin = 1
    lpDrawTextParams.iTabLength = 2
    lpDrawTextParams.cbSize = 20
    
    If m_ShowMenu = True Then
        If MainMenuRect(rcItem) Then
            If m_ClickedMain Then
                DrawEdge UserControl.hdc, rcItem, EDGE_SUNKEN, BF_RECT
            Else
                DrawEdge UserControl.hdc, rcItem, EDGE_RAISED, BF_RECT
            End If
        End If
        ' draw the text
        rcItem.Top = (rcItem.Bottom - rcItem.Top) / 4
        If m_MenuButtonIcon <> -1 Then
            If hImageList <> 0 Then
                ImageList_DrawEx hImageList, m_MenuButtonIcon, UserControl.hdc, rcItem.Left + 5, rcItem.Top - 1, 16, 16, CLR_NONE, CLR_DEFAULT, ILD_TRANSPARENT
                ' has an icon. so add space for it.
                rcItem.Left = rcItem.Left + ICON_WIDTH
            End If
        End If
        rcItem.Right = rcItem.Right - 2
        bRet = UserControl.FontBold
        UserControl.FontBold = True
        DrawTextEx UserControl.hdc, m_MenuCaption, Len(m_MenuCaption), rcItem, _
        DT_CENTER Or DT_VCENTER Or DT_WORD_ELLIPSIS, lpDrawTextParams
        UserControl.FontBold = bRet
    End If
    
    For Each Icn In m_colIcons
        If ItemRect(I + 1, rcItem) Then
            ' three states of a push button
            If Icn Is m_refActive Then
                
                ' down, currently selected button
                If m_ShowActive Then
                    ' they want to see the active button
                    vEdge = EDGE_SUNKEN
                    UserControl.FontBold = True
                    hBrush = CreateSolidBrush(TranslateColor(Icn.Selected_BackColour))
                    ' they dont want to see the active button
                    ' differently, draw it with the same code
                    ' that draws the inactvie buttons.
                
                Else
                    If Icn.IsFlashing And Icn.FlashOn Then
                        hBrush = CreateSolidBrush(TranslateColor(Icn.FlashColour))
                    Else
                        hBrush = CreateSolidBrush(TranslateColor(Icn.Unselected_BackColour))
                    End If
                    If m_Style = Default Then
                        ' if it is default then it looks like a raised button
                        vEdge = EDGE_RAISED
                    End If
                    UserControl.FontBold = False
                End If
                
            ElseIf I + 1 = m_nIndexBeingSelected And m_bInsetSelected Then
                ' being pushed at the moment
                vEdge = EDGE_SUNKEN
                UserControl.FontBold = False
                hBrush = CreateSolidBrush(TranslateColor(m_SelectingBackColor))
                
            Else
                ' up, inactive
                If m_Style = Default Then
                    vEdge = EDGE_RAISED
                Else
                    vEdge = 0
                End If
                UserControl.FontBold = False
                
                If Icn.IsFlashing Then
                    hBrush = CreateSolidBrush(TranslateColor(Icn.FlashColour))
                Else
                    hBrush = CreateSolidBrush(TranslateColor(Icn.Unselected_BackColour))
                End If
                
            End If
            
            ' draw the edge.
            If vEdge <> 0 Then
                DrawEdge UserControl.hdc, rcItem, vEdge, BF_RECT
            End If
            
            ' we selected the right hBrush up above, now use it to
            ' draw the back color.
            lRet = CopyRect(r, rcItem)
            r.Bottom = r.Bottom - 2
            r.Top = r.Top + 1
            r.Left = r.Left + 1
            r.Right = r.Right - 2
            lRet = FillRect(UserControl.hdc, r, hBrush)  ' fill the rectangle using the brush
            lRet = DeleteObject(hBrush) ' clean up
                
            ' fix the rect to fit inside the new border thats
            ' been drawn
            rcItem.Left = rcItem.Left + m_cxBorder + 1
            rcItem.Top = rcItem.Top + m_cyBorder
            rcItem.Right = rcItem.Right - m_cxBorder - 1
            rcItem.Bottom = rcItem.Bottom - m_cyBorder - 1
            
            ' used to calculate the position to draw the icon
            nDiff = rcItem.Bottom - rcItem.Top
            
            ' draw the icon
            ' calculate the position to draw the icon
            nIconTop = rcItem.Top + (nDiff - ICON_WIDTH) \ 2
            If Icn.IconPtr <> 0 Then
                DrawIconEx UserControl.hdc, rcItem.Left, nIconTop + 2, Icn.IconPtr, 16, 16, 0, 0, DI_NORMAL
            Else
                ' no icon was returned, so we cant draw anything.
            End If
            
            ' drawing a text with default font for a control
            If Icn.IconPtr <> 0 Then
                ' has an icon. so add space for it.
                rcItem.Left = rcItem.Left + ICON_WIDTH + 2
            Else
                ' no icon, so draw it over to the left.
                rcItem.Left = rcItem.Left + 2
            End If
            
            lpDrawTextParams.iLeftMargin = 1
            lpDrawTextParams.iRightMargin = 1
            lpDrawTextParams.iTabLength = 2
            lpDrawTextParams.cbSize = 20
            
            ' calculate all the dimensions for the rect to
            ' draw the text in.
            GetTextExtentPoint32 UserControl.hdc, Icn.Title, Len(Icn.Title), oTest
            nTextH = oTest.y
            If nTextH < nDiff Then
                nDiff = (nDiff - nTextH) \ 2
                rcItem.Bottom = rcItem.Bottom - nDiff
                rcItem.Top = rcItem.Top + nDiff
            End If
            
            'Set the Font Style.
            Set UserControl.Font = m_FontStyle
            
            If Icn.Change_Font_Colour Then
               UserControl.ForeColor = Icn.FontColour
               
            Else
               UserControl.ForeColor = Default_Font_Colour
               
            End If
                
            ' draw the text
            DrawTextEx UserControl.hdc, Icn.Title, Len(Icn.Title), rcItem, _
            DT_LEFT Or DT_VCENTER Or DT_WORD_ELLIPSIS, lpDrawTextParams
            
            
            'Set the default colour back to normal
            If Icn.Change_Font_Colour Then
               UserControl.ForeColor = Default_Font_Colour
                 
            End If
        End If
        
        I = I + 1
        ' now if it is CoolBar style, we draw the separator
        ' line inbetween buttons
        If m_colIcons.Count > 1 And I < m_colIcons.Count Then
            If m_Style = CoolBar And m_CoolBarSeparator = True Then
                If m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
                    ' draw 2 lines to give it the 3d look.
                    UserControl.Line (r.Right + 3, r.Top)-(r.Right + 3, r.Bottom), vb3DShadow
                    UserControl.Line (r.Right + 4, r.Top + 1)-(r.Right + 4, r.Bottom + 1), TranslateColor(m_SunkenBackColor)
                ElseIf m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
                    UserControl.Line (r.Left, r.Bottom + 3)-(r.Right, r.Bottom + 3), vb3DShadow
                    UserControl.Line (r.Left, r.Bottom + 4)-(r.Right, r.Bottom + 4), TranslateColor(m_SunkenBackColor)
                End If
            End If
        End If
    Next
    
    ' draw the tray
    If m_ShowTray And (Not m_colTrayIcons Is Nothing) Then
        TrayRect rcItem
        DrawEdge UserControl.hdc, rcItem, BDR_SUNKENOUTER, BF_RECT
        I = -1
        For Each oTrayIcon In m_colTrayIcons
            I = I + 1
            If TrayIconPoint(I, oTrayPoint) Then
                If hImageList <> 0 Then
                    ImageList_DrawEx hImageList, oTrayIcon.Icon, UserControl.hdc, oTrayPoint.x, oTrayPoint.y, 16, 16, CLR_NONE, CLR_DEFAULT, ILD_TRANSPARENT
                End If
            End If
        Next
        Set oTrayIcon = Nothing
    End If
    
Done:
    DeleteObject hBrush
    
    
    'Stop the Flash timer
    FlashTimer.Enabled = True


    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_UserControl_Paint" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set m_FontStyle = PropBag.ReadProperty("ButtonFont", UserControl.Font)
    
    m_Task_Height = PropBag.ReadProperty("ButtonHeight", 21)
    m_MenuBarText = PropBag.ReadProperty("MenuBarText", m_def_MenuBarText)
    m_MenuForeColor = PropBag.ReadProperty("MenuForeColor", m_def_MenuForeColor)
    gMenuForeColor = TranslateColor(m_MenuForeColor)
    m_MenuBackColor = PropBag.ReadProperty("MenuBackColor", m_def_MenuBackColor)
    gMenuBackColor = TranslateColor(m_MenuBackColor)
    m_MenuHighlightColor = PropBag.ReadProperty("MenuHighlightColor", m_def_MenuHighlightColor)
    gMenuHighlight = TranslateColor(m_MenuHighlightColor)
    m_MenuHighlightTextColor = PropBag.ReadProperty("MenuHighlightTextColor", m_def_MenuHighlightTextColor)
    gMenuHighlightText = TranslateColor(m_MenuHighlightTextColor)
    m_MenuBarTextColor = PropBag.ReadProperty("MenuBarTextColor", m_def_MenuBarTextColor)
    m_MenuBarColor = PropBag.ReadProperty("MenuBarColor", m_def_MenuBarColor)
    m_MenuButtonWidth = PropBag.ReadProperty("MenuButtonWidth", m_def_MenuButtonWidth)
    m_MenuButtonIcon = PropBag.ReadProperty("MenuButtonIcon", m_def_MenuButtonIcon)
    m_AutoHideAnimateFrames = PropBag.ReadProperty("AutoHideAnimateFrames", m_def_AutoHideAnimateFrames)
    m_AutoHideAnimate = PropBag.ReadProperty("AutoHideAnimate", m_def_AutoHideAnimate)
    m_AutoHideWait = PropBag.ReadProperty("AutoHideWait", m_def_AutoHideWait)
    m_AutoHide = PropBag.ReadProperty("AutoHide", m_def_AutoHide)
    m_ShowActive = PropBag.ReadProperty("ShowActive", m_def_ShowActive)
    m_ShowTray = PropBag.ReadProperty("ShowTray", m_def_ShowTray)
    m_ShowMenu = PropBag.ReadProperty("ShowMenu", m_def_ShowMenu)
    m_MenuCaption = PropBag.ReadProperty("MenuCaption", m_def_MenuCaption)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_CoolBarSeparator = PropBag.ReadProperty("CoolBarSeparator", m_def_CoolBarSeparator)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_RaisedBackColor = PropBag.ReadProperty("RaisedBackColor", m_def_RaisedBackColor)
    m_SunkenBackColor = PropBag.ReadProperty("SunkenBackColor", m_def_SunkenBackColor)
    m_SelectingBackColor = PropBag.ReadProperty("SelectingBackColor", m_def_SelectingBackColor)
End Sub

Private Sub UserControl_Resize()
    UserControl.Cls
    UserControl_Paint
End Sub

Private Sub UserControl_Terminate()
On Error Resume Next

    If m_Menu <> 0 Then
        ' destroy the menu
        DestroyMenu m_Menu
    End If
    FreeMenus
    tmrRefresh.Enabled = False
    tmrMouse.Enabled = False
    FlashTimer.Enabled = False
    Set m_colTrayIcons = Nothing
    Set m_colIcons = Nothing
    Set MainMenu = Nothing
    Set m_refActive = Nothing
    Set m_FontStyle = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ButtonHeight", m_Task_Height, 21
    PropBag.WriteProperty "MenuBarText", m_MenuBarText, m_def_MenuBarText
    PropBag.WriteProperty "MenuBackColor", m_MenuBackColor, m_def_MenuBackColor
    PropBag.WriteProperty "MenuForeColor", m_MenuForeColor, m_def_MenuForeColor
    PropBag.WriteProperty "MenuHighlightColor", m_MenuHighlightColor, m_def_MenuHighlightColor
    PropBag.WriteProperty "MenuHighlightTextColor", m_MenuHighlightTextColor, m_def_MenuHighlightTextColor
    PropBag.WriteProperty "MenuBarTextColor", m_MenuBarTextColor, m_def_MenuBarTextColor
    PropBag.WriteProperty "MenuBarColor", m_MenuBarColor, m_def_MenuBarColor
    PropBag.WriteProperty "MenuButtonWidth", m_MenuButtonWidth, m_def_MenuButtonWidth
    PropBag.WriteProperty "MenuButtonIcon", m_MenuButtonIcon, m_def_MenuButtonIcon
    PropBag.WriteProperty "AutoHideAnimateFrames", m_AutoHideAnimateFrames, m_def_AutoHideAnimateFrames
    PropBag.WriteProperty "AutoHideAnimate", m_AutoHideAnimate, m_def_AutoHideAnimate
    PropBag.WriteProperty "AutoHideWait", m_AutoHideWait, m_def_AutoHideWait
    PropBag.WriteProperty "AutoHide", m_AutoHide, m_def_AutoHide
    PropBag.WriteProperty "CoolBarSeparator", m_CoolBarSeparator, m_def_CoolBarSeparator
    PropBag.WriteProperty "ShowActive", m_ShowActive, m_def_ShowActive
    PropBag.WriteProperty "ShowTray", m_ShowTray, m_def_ShowTray
    PropBag.WriteProperty "ShowMenu", m_ShowMenu, m_def_ShowMenu
    PropBag.WriteProperty "MenuCaption", m_MenuCaption, m_def_MenuCaption
    PropBag.WriteProperty "Style", m_Style, m_def_Style
    PropBag.WriteProperty "ForeColor", m_ForeColor, m_def_ForeColor
    PropBag.WriteProperty "BackColor", m_BackColor, m_def_BackColor
    PropBag.WriteProperty "RaisedBackColor", m_RaisedBackColor, m_def_RaisedBackColor
    PropBag.WriteProperty "SunkenBackColor", m_SunkenBackColor, m_def_SunkenBackColor
    PropBag.WriteProperty "SelectingBackColor", m_SelectingBackColor, m_def_SelectingBackColor
    PropBag.WriteProperty "ButtonFont", m_FontStyle, UserControl.Font

End Sub

Private Sub UserControl_Show()
On Error GoTo ErrorHandler
    ' showing the usercontrol, set it all up
    If Ambient.UserMode() Then
        SubClassParentWnd Me, UserControl.hWnd
        If UserControl.Extender.ToolTipText = vbNullString Then
            UserControl.Extender.ToolTipText = UserControl.Parent.Caption
        End If
        m_strOriginalTooltip = UserControl.Extender.ToolTipText
        
        OnRefresh
    End If
    ' get all the right heights and border widths
    m_cxBorder = GetSystemMetrics(SM_CXEDGE)
   ' m_cyBorder = GetSystemMetrics(SM_CYEDGE)
    m_nOptimalHeight = m_Task_Height
    m_nOptimalHeight = m_nOptimalHeight + 2 * m_cyBorder + 3
    
    ' set the correct height
    m_nAlign = UserControl.Extender.Align
    If m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
        UserControl.Extender.Height = ScaleY(m_nOptimalHeight, vbPixels, vbTwips)
        m_ActualHeight = UserControl.Extender.Height
    ElseIf m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
        m_ActualWidth = UserControl.Extender.Width
    End If
Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_UserControl_Show" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Sub

Private Sub UpdateIconsCollection(ByVal hActive As Long)
On Error GoTo ErrorHandler
    ' updates icons collection
    
    ' this function goes through every child window
    ' of the MDIClient window, and if it is a MDI Child
    ' window, it calls the VerifyItem function on it.
    ' VerifyItem just updates the information we have
    ' for a window, and adds any new windows.
    
    Dim hWnd As Long
    Dim hWndStart As Long
    Dim sClassName As String  ' receives the name of the class
    Dim lLen As Long  ' length of the string retrieved
    Dim Icn As clsIcon
    Dim iColIdx As Integer
    
    InitCollection 'Set all as old or deleted,and rely on the Loop to find all windows.
    
    hWnd = FindWindowEx(ghWnd, 0, vbNullString, vbNullString)
    hWndStart = hWnd
    Dim iCount As Integer
    Dim nCount As Integer
    ' loop through all the child windows
    Do Until hWnd = 0
        If IsChild(ghWnd, hWnd) Then
            sClassName = Space$(255)
            lLen = GetClassName(hWnd, sClassName, 255)
            sClassName = Trim$(Left$(sClassName, lLen))
            If sClassName = "ThunderFormDC" Or sClassName = "ThunderRT6FormDC" Then
                ' those 2 class names are the names for MDI Child windows
                ' ThunderFormDC is the class if you are in the IDE
                ' and ThunderRT6Form is the class name if it is compiled.
                ' (no idea why microsoft did that)
                    VerifyItem hWnd
            End If
        End If
        hWnd = FindWindowEx(ghWnd, hWnd, vbNullString, vbNullString)
        If hWnd = hWndStart Then
            Exit Do
        End If
    Loop
    
    ' cleanup all the icons. removing the old/unneeded ones
    iColIdx = 1
    Set m_refActive = Nothing
    Do While iColIdx <= m_colIcons.Count
     Set Icn = m_colIcons.Item(iColIdx)
        
        If Not Icn.IsTaught Then
            m_colIcons.Remove iColIdx
        Else
            iColIdx = iColIdx + 1
        End If
     
    Loop
    
    If hActive > 0 Then
        'Now lets find the active one.
        For Each Icn In m_colIcons
            If Icn.hWnd = hActive Then
                Set m_refActive = Icn
                Exit For
            End If
        Next
        'Debug.Print hActive
    End If
    
    m_maxCount = m_colIcons.Count

        
Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_UpdateIconsCollection" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "hActive=" & hActive)
    GoTo Done
End Sub

Private Sub InitCollection()
On Error GoTo ErrorHandler
    ' initializes collection object by clearing status
    ' of all collection elements
    Dim Icn As clsIcon
    
    If m_colIcons Is Nothing Then
        Set m_colIcons = New Collection
        m_maxCount = 0
        Exit Sub
    End If
        
    ' setup the tray icons
    If m_colTrayIcons Is Nothing Then
        Set m_colTrayIcons = New Collection
    End If
    
    m_maxCount = 0
    For Each Icn In m_colIcons
        Icn.ClearTouch
    Next
    m_maxCount = m_colIcons.Count
Done:
    Set Icn = Nothing
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_InitCollection" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Sub

Private Sub MapIconCollection()
On Error GoTo ErrorHandler


    Dim Icn As clsIcon
    Dim strState As String
    Dim nElementIndex As Integer
    
    ' this function is to cause refreshing of proper parts
    ' of the control
    Static nLastPaintedCnt As Integer
    Static nLastPaintedAct As Integer
        
    If Not nLastPaintedCnt = m_colIcons.Count Then
        ' something added or removed - repaint all
        UserControl.Refresh
    Else
        ' refresh every changed element taking
        ' special care of minimized windows
        ' (we hide them actually)
        nElementIndex = 1
        For Each Icn In m_colIcons
            If Icn.IsNew Then
                If Icn.State = vbMinimized Then
                    ShowWindow Icn.hWnd, SW_HIDE
                    Icn.Touch
                    
                End If
                UserControl.Refresh
                Exit For
            ElseIf Icn.IsChanged Then
                If Icn.State = vbMinimized Then
                    ShowWindow Icn.hWnd, SW_HIDE
                    Icn.Touch
                
                End If
                InvalidateElement nElementIndex
            ElseIf Icn Is m_refActive And nElementIndex <> nLastPaintedAct Then
                InvalidateElement nLastPaintedAct
                InvalidateElement nElementIndex
                nLastPaintedAct = nElementIndex
            End If
            
            nElementIndex = nElementIndex + 1
        Next
    End If
    nLastPaintedCnt = m_colIcons.Count
Done:
    Set Icn = Nothing
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MapIconCollection" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Sub

Private Sub VerifyItem(ByVal hWnd As Long)
On Error GoTo ErrorHandler
    ' find item, update it if found
    ' and mark as touched
    
    Dim Icn As clsIcon
    Dim lWinStyle As Long
    Dim sCaption As String
    'Debug.Print "Start: ", m_colIcons.Count
    For Each Icn In m_colIcons
        If hWnd = Icn.hWnd Then
            ' get the windows caption/title
            sCaption = WindowText(hWnd)
            
            Icn.Title = sCaption
            
            ' get the WindowState of the window.
            lWinStyle = GetWindowLong(hWnd, GWL_STYLE)
            If lWinStyle And WS_MAXIMIZE Then
                Icn.State = vbMaximized
            ElseIf lWinStyle And WS_MINIMIZE Then
                Icn.State = vbMinimized
            Else
                Icn.State = vbNormal
            End If
            
            Icn.Touch
            ' we found the item, so we arent adding a new one.
            ' exit the sub before the code to add an item runs.
            ' Debug.Print "Found :", hWnd, Icn.Title
            ' Debug.Print "Finish: ", m_colIcons.Count
            Exit Sub
        End If
        'Debug.Print "Checking New:", hWnd, Icn.hWnd
    Next
    
    
    ' new element to be added
    
    If IsWindowVisible(hWnd) Then
        
        Set Icn = New clsIcon
        
        sCaption = WindowText(hWnd)
        Icn.Title = sCaption
        
        lWinStyle = GetWindowLong(hWnd, GWL_STYLE)
        If lWinStyle And WS_MAXIMIZE Then
            Icn.State = vbMaximized
        ElseIf lWinStyle And WS_MINIMIZE Then
            Icn.State = vbMinimized
        Else
            Icn.State = vbNormal
        End If
        Icn.hWnd = hWnd
        ' the IconPtr
        m_colIcons.Add Icn
    End If

Done:
    Set Icn = Nothing
    Exit Sub
    
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_VerifyItem" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "hWnd=" & hWnd)
    GoTo Done
End Sub

Private Sub InvalidateElement(ByVal nElIdx As Integer)
On Error GoTo ErrorHandler
    ' refreshes particular element
    ' (takes any border we have drawn off, so we can redraw it)
    Dim nAllCnt As Integer
    Dim lpRect As RECT
    
    If nElIdx < 1 Then Exit Sub
    
    nAllCnt = m_colIcons.Count
    If nElIdx > nAllCnt Then Exit Sub
    
    'now calculate position and call invalidate rect
    If ItemRect(nElIdx, lpRect) Then
        InvalidateRect UserControl.hWnd, lpRect, False
    End If
Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_InvalidateElement" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "nElIdx=" & nElIdx)
    GoTo Done
End Sub

Private Function TrayIconRect(ByVal Icon As Integer, ByRef oRect As RECT) As Boolean
On Error GoTo ErrorHandler

    Dim oPoint As POINTAPI
    If Not m_colTrayIcons Is Nothing Then
        If TrayIconPoint(Icon, oPoint) Then
            oRect.Left = oPoint.x
            oRect.Top = oPoint.y
            oRect.Right = oPoint.x + 18
            oRect.Bottom = oPoint.y + 18
            If oRect.Top > UserControl.ScaleTop Then oRect.Top = UserControl.ScaleTop
            If oRect.Bottom > UserControl.ScaleHeight - oRect.Top Then oRect.Bottom = UserControl.ScaleHeight - oRect.Top
            TrayIconRect = True
        Else
            TrayIconRect = False
        End If
    End If

Done:
    Exit Function
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_TrayIconRect" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "Icon=" & Icon, "oRect.Bottom=" & oRect.Bottom, "oRect.Top=" & oRect.Top, "oRect.Left=" & oRect.Left, "oRect.Right=" & oRect.Right)
    GoTo Done
End Function

Private Function TrayIconPoint(ByVal Icon As Integer, ByRef oPoint As POINTAPI) As Boolean
On Error GoTo ErrorHandler
    Dim oTray As RECT
    Dim lWidth As Long
    Dim dTemp As Double
    Dim lIcons As Long
    Dim lRow As Long
    Dim lMod As Long
    Dim lTempIcon As Long
    
    TrayRect oTray
    
    If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
        lWidth = oTray.Right - oTray.Left
        ' calculate the # of icons that fit per row
        dTemp = lWidth / 20
        lIcons = dTemp
        If lIcons > dTemp Then
            lIcons = lIcons - 1
        End If
        
        ' calculate what row to draw on
        dTemp = (Icon + 1) / lIcons
        lRow = dTemp
        If lRow < dTemp Then
            lRow = lRow + 1
        End If
        If lRow = 0 Then lRow = 1
        
        oPoint.x = (oTray.Left + 2) + ((Icon Mod lIcons) * 20)
        
        oPoint.y = (oTray.Top + 4) + ((lRow - 1) * 20)
        
        TrayIconPoint = True
    ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
        oPoint.x = (oTray.Left + 2) + (Icon * 20)
        oPoint.y = oTray.Top + 4
        TrayIconPoint = True
    End If
Done:
    Exit Function
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_TrayIconPoint" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "Icon=" & Icon, "oPoint.x=" & oPoint.x, "oPoint.y=" & oPoint.y)
    GoTo Done
End Function

Private Function TrayRect(ByRef rItem As RECT) As Boolean
On Error GoTo ErrorHandler
    ' returns true for existing (worth painting at least) tray
    ' and fills the RECT object we pass in, with the correct dimentions
    ' for the tray

    Dim nItemH As Long
    Dim nItemW As Long
    Dim lIcons As Long
    Dim lTray As Long
    Dim dTest As Double
    
    If Not m_colTrayIcons Is Nothing Then
        m_nAlign = UserControl.Extender.Align
        
        If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
            nItemH = m_nOptimalHeight - 3
            rItem.Left = UserControl.ScaleLeft + 1
            rItem.Right = UserControl.ScaleWidth - UserControl.ScaleLeft - 1
            If rItem.Right - rItem.Left > 0 Then
                'UserControl.ScaleWidth - ((20 * m_colTrayIcons.Count) + 4)
                ' get total incons width, check for remainder, fix it.
                lIcons = ((20 * m_colTrayIcons.Count) + 4)
                If lIcons > (rItem.Right - rItem.Left) Then
                    ' wider. make the tray taller
                    dTest = (rItem.Right - rItem.Left) / 20
                    lIcons = dTest
                    If lIcons > dTest Then
                        lIcons = lIcons - 1
                    End If
                    dTest = m_colTrayIcons.Count / lIcons
                    lTray = dTest
                    If lTray < dTest Then
                        ' if the return from that division is 1.1 or for that mater
                        ' 1.Anything (any remainder), then we need to add 1 to
                        ' the # of rows to make
                        lTray = lTray + 1
                    End If
                    'If lTray > 1 Then lTray = lTray + 2
                    rItem.Top = UserControl.ScaleHeight - ((m_nOptimalHeight - 3) * lTray)
                    rItem.Bottom = rItem.Top + ((m_nOptimalHeight - 3) * lTray)
                Else
                    rItem.Top = UserControl.ScaleHeight - (m_nOptimalHeight - 3)
                    rItem.Bottom = rItem.Top + (m_nOptimalHeight - 3)
                End If
                
                If rItem.Bottom > UserControl.ScaleTop + UserControl.ScaleHeight Then
                    rItem.Left = 0
                    rItem.Right = 0
                    rItem.Top = 0
                    rItem.Bottom = 0
                    TrayRect = False
                Else
                    TrayRect = True
                End If
            Else
                rItem.Left = 0
                rItem.Right = 0
                TrayRect = False
            End If
        ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
            rItem.Top = UserControl.ScaleTop + 1
            'rItem.Bottom = UserControl.ScaleHeight - UserControl.ScaleTop - 1
            rItem.Bottom = 10 - UserControl.ScaleTop - 1
            If rItem.Bottom - rItem.Top > 0 Then
                rItem.Left = UserControl.ScaleWidth - ((20 * m_colTrayIcons.Count) + 4)
                rItem.Right = rItem.Left + ((20 * m_colTrayIcons.Count) + 2)
                TrayRect = True
            Else
                rItem.Top = 0
                rItem.Bottom = 0
                TrayRect = False
            End If
        End If
    End If
Done:
    Exit Function
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_TrayRect" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "rItem.Bottom=" & rItem.Bottom, "rItem.Top=" & rItem.Top, "rItem.Left=" & rItem.Left, "rItem.Right=" & rItem.Right)
    GoTo Done
End Function

Private Function CountHiddenButtons() As Long
On Error GoTo ErrorHandler
    ' this returns the # of buttons that cant be drawn
    ' when the control is Left or Right aligned.
    ' (cant be drawn, means they fall off the end, not enough room)
    Dim lTotal As Long
    Dim dReal As Double
    dReal = m_AvailableHeight / (m_nOptimalHeight - 3)
    lTotal = dReal
    If lTotal > dReal Then
        lTotal = lTotal - 1
    End If
    lTotal = lTotal - 2
    CountHiddenButtons = IIf(lTotal >= m_colIcons.Count, 0, m_colIcons.Count - lTotal)
Done:
    Exit Function
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_CountHiddenButtons" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Function

Private Function IsPointInMainMenu(ByVal x As Long, ByVal y As Long) As Boolean

    Dim oRect As RECT
    
    If MainMenuRect(oRect) Then
        IsPointInMainMenu = CBool(PtInRect(oRect, x, y))
    End If

End Function

Private Function MainMenuRect(ByRef rItem As RECT) As Boolean
On Error GoTo ErrorHandler
    
    m_nAlign = UserControl.Extender.Align
    Dim nItem As Long
    
    If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
        nItem = m_nOptimalHeight - 3
        rItem.Left = UserControl.ScaleLeft + 1
        rItem.Right = UserControl.ScaleWidth - UserControl.ScaleLeft - 1
        rItem.Top = (FIRST_OFFSET)
        rItem.Bottom = rItem.Top + nItem
        MainMenuRect = True
    ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
        rItem.Top = UserControl.ScaleTop + 1
        rItem.Bottom = UserControl.ScaleHeight - UserControl.ScaleTop - 1
        rItem.Left = UserControl.ScaleLeft + FIRST_OFFSET
        rItem.Right = rItem.Left + (m_MenuButtonWidth - 5)
        MainMenuRect = True
    End If

Done:
    Exit Function
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MainMenuRect" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    MainMenuRect = False
    GoTo Done
End Function

Private Function ItemRect(ByVal itmIdx As Integer, ByRef rItem As RECT) As Boolean
On Error GoTo ErrorHandler
    ' returns true for existing (worth painting at least) buttons
    ' and fills the RECT object we pass in, with the correct dimentions
    ' for itmIdx
    'Debug.Assert itmIdx > 0 And itmIdx <= m_colIcons.Count
    
    Dim nItemH As Long
    Dim nItemW As Long
    Dim lTray As Long
    Dim oRect As RECT
    Dim lMenu As Long
    
    m_nAlign = UserControl.Extender.Align
    
    If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
        nItemH = m_nOptimalHeight - 3
        rItem.Left = UserControl.ScaleLeft + 1
        rItem.Right = UserControl.ScaleWidth - UserControl.ScaleLeft - 1
        lTray = 0
        lMenu = 0
        ' calculate the tray size
        If m_ShowTray = True Then
            If TrayRect(oRect) Then
                lTray = oRect.Bottom - oRect.Top
            End If
        End If
        If m_ShowMenu = True Then
            lMenu = nItemH + 2
        End If
        If rItem.Right - rItem.Left > 0 Then
            rItem.Top = (FIRST_OFFSET + lMenu) + (itmIdx - 1) * (nItemH + STANDARD_OFFSET)
            rItem.Bottom = rItem.Top + nItemH
            If rItem.Bottom > UserControl.ScaleTop + (UserControl.ScaleHeight - lTray - lMenu) Then
                rItem.Left = 0
                rItem.Right = 0
                rItem.Top = 0
                rItem.Bottom = 0
                ItemRect = False
            Else
                m_AvailableHeight = UserControl.ScaleTop + (UserControl.ScaleHeight - lTray - lMenu)
                ItemRect = True
            End If
        Else
            rItem.Left = 0
            rItem.Right = 0
            ItemRect = False
        End If
    ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
        rItem.Top = UserControl.ScaleTop + 1
        rItem.Bottom = UserControl.ScaleHeight - UserControl.ScaleTop - 1
        If rItem.Bottom - rItem.Top > 0 Then
            If m_maxCount = 0 Then Exit Function
            ' calculate the size of the tray. if there is one.
            ' and remove it from the drawable area of the bar
            ' for buttons.
            lTray = 0
            lMenu = 0
            If m_ShowTray = True Then
                If TrayRect(oRect) Then
                    lTray = oRect.Right - oRect.Left
                End If
            End If
            If m_ShowMenu = True Then
                lMenu = m_MenuButtonWidth
            End If
            nItemW = (UserControl.ScaleWidth - 3 - lTray - lMenu) \ m_maxCount - 3
            m_AvailableWidth = (UserControl.ScaleWidth - 3 - lTray - lMenu)
            nItemW = IIf(nItemW > DEFAULT_ITEM_WIDTH, DEFAULT_ITEM_WIDTH, nItemW)
            rItem.Left = UserControl.ScaleLeft + FIRST_OFFSET + (itmIdx - 1) * (nItemW + STANDARD_OFFSET)
            If lMenu > 0 Then
                rItem.Left = rItem.Left + lMenu
            End If
            rItem.Right = rItem.Left + nItemW
            ItemRect = True
        Else
            rItem.Top = 0
            rItem.Bottom = 0
            ItemRect = False
        End If
    End If
Done:
    Exit Function
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_ItemRect" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "itmIdx=" & itmIdx, "rItem.Bottom=" & rItem.Bottom, "rItem.Top=" & rItem.Top, "rItem.Left=" & rItem.Left, "rItem.Right=" & rItem.Right)
    GoTo Done
End Function

Private Function IsPtInTray(ByVal x As Long, ByVal y As Long) As Boolean
On Error GoTo ErrorHandler
    Dim oRect As RECT
    
    If TrayRect(oRect) Then
        If CBool(PtInRect(oRect, x, y)) Then
            IsPtInTray = True
        Else
            IsPtInTray = False
        End If
    Else
        IsPtInTray = False
    End If

Done:
    Exit Function
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_IsPtInTray" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "x=" & x, "y=" & y)
    GoTo Done
End Function

Private Function TrayIconFromPoint(ByVal x As Single, ByVal y As Single) As Integer
On Error GoTo ErrorHandler
    'returns index of an element, the point of coordinates
    'given is in
    
    Dim nEl As Integer
    Dim Icn As clsIcon
    Dim I As Integer
    
    If m_ShowTray Then
        ' not 0, we have items.
        m_nAlign = UserControl.Extender.Align
        ' default to 0
        TrayIconFromPoint = -1
                
        ' all this if/else does, is check to see that the x/y is within
        ' the borders of our task bar, in the drawable area. if not
        ' we dont need to do anything else. just exit, leaving the
        ' default of 0
        If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
            If Not (x > UserControl.ScaleLeft + 1 And x < UserControl.ScaleWidth - UserControl.ScaleLeft - 1) Then
                Exit Function
            End If
        ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
            If Not (y > UserControl.ScaleTop + 1 And y < UserControl.ScaleHeight - UserControl.ScaleTop - 1) Then
                Exit Function
            End If
        End If
        
        For I = 0 To m_colTrayIcons.Count - 1
            If IsPointInTrayIcon(x, y, I) Then
                TrayIconFromPoint = I
                Exit For
            End If
        Next I
    End If
    
Done:
    Set Icn = Nothing
    Exit Function
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_TrayIconFromPoint" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "x=" & x, "y=" & y)
    GoTo Done
End Function

'Private Function PointInElement(ByVal x As Single, ByVal y As Single) As Integer
'    'returns index of an element, the point of coordinates
'    'given is in
'    If m_maxCount <> 0 Then
'        m_nAlign = UserControl.Extender.Align
'        Dim nEl As Integer
'        If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
'            PointInElement = 0
'            If x > UserControl.ScaleLeft + 1 And x < UserControl.ScaleWidth - UserControl.ScaleLeft - 1 Then
'                Dim nItemH As Long
'                nItemH = m_nOptimalHeight - 3
'
'                nEl = Int((y - UserControl.ScaleTop - FIRST_OFFSET) / (nItemH + STANDARD_OFFSET)) + 1
'                If Not (nEl > m_maxCount Or nEl < 0) Then
'                    If (y - UserControl.ScaleTop - FIRST_OFFSET) - (nEl - 1) * (nItemH + STANDARD_OFFSET) > -2 Then PointInElement = nEl
'                End If
'            End If
'        ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
'            PointInElement = 0
'            If y > UserControl.ScaleTop + 1 And y < UserControl.ScaleHeight - UserControl.ScaleTop - 1 Then
'
'                Dim nItemW As Long
'
'                nItemW = (UserControl.ScaleWidth - 3) / m_maxCount - 3
'                nItemW = IIf(nItemW > DEFAULT_ITEM_WIDTH, DEFAULT_ITEM_WIDTH, nItemW)
'
'                nEl = Int((x - UserControl.ScaleLeft - FIRST_OFFSET) / (nItemW + STANDARD_OFFSET)) + 1
'                If Not (nEl > m_maxCount Or nEl < 0) Then
'                    If (x - UserControl.ScaleLeft - FIRST_OFFSET) - (nEl - 1) * (nItemW + STANDARD_OFFSET) > -2 Then PointInElement = nEl
'                End If
'
'            End If
'        End If
'    End If
'
'End Function

Private Function ElementFromPoint(ByVal x As Single, ByVal y As Single) As Integer
On Error GoTo ErrorHandler
    'returns index of an element, the point of coordinates
    'given is in
    
    ' this is my new version of the PointInElement function
    ' renamed to better suite the purpose of the function
    Dim nEl As Integer
    Dim Icn As clsIcon
    Dim I As Integer
    
    If m_maxCount <> 0 Then
        ' not 0, we have items.
        m_nAlign = UserControl.Extender.Align
        ' default to 0
        ElementFromPoint = 0
                
        ' all this if/else does, is check to see that the x/y is within
        ' the borders of our task bar, in the drawable area. if not
        ' we dont need to do anything else. just exit, leaving the
        ' default of 0
        If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
            If Not (x > UserControl.ScaleLeft + 1 And x < UserControl.ScaleWidth - UserControl.ScaleLeft - 1) Then
                Exit Function
            End If
        ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
            If Not (y > UserControl.ScaleTop + 1 And y < UserControl.ScaleHeight - UserControl.ScaleTop - 1) Then
                Exit Function
            End If
        End If
        
        For I = 1 To m_colIcons.Count
            If IsPointInElement(x, y, I) Then
                ElementFromPoint = I
                Exit For
            End If
        Next I
    End If
    
Done:
    Set Icn = Nothing
    Exit Function
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_ElementFromPoint" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "x=" & x, "y=" & y)
    GoTo Done
End Function

'Private Function IsPointInElement2(ByVal x As Single, ByVal y As Single, ByVal idx As Integer) As Boolean
'    'checks, whether the point is within area of the point
'    'of given index or not
'
'    'returns index of an element, the point of coordinates
'    'given is in
'    m_nAlign = UserControl.Extender.Align
'    Dim nEl As Integer
'    If m_maxCount <> 0 Then
'        If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
'            IsPointInElement = False
'            If x > UserControl.ScaleLeft + 1 And x < UserControl.ScaleWidth - UserControl.ScaleLeft - 1 Then
'
'
'                Dim nItemH As Long
'                nItemH = m_nOptimalHeight - 3
'
'                Dim yOffs As Single
'                yOffs = y - UserControl.ScaleTop - FIRST_OFFSET
'
'                IsPointInElement = (y > (idx - 1) * (nItemH + STANDARD_OFFSET)) And (y < idx * (nItemH + STANDARD_OFFSET) - STANDARD_OFFSET)
'            End If
'        ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
'            IsPointInElement = False
'            If y > UserControl.ScaleTop + 1 And y < UserControl.ScaleHeight - UserControl.ScaleTop - 1 Then
'
'                Dim nItemW As Long
'                nItemW = (UserControl.ScaleWidth - 3) / m_maxCount - 3
'                nItemW = IIf(nItemW > DEFAULT_ITEM_WIDTH, DEFAULT_ITEM_WIDTH, nItemW)
'
'                Dim xOffs As Single
'                xOffs = x - UserControl.ScaleLeft - FIRST_OFFSET
'
'                IsPointInElement = (x > (idx - 1) * (nItemW + STANDARD_OFFSET)) And (x < idx * (nItemW + STANDARD_OFFSET) - STANDARD_OFFSET)
'            End If
'        End If
'    End If
'End Function

Private Function IsPointInElement(ByVal x As Single, ByVal y As Single, ByVal idx As Integer) As Boolean
On Error GoTo ErrorHandler
    'checks, whether the point is within area of the point
    'of given index or not
    
    ' NEW version of this fuction, easier to read.
    
    Dim nEl As Integer
    Dim oRect As RECT
    
    m_nAlign = UserControl.Extender.Align
    ' default to False
    IsPointInElement = False
    
    If m_maxCount <> 0 Then
        ' all this if/else does, is check to see that the x/y is within
        ' the borders of our task bar, in the drawable area. if not
        ' we dont need to do anything else. just exit, leaving the
        ' default of false
        If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
            If Not (x > UserControl.ScaleLeft + 1 And x < UserControl.ScaleWidth - UserControl.ScaleLeft - 1) Then
                Exit Function
            End If
        ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
            If Not (y > UserControl.ScaleTop + 1 And y < UserControl.ScaleHeight - UserControl.ScaleTop - 1) Then
                Exit Function
            End If
        End If
        
        If ItemRect(idx, oRect) Then
            IsPointInElement = CBool(PtInRect(oRect, x, y))
        End If
    End If
Done:
    Exit Function
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_IsPointInElement" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "x=" & x, "y=" & y, "idx=" & idx)
    GoTo Done
End Function


Private Function IsPointInTrayIcon(ByVal x As Single, ByVal y As Single, ByVal idx As Integer) As Boolean
On Error GoTo ErrorHandler
    'checks, whether the point is within area of the point
    'of given index or not
    
    Dim oRect As RECT
    
    m_nAlign = UserControl.Extender.Align
    ' default to False
    IsPointInTrayIcon = False
    
    If m_ShowTray Then
        ' all this if/else does, is check to see that the x/y is within
        ' the borders of our task bar, in the drawable area. if not
        ' we dont need to do anything else. just exit, leaving the
        ' default of false
        If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
            If Not (x > UserControl.ScaleLeft + 1 And x < UserControl.ScaleWidth - UserControl.ScaleLeft - 1) Then
                Exit Function
            End If
        ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
            If Not (y > UserControl.ScaleTop + 1 And y < UserControl.ScaleHeight - UserControl.ScaleTop - 1) Then
                Exit Function
            End If
        End If
        
        If TrayIconRect(idx, oRect) Then
            IsPointInTrayIcon = CBool(PtInRect(oRect, x, y))
        End If
    End If

Done:
    Exit Function
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_IsPointInTrayIcon" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "x=" & x, "y=" & y, "idx=" & idx)
    GoTo Done
End Function
Private Sub ActivateWindow(ByVal nEl As Integer)
    ' this function is called when we need to activate
    ' a window because a button on the taskbar was clicked.
    
    ' major re-write of this function also
    ' about half the size it used to be.
    Dim hWnd As Long
    Dim hWndStart As Long
    Dim hWndLast As Long
    Dim sClassName As String
    Dim lLen As Long
    Dim Old_Active As clsIcon
    
    If nEl < 1 Or nEl > m_maxCount Then Exit Sub
    
    Set Old_Active = m_refActive
    
    Set m_refActive = m_colIcons(nEl)
    If m_refActive.State = vbMinimized Then
        SendMessage ghWnd, WM_MDIRESTORE, m_refActive.hWnd, 0&
        m_refActive.State = vbNormal
    Else
        If m_refActive.IsFlashing Then
            m_refActive.FlashCount = m_refActive.SetFlashCount
            m_refActive.FlashOn = True
        End If
        
        BringWindowToTop m_refActive.hWnd
        Bol_Refresh = False
        'If m_refActive.hWnd = Old_Active.hWnd Then 'added and removed by syntax
        '    SendMessage ghWnd, WM_MDIACTIVATE, m_refActive.hWnd, 0
        'Else
            SendMessage ghWnd, WM_MDIACTIVATE, m_refActive.hWnd, Old_Active.hWnd
        'End If
        Bol_Refresh = True
        
    End If
    Set Old_Active = Nothing
Done:
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_ActivateWindow" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "nEl=" & nEl)
    GoTo Done
End Sub

Private Sub ClearCollection()
On Error GoTo ErrorHandler
    ' this makes sure all the windows that we had hidden
    ' are shown again. Just cleaning up on a Usercontrol_Hide.
    Dim Icn As clsIcon
    
    m_nIndexBeingSelected = 0
    m_bInsetSelected = False
    If m_colIcons Is Nothing Then GoTo Done:
    For Each Icn In m_colIcons
        If Icn.State = vbMinimized Then
            ShowWindow Icn.hWnd, SW_SHOW
        End If
    Next
Done:
    Set Icn = Nothing
    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_ClearCollection" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Sets the text color for the task bar."
On Error GoTo ErrorHandler
    ForeColor = m_ForeColor
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_ForeColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
On Error GoTo ErrorHandler
    m_ForeColor = vNewValue
    UserControl_Paint
    PropertyChanged "ForeColor"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_ForeColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Sets the BackColor for the task bar (not the buttons)"
On Error GoTo ErrorHandler
    BackColor = m_BackColor
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_BackColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
On Error GoTo ErrorHandler
    m_BackColor = vNewValue
    UserControl_Paint
    PropertyChanged "BackColor"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_BackColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Property Get MenuBarColor() As OLE_COLOR
Attribute MenuBarColor.VB_Description = "This sets the color of the bar that goes up the left side of our main menu."
On Error GoTo ErrorHandler
    MenuBarColor = m_MenuBarColor
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuBarColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let MenuBarColor(ByVal vNewValue As OLE_COLOR)
On Error GoTo ErrorHandler
    m_MenuBarColor = vNewValue
    PropertyChanged "MenuBarColor"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuBarColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Property Get MenuBarTextColor() As OLE_COLOR
Attribute MenuBarTextColor.VB_Description = "This sets the color of the text that gets drawn on the MenuBar (MenuBarText property.)"
On Error GoTo ErrorHandler
    MenuBarTextColor = m_MenuBarTextColor
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuBarTextColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let MenuBarTextColor(ByVal vNewValue As OLE_COLOR)
On Error GoTo ErrorHandler
    m_MenuBarTextColor = vNewValue
    PropertyChanged "MenuBarTextColor"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuBarTextColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Property Get RaisedBackColor() As OLE_COLOR
Attribute RaisedBackColor.VB_Description = "This sets the BackColor of items on the bar that are currently in a raised state."
On Error GoTo ErrorHandler
    RaisedBackColor = m_RaisedBackColor
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_RaisedBackColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let RaisedBackColor(ByVal vNewValue As OLE_COLOR)
On Error GoTo ErrorHandler
    m_RaisedBackColor = vNewValue
    UserControl_Paint
    PropertyChanged "RaisedBackColor"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_RaisedBackColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Property Get SunkenBackColor() As OLE_COLOR
Attribute SunkenBackColor.VB_Description = "Sets the BackColor of the button that is currently active or Sunken."
On Error GoTo ErrorHandler
    SunkenBackColor = m_SunkenBackColor
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_SunkenBackColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let SunkenBackColor(ByVal vNewValue As OLE_COLOR)
On Error GoTo ErrorHandler
    m_SunkenBackColor = vNewValue
    UserControl_Paint
    PropertyChanged "SunkenBackColor"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_SunkenBackColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Property Get SelectingBackColor() As OLE_COLOR
Attribute SelectingBackColor.VB_Description = "Sets the BackColor of buttons that are currently being selected (clicked/held down with the mouse.)"
On Error GoTo ErrorHandler
    SelectingBackColor = m_SelectingBackColor
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_SelectingBackColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property
'
Public Property Let SelectingBackColor(ByVal vNewValue As OLE_COLOR)
On Error GoTo ErrorHandler
    m_SelectingBackColor = vNewValue
    UserControl_Paint
    PropertyChanged "SelectingBackColor"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_SelectingBackColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Property Get Style() As enmStyles
Attribute Style.VB_Description = "Sets the style to use when drawing the task bar."
On Error GoTo ErrorHandler
    Style = m_Style
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_Style" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let Style(ByVal vNewValue As enmStyles)
On Error GoTo ErrorHandler
    m_Style = vNewValue
    PropertyChanged "Style"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_Style" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Property Get CoolBarSeparator() As Boolean
Attribute CoolBarSeparator.VB_Description = "True/False Show a separator bar inbetween the buttons when Style is set to CoolBar."
On Error GoTo ErrorHandler
    CoolBarSeparator = m_CoolBarSeparator
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_CoolBarSeparator" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let CoolBarSeparator(ByVal vNewValue As Boolean)
On Error GoTo ErrorHandler
    m_CoolBarSeparator = vNewValue
    PropertyChanged "CoolBarSeparator"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_CoolBarSeparator" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Property Get ShowActive() As Boolean
Attribute ShowActive.VB_Description = "This tells the control whether or not to show active items on the bar. If this is set to False then all the items, be them active or not, will look the same."
On Error GoTo ErrorHandler
    ShowActive = m_ShowActive
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_ShowActive" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let ShowActive(ByVal vNewValue As Boolean)
On Error GoTo ErrorHandler
    m_ShowActive = vNewValue
    PropertyChanged "ShowActive"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_ShowActive" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Property Get AutoHide() As Boolean
Attribute AutoHide.VB_Description = "AutoHide works just like windows taskbar's autohide feature.  When you move the mouse off of the taskbar it hides (after the specified time in AutoHideWait has passed), and when you move back over the smaller version of the bar, it shows again."
On Error GoTo ErrorHandler
    AutoHide = m_AutoHide
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_AutoHide" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let AutoHide(ByVal vNewValue As Boolean)
On Error GoTo ErrorHandler
    m_AutoHide = vNewValue
    m_NoDraw = False    '//added by syntax to fix problem after un-auto-hiding
                        '//and the taskicons not showing
    PropertyChanged "AutoHide"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_AutoHide" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Property Get ButtonHeight() As Integer
On Error GoTo ErrorHandler
    ButtonHeight = m_Task_Height
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_AutoHide" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let ButtonHeight(ByVal Height As Integer)
On Error GoTo ErrorHandler
    m_Task_Height = Height
    
    PropertyChanged "ButtonHeight"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_AutoHide" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & Height)
    GoTo Done
End Property


Public Property Get AutoHideWait() As Integer
Attribute AutoHideWait.VB_Description = "This is the # of milliseconds (1000 = 1 second) to wait until hiding the bar, after moving off of it when AutoHide = True."
On Error GoTo ErrorHandler
    AutoHideWait = m_AutoHideWait
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_AutoHideWait" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let AutoHideWait(ByVal vNewValue As Integer)
On Error GoTo ErrorHandler
    m_AutoHideWait = vNewValue
    PropertyChanged "AutoHideWait"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_AutoHideWait" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Property Get AutoHideAnimate() As Boolean
Attribute AutoHideAnimate.VB_Description = "If AutoHide is true, and you would like the bar to Slide out of side, instead of just disappearing, then set this to true."
On Error GoTo ErrorHandler
    AutoHideAnimate = m_AutoHideAnimate
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_AutoHideAnimate" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let AutoHideAnimate(ByVal vNewValue As Boolean)
On Error GoTo ErrorHandler
    m_AutoHideAnimate = vNewValue
    PropertyChanged "AutoHideAnimate"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_AutoHideAnimate" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Property Get AutoHideAnimateFrames() As Integer
Attribute AutoHideAnimateFrames.VB_Description = "This sets how many frames to use when animating the AutoHide feature."
On Error GoTo ErrorHandler
    AutoHideAnimateFrames = m_AutoHideAnimateFrames
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_AutoHideAnimateFrames" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let AutoHideAnimateFrames(ByVal vNewValue As Integer)
On Error GoTo ErrorHandler
    m_AutoHideAnimateFrames = vNewValue
    PropertyChanged "AutoHideAnimateFrames"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_AutoHideAnimateFrames" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Property Get ShowTray() As Boolean
Attribute ShowTray.VB_Description = "This tells the bar whether or not to show the tray."
On Error GoTo ErrorHandler
    ShowTray = m_ShowTray
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_ShowTray" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let ShowTray(ByVal vNewValue As Boolean)
On Error GoTo ErrorHandler
    m_ShowTray = vNewValue
    PropertyChanged "ShowTray"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_ShowTray" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Property Get ShowMenu() As Boolean
Attribute ShowMenu.VB_Description = "This tells the control whether to show the menu button  or not."
On Error GoTo ErrorHandler
    ShowMenu = m_ShowMenu
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_ShowMenu" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let ShowMenu(ByVal bNewValue As Boolean)
On Error GoTo ErrorHandler
    m_ShowMenu = bNewValue
    PropertyChanged "ShowMenu"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_ShowMenu" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "bNewValue=" & bNewValue)
    GoTo Done
End Property

Public Property Get MenuCaption() As String
Attribute MenuCaption.VB_Description = "This sets the Caption text for the menu button."
On Error GoTo ErrorHandler
    MenuCaption = m_MenuCaption
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuCaption" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let MenuCaption(ByVal sNewValue As String)
On Error GoTo ErrorHandler
    m_MenuCaption = sNewValue
    PropertyChanged "MenuCaption"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuCaption" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "sNewValue=" & sNewValue)
    GoTo Done
End Property

Public Property Get MenuButtonIcon() As Long
Attribute MenuButtonIcon.VB_Description = "This holds the index in the ImageList that was set as the hImageList property, of the icon to draw on the menu button."
Attribute MenuButtonIcon.VB_MemberFlags = "400"
On Error GoTo ErrorHandler
    MenuButtonIcon = m_MenuButtonIcon
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuButtonIcon" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let MenuButtonIcon(ByVal lNewValue As Long)
On Error GoTo ErrorHandler
    m_MenuButtonIcon = lNewValue
    PropertyChanged "MenuButtonIcon"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuButtonIcon" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "lNewValue=" & lNewValue)
    GoTo Done
End Property

Public Property Get MenuButtonWidth() As Long
Attribute MenuButtonWidth.VB_Description = "This sets the minimum width for the Menu Button."
On Error GoTo ErrorHandler
    MenuButtonWidth = m_MenuButtonWidth
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuButtonWidth" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let MenuButtonWidth(ByVal lNewValue As Long)
On Error GoTo ErrorHandler
    If lNewValue > 50 Then
        m_MenuButtonWidth = lNewValue
        PropertyChanged "MenuButtonWidth"
    Else
        MsgBox "Value '" & CStr(lNewValue) & "' is to low, the lowest this property can be set at is 50.", vbInformation, "Error"
    End If
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuButtonWidth" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "lNewValue=" & lNewValue)
    GoTo Done
End Property

Public Property Get MenuBarText() As String
Attribute MenuBarText.VB_Description = "This sets the text to be drawn on the MenuBar that goes up the left side of our main menu."
On Error GoTo ErrorHandler
    MenuBarText = m_MenuBarText
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuBarText" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let MenuBarText(ByVal sNewValue As String)
On Error GoTo ErrorHandler
    If Len(sNewValue) > 0 Then
        m_MenuBarText = sNewValue
        PropertyChanged "MenuBarText"
    End If
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuBarText" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "sNewValue=" & sNewValue)
    GoTo Done
End Property

Public Property Get hImageList() As Long
Attribute hImageList.VB_Description = "This is the imagelist to use for all functions requiring images (tray icons, menu items, ect...)"
Attribute hImageList.VB_MemberFlags = "400"
On Error GoTo ErrorHandler
    hImageList = m_hImageList
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_hImageList" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let hImageList(ByVal vNewValue As Long)
On Error GoTo ErrorHandler
    m_hImageList = vNewValue
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_hImageList" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "vNewValue=" & vNewValue)
    GoTo Done
End Property

Public Function CreateSubMenu(ByVal hMenu As Long, ByRef MenuItem As clsMenuItem, ByVal iCurrent As Long)
Attribute CreateSubMenu.VB_MemberFlags = "40"


    Dim hSubMenu As Long
    Dim oSubMenuItem As clsMenuItem
    Dim iCount As Long
    m_Main = m_Main + 1
    hSubMenu = CreatePopupMenu()
    mID = mID + 1
    AppendMenu hMenu, MF_POPUP, hSubMenu, MenuItem.Caption
    RegisterMenu hMenu, iCurrent, UserControl.hWnd, MenuItem.Caption, MenuItem.Icon, m_hImageList, mID + IIf(m_Main = 1, 9001, 0), MenuItem.Key, MenuItem.Style, m_MenuBarText, MenuItem.Tag, (MenuItem.SubItems.Count > 0)
    
    iCount = -1
    For Each oSubMenuItem In MenuItem.SubItems
        iCount = iCount + 1
        If oSubMenuItem.SubItems.Count > 0 Then
            CreateSubMenu hSubMenu, oSubMenuItem, iCount
        Else
            mID = mID + 1
            If oSubMenuItem.Style = 0 Then
                AppendMenu hSubMenu, MF_STRING, mID, oSubMenuItem.Caption
                RegisterMenu hSubMenu, iCount, UserControl.hWnd, oSubMenuItem.Caption, oSubMenuItem.Icon, m_hImageList, mID, oSubMenuItem.Key, oSubMenuItem.Style, m_MenuBarText, oSubMenuItem.Tag, (oSubMenuItem.SubItems.Count > 0)
            Else
                AppendMenu hSubMenu, MF_SEPARATOR, ByVal 0&, ByVal 0&
                RegisterMenu hSubMenu, iCount, UserControl.hWnd, "", 0, 0, mID, oSubMenuItem.Key, oSubMenuItem.Style, m_MenuBarText, oSubMenuItem.Tag, (oSubMenuItem.SubItems.Count > 0)
            End If
        End If
    Next oSubMenuItem
    
    Set oSubMenuItem = Nothing
End Function

Public Function BuildMenu()
Attribute BuildMenu.VB_Description = "After modifying the MainMenu object, you will need to call BuildMenu to actuall create the menu."
On Error Resume Next
    mID = 0
    Dim hMenu As Long
    Dim oMenuItem As clsMenuItem
    Dim iCount As Long
    m_Main = 0
    If MainMenu.Count > 0 Then
        If m_Menu <> 0 Then
            ' destroy the old menu if we've already run
            DestroyMenu m_Menu
        End If
        FreeMenus
        iCount = -1
        hMenu = CreatePopupMenu()
        For Each oMenuItem In MainMenu
            iCount = iCount + 1
            If oMenuItem.SubItems.Count > 0 Then
                CreateSubMenu hMenu, oMenuItem, iCount
                m_Main = 0
            Else
                m_Main = 0
                mID = mID + 1
                If oMenuItem.Style = 0 Then
                    ' Add the separator to the end of the system menu.
                    AppendMenu hMenu, MF_STRING, mID + 9001, oMenuItem.Caption
                    RegisterMenu hMenu, iCount, UserControl.hWnd, oMenuItem.Caption, oMenuItem.Icon, m_hImageList, mID + 9001, oMenuItem.Key, oMenuItem.Style, m_MenuBarText, oMenuItem.Tag, (oMenuItem.SubItems.Count > 0)
                Else
                    AppendMenu hMenu, MF_SEPARATOR, mID + 9001, ByVal 0&
                    RegisterMenu hMenu, iCount, UserControl.hWnd, "", 0, 0, mID + 9001, oMenuItem.Key, oMenuItem.Style, m_MenuBarText, oMenuItem.Tag, (oMenuItem.SubItems.Count > 0)
                End If
            End If
        Next oMenuItem
    End If
    
    m_Menu = hMenu
    
    Set oMenuItem = Nothing

End Function

Public Function ShowMainMenu()
Attribute ShowMainMenu.VB_Description = "This tells the control whether or not to show the menu."
Attribute ShowMainMenu.VB_MemberFlags = "40"
On Error Resume Next
'On Error GoTo ErrorHandler
    Dim oRect As RECT
    Dim lTop As Long
    Dim lCommand As Long
    Dim oMenu As clsMenuItem
    
    If m_nOMCount > 0 Then
        m_nAlign = UserControl.Extender.Align
        GetWindowRect UserControl.hWnd, oRect
        If m_nAlign = vbAlignLeft Then
            lCommand = TrackPopupMenu(m_Menu, TPM_HORPOSANIMATION Or TPM_VERTICAL, oRect.Right + 2, oRect.Top + 1, 0, UserControl.hWnd, 0)
        ElseIf m_nAlign = vbAlignRight Then
            lCommand = TrackPopupMenu(m_Menu, TPM_HORNEGANIMATION Or TPM_VERTICAL, oRect.Left - (oRect.Right - oRect.Left) * 1.155, oRect.Top + 1, 0, UserControl.hWnd, 0)
        ElseIf m_nAlign = vbAlignBottom Then
            For Each oMenu In MainMenu
                Select Case oMenu.Style
                    Case 0 ' normal menu item
                        lTop = lTop + 30
                    Case 1 ' separator bar
                        lTop = lTop + 17
                End Select
            Next
            lCommand = TrackPopupMenu(m_Menu, TPM_VERNEGANIMATION Or TPM_VERTICAL, oRect.Left + 2, oRect.Top - lTop, 0, UserControl.hWnd, 0)
        ElseIf m_nAlign = vbAlignTop Then
            lCommand = TrackPopupMenu(m_Menu, TPM_VERPOSANIMATION Or TPM_VERTICAL, oRect.Left + 2, oRect.Bottom + 2, 0, UserControl.hWnd, 0)
        End If

    End If
Done:
    Set oMenu = Nothing
    Exit Function
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_ShowMenu" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Function

Public Function AddTrayIcon(ByVal Key As String, ByVal ToolTip As String, IconIndex As Integer) As Boolean
Attribute AddTrayIcon.VB_Description = "This method is used to add an icon to the tray area."
    On Error GoTo err_routine
    Dim oTray As RECT
    Dim oIcon As clsTrayIcon
    Set oIcon = New clsTrayIcon
    
    oIcon.Key = Key
    oIcon.Icon = IconIndex
    oIcon.ToolTip = ToolTip
    If m_colTrayIcons Is Nothing Then
        Set m_colTrayIcons = New Collection
    End If
    m_colTrayIcons.Add oIcon, Key
    AddTrayIcon = True
    
    'redraw the tray
    TrayRect oTray
    InvalidateRect UserControl.hWnd, oTray, False
    UserControl.Cls
    UserControl_Paint
    
exit_routine:
    Set oIcon = Nothing
    Exit Function

err_routine:
    AddTrayIcon = False
    Err.Clear
    GoTo exit_routine
    
End Function

Public Function RemoveTrayIcon(ByVal Key As String) As Boolean
Attribute RemoveTrayIcon.VB_Description = "This method removes items from the tray area."
    On Error GoTo err_routine
    
    Dim oTray As RECT
    
    TrayRect oTray
    InvalidateRect UserControl.hWnd, oTray, False
    
    m_colTrayIcons.Remove Key
    RemoveTrayIcon = True
    
    'redraw the tray
    UserControl.Cls
    UserControl_Paint
    
exit_routine:
    Exit Function

err_routine:
    RemoveTrayIcon = False
    GoTo exit_routine

End Function

Public Function ChangeTrayIcon(ByVal Key As String, ByVal ToolTip As String, ByVal Icon As Long) As Boolean
    On Error GoTo err_routine
    
    m_colTrayIcons.Item(Key).ToolTip = ToolTip
    m_colTrayIcons.Item(Key).Icon = Icon
    
    ChangeTrayIcon = True
    
    UserControl_Paint
    
exit_routine:
    Exit Function

err_routine:
    ChangeTrayIcon = False
    GoTo exit_routine
End Function

Public Property Get MenuHighlightColor() As OLE_COLOR
On Error GoTo ErrorHandler
    MenuHighlightColor = m_MenuHighlightColor
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuHighlightColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let MenuHighlightColor(ByVal vNewValue As OLE_COLOR)
On Error GoTo ErrorHandler
    m_MenuHighlightColor = vNewValue
    gMenuHighlight = TranslateColor(m_MenuHighlightColor)
    PropertyChanged "MenuHighlightColor"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuHighlightColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Get MenuHighlightTextColor() As OLE_COLOR
On Error GoTo ErrorHandler
    MenuHighlightTextColor = m_MenuHighlightTextColor
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuHighlightTextColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let MenuHighlightTextColor(ByVal vNewValue As OLE_COLOR)
On Error GoTo ErrorHandler
    m_MenuHighlightTextColor = vNewValue
    gMenuHighlightText = TranslateColor(m_MenuHighlightTextColor)
    PropertyChanged "MenuHighlightTextColor"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuHighlightTextColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Get MenuBackColor() As OLE_COLOR
On Error GoTo ErrorHandler
    MenuBackColor = m_MenuBackColor
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuBackColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let MenuBackColor(ByVal vNewValue As OLE_COLOR)
On Error GoTo ErrorHandler
    m_MenuBackColor = vNewValue
    gMenuBackColor = TranslateColor(m_MenuBackColor)
    PropertyChanged "MenuBackColor"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuBackColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Get MenuForeColor() As OLE_COLOR
On Error GoTo ErrorHandler
    MenuForeColor = m_MenuForeColor
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuForeColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Property Let MenuForeColor(ByVal vNewValue As OLE_COLOR)
On Error GoTo ErrorHandler
    m_MenuForeColor = vNewValue
    gMenuForeColor = TranslateColor(m_MenuForeColor)
    PropertyChanged "MenuForeColor"
Done:
    Exit Property
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_MenuForeColor" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Property

Public Sub FlashMe(ByVal hWnd As Long, Optional Colour As Long, Optional FlashTimes As Integer)
Dim Icn As clsIcon, rcItem As RECT
Dim I As Integer

If Colour = 0 Then
    Colour = &H8080FF
End If

If FlashTimes = 0 Then
    FlashTimes = 5
End If

For I = 1 To m_colIcons.Count
    Set Icn = m_colIcons.Item(I)
    ItemRect I, rcItem
    If Icn.hWnd = hWnd Then
        Icn.FlashColour = Colour
        Icn.SetFlashCount = FlashTimes
        
        Icn.IsFlashing = True
        Exit For
    End If
    DoEvents
Next

End Sub

Public Sub Set_Unselected_Colour(ByVal hWnd As Long, ByVal Colour As Long)
If hWnd = 0 Then
    Exit Sub
End If
If Colour = 0 Then
    Colour = vbButtonFace
End If

Dim Icn As clsIcon, I As Integer
For I = 1 To m_colIcons.Count
    Set Icn = m_colIcons.Item(I)
    If Icn.hWnd = hWnd Then
        Icn.Unselected_BackColour = Colour
        Exit For
    End If
Next
Refresh

End Sub

Public Sub Set_Selected_Colour(ByVal hWnd As Long, ByVal Colour As Long)
If hWnd = 0 Then
    Exit Sub
End If
If Colour = 0 Then
    Colour = vbButtonFace
End If

Dim Icn As clsIcon, I As Integer
For I = 1 To m_colIcons.Count
    Set Icn = m_colIcons.Item(I)
    If Icn.hWnd = hWnd Then
        Icn.Selected_BackColour = Colour
        Exit For
    End If
Next
Refresh

End Sub

Public Sub Font_Color(ByVal hWnd As Long, Colour As Long)
Dim Icn As clsIcon, I As Integer

If hWnd = 0 Then
    Exit Sub
End If

If Colour = 0 Then
    Colour = vbBlack
End If

For I = 1 To m_colIcons.Count
    Set Icn = m_colIcons.Item(I)
    If Icn.hWnd = hWnd Then
        Icn.Change_Font_Colour = True
        Icn.FontColour = Colour
        Exit For
    End If
Next
Refresh

End Sub

Private Sub PaintOne(Buttonhwnd As Long)
On Error GoTo ErrorHandler
    ' painting whole area
    
    Dim I As Integer
    Dim rcItem As RECT
    Dim Icn As clsIcon
    Dim lEdgeParam As Long
    Dim hBrush As Long  ' receives handle to the blue hatched brush to use
    Dim r As RECT  ' rectangular area to fill
    Dim lRet As Long  ' return value
    Dim nDiff As Single
    Dim nIconTop As Single
    Dim rcIcon As RECT
    Dim lpDrawTextParams As DRAWTEXTPARAMS
    Dim nTextH As Single
    Dim oTest As POINTAPI
    Dim oTrayPoint As POINTAPI
    Dim vEdge As Variant
    Dim oTrayIcon As clsTrayIcon
    Dim bRet As Boolean
    
    If Not Ambient.UserMode() Then
        Exit Sub
    End If
            
    'Stop the Flash timer
    FlashTimer.Enabled = False
    
    I = 0
    For Each Icn In m_colIcons
        If Icn.hWnd = Buttonhwnd Then
            Exit For
        End If
        I = I + 1
    Next
    
    ' set the colors
    UserControl.BackColor = m_BackColor
    UserControl.ForeColor = m_ForeColor
    
    lpDrawTextParams.iLeftMargin = 1
    lpDrawTextParams.iRightMargin = 1
    lpDrawTextParams.iTabLength = 2
    lpDrawTextParams.cbSize = 20
        
    If ItemRect(I + 1, rcItem) Then
        ' three states of a push button
        If Icn Is m_refActive Then
            
            ' down, currently selected button
            If m_ShowActive Then
                ' they want to see the active button
                vEdge = EDGE_SUNKEN
                UserControl.FontBold = True
                hBrush = CreateSolidBrush(TranslateColor(Icn.Selected_BackColour))
                ' they dont want to see the active button
                ' differently, draw it with the same code
                ' that draws the inactvie buttons.
            
            Else
                If Icn.IsFlashing And Icn.FlashOn Then
                    hBrush = CreateSolidBrush(TranslateColor(Icn.FlashColour))
                Else
                    hBrush = CreateSolidBrush(TranslateColor(Icn.Unselected_BackColour))
                End If
                If m_Style = Default Then
                    ' if it is default then it looks like a raised button
                    vEdge = EDGE_RAISED
                End If
                UserControl.FontBold = False
            End If
            
        ElseIf I + 1 = m_nIndexBeingSelected And m_bInsetSelected Then
            ' being pushed at the moment
            vEdge = EDGE_SUNKEN
            UserControl.FontBold = False
            hBrush = CreateSolidBrush(TranslateColor(m_SelectingBackColor))
            
        Else
            ' up, inactive
            If m_Style = Default Then
                vEdge = EDGE_RAISED
            Else
                vEdge = 0
            End If
            UserControl.FontBold = False
            
            If Icn.IsFlashing Then
                hBrush = CreateSolidBrush(TranslateColor(Icn.FlashColour))
            Else
                hBrush = CreateSolidBrush(TranslateColor(Icn.Unselected_BackColour))
            End If
            
        End If
        
        ' draw the edge.
        If vEdge <> 0 Then
            DrawEdge UserControl.hdc, rcItem, vEdge, BF_RECT
        End If
        
        ' we selected the right hBrush up above, now use it to
        ' draw the back color.
        lRet = CopyRect(r, rcItem)
        r.Bottom = r.Bottom - 2
        r.Top = r.Top + 1
        r.Left = r.Left + 1
        r.Right = r.Right - 2
        lRet = FillRect(UserControl.hdc, r, hBrush)  ' fill the rectangle using the brush
        lRet = DeleteObject(hBrush) ' clean up
            
        ' fix the rect to fit inside the new border thats
        ' been drawn
        rcItem.Left = rcItem.Left + m_cxBorder + 1
        rcItem.Top = rcItem.Top + m_cyBorder
        rcItem.Right = rcItem.Right - m_cxBorder - 1
        rcItem.Bottom = rcItem.Bottom - m_cyBorder - 1
        
        ' used to calculate the position to draw the icon
        nDiff = rcItem.Bottom - rcItem.Top
        
        ' draw the icon
        ' calculate the position to draw the icon
        nIconTop = rcItem.Top + (nDiff - ICON_WIDTH) \ 2
        If Icn.IconPtr <> 0 Then
            DrawIconEx UserControl.hdc, rcItem.Left, nIconTop + 2, Icn.IconPtr, 16, 16, 0, 0, DI_NORMAL
        Else
            ' no icon was returned, so we cant draw anything.
        End If
        
        ' drawing a text with default font for a control
        If Icn.IconPtr <> 0 Then
            ' has an icon. so add space for it.
            rcItem.Left = rcItem.Left + ICON_WIDTH + 2
        Else
            ' no icon, so draw it over to the left.
            rcItem.Left = rcItem.Left + 2
        End If
        
        lpDrawTextParams.iLeftMargin = 1
        lpDrawTextParams.iRightMargin = 1
        lpDrawTextParams.iTabLength = 2
        lpDrawTextParams.cbSize = 20
        
        ' calculate all the dimensions for the rect to
        ' draw the text in.
        GetTextExtentPoint32 UserControl.hdc, Icn.Title, Len(Icn.Title), oTest
        nTextH = oTest.y
        If nTextH < nDiff Then
            nDiff = (nDiff - nTextH) \ 2
            rcItem.Bottom = rcItem.Bottom - nDiff
            rcItem.Top = rcItem.Top + nDiff
        End If
        
        'Set the Font Style.
        Set UserControl.Font = m_FontStyle
        
        If Icn.Change_Font_Colour Then
           UserControl.ForeColor = Icn.FontColour
           
        Else
           UserControl.ForeColor = Default_Font_Colour
           
        End If
            
        ' draw the text
        DrawTextEx UserControl.hdc, Icn.Title, Len(Icn.Title), rcItem, _
        DT_LEFT Or DT_VCENTER Or DT_WORD_ELLIPSIS, lpDrawTextParams
        
        
        'Set the default colour back to normal
        If Icn.Change_Font_Colour Then
           UserControl.ForeColor = Default_Font_Colour
             
        End If
    End If
    
    I = I + 1
    ' now if it is CoolBar style, we draw the separator
    ' line inbetween buttons
    If m_colIcons.Count > 1 And I < m_colIcons.Count Then
        If m_Style = CoolBar And m_CoolBarSeparator = True Then
            If m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
                ' draw 2 lines to give it the 3d look.
                UserControl.Line (r.Right + 3, r.Top)-(r.Right + 3, r.Bottom), vb3DShadow
                UserControl.Line (r.Right + 4, r.Top + 1)-(r.Right + 4, r.Bottom + 1), TranslateColor(m_SunkenBackColor)
            ElseIf m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
                UserControl.Line (r.Left, r.Bottom + 3)-(r.Right, r.Bottom + 3), vb3DShadow
                UserControl.Line (r.Left, r.Bottom + 4)-(r.Right, r.Bottom + 4), TranslateColor(m_SunkenBackColor)
            End If
        End If
    End If
    
    ' draw the tray
    If m_ShowTray And (Not m_colTrayIcons Is Nothing) Then
        TrayRect rcItem
        DrawEdge UserControl.hdc, rcItem, BDR_SUNKENOUTER, BF_RECT
        I = -1
        For Each oTrayIcon In m_colTrayIcons
            I = I + 1
            If TrayIconPoint(I, oTrayPoint) Then
                If hImageList <> 0 Then
                    ImageList_DrawEx hImageList, oTrayIcon.Icon, UserControl.hdc, oTrayPoint.x, oTrayPoint.y, 16, 16, CLR_NONE, CLR_DEFAULT, ILD_TRANSPARENT
                End If
            End If
        Next
        Set oTrayIcon = Nothing
    End If
    
Done:
    DeleteObject hBrush
    
    
    'Stop the Flash timer
    FlashTimer.Enabled = True


    Exit Sub
    
ErrorHandler:
    Dim lngErrNum As Long: Dim strErrDesc As String: lngErrNum = Err.Number: strErrDesc = Err.Description
    If InDesign = True Then: Stop: Else: Call HandleError("Class " & TypeName(Me) & "_UserControl_Paint" & vbCrLf & "Line# " & Erl & vbCrLf & "Err#" & Err.Number & vbCrLf & "Desc. " & Err.Description, App.Title, "")
    GoTo Done
End Sub
