Attribute VB_Name = "modANSIStuff"
Option Base 0
Option Explicit

Public doNotAdvance
Public lastKeyAlpha As Boolean

Public AlreadyUpdating As Boolean

'Private Declare Function ScrollWindow Lib "user32" (ByVal hWnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As RECT, lpClipRect As RECT) As Long
Private Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hRgnUpdate As Long, lprcUpdate As RECT) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function GetBkMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal newColor As Long) As Long
Private Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long


Public strAnsi As String
Public MyForeColor As Integer
Public MyBackColor As Integer
Public isBold As Boolean


'Private Const PATCOPY = &HF00021         ' (DWORD) dest = pattern
'Private Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
'Private Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Private Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
'Private Const BLACKNESS = &H42&          ' (DWORD) dest = BLACK
'Private Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE
'Private Const TRANSPARENT = 1
Private Const OPAQUE = 2
'Private Const GO_IAC1 = 6

Public Const LinesPerPage = 25
Public Const CharsPerLine = 80
Public Const TabsPerPage = 20
Public Const LastLine = LinesPerPage - 1
Public Const LastChar = CharsPerLine - 1
'Private Const LastTab = 19

'Private ScrImage(LinesPerPage)    As String * CharsPerLine
'Private ScrAttr(LinesPerPage)     As String * CharsPerLine

'Private Norm_Attr                 As String * CharsPerLine
Private Blank_Line                As String * CharsPerLine
Private Block_Line                As String * CharsPerLine

Private TermTextColor             As Long
Private TermBkColor               As Long

Private tabno                     As Integer
Private tab_table(TabsPerPage)    As Integer
'Private curattr                   As String

Private lprcScroll                As RECT
Private lprcClip                  As RECT
Private hRgnUpdate                As Integer
Private lprcUpdate                As RECT

'
'   Current Buffered Text waiting for output on screen
'

'Private OutStr          As String
'Private outlen          As Integer

'
'   Flag to indicate that we're ready to run
'
'Private FlagInit As Integer

Public CurX As Integer
Public CurY As Integer
Private SavecurX As Integer
Private SavecurY As Integer

Public inEscape As Boolean    ' Processing an escape seq?
Private EscString As String     ' String so far

Public charHeight As Single
Public charWidth As Single

Private oldForeColor As Long
Private oldBackColor As Long

Public previewMatrix(LastChar, LastLine) As String

Public CurState As Boolean
Private ret As Long
'Private tAttrBold As String * CharsPerLine

Public Function term_process_char(CH As Byte)

       
        
        If (inEscape) Then
           term_escapeProcess (CH)
        Else
       
        Select Case CH
        
        Case 0

        Case 7
            Beep
        Case 8
            'If CurX > 0 Then                    '   if not at line begin
            '    CurX = CurX - 1                 '   adjust back 1 spc

            'End If

        Case 9
            Dim tY As Integer
            For tY = 0 To 19
              If CurY < tab_table(tY) Then
                Exit For
              End If
            Next tY
            CurY = tab_table(tY)
            

        Case 10, 11, 12, 13
            If frmPreview.CursorTimer.Enabled Then
                frmPreview.CursorTimer.Enabled = False
                term_Carethide
                frmPreview.refresh
                
                CurY = CurY + 1
                CurX = 0
                
                term_Caretshow
                frmPreview.CursorTimer.Enabled = True
                frmPreview.refresh
            Else
                CurY = CurY + 1
                CurX = 0
            End If

        
        Case 27
            
            inEscape = True
            EscString = ""

        Case Else
                
                term_write CH
        
        End Select
        
    End If

End Function

Private Sub term_escapeProcess(CH As Byte)

Dim C           As String
Dim yDiff       As Integer
Dim xDiff       As Integer

    C = Chr$(CH)
    If EscString = "" Then
      'No start character yet
      Select Case C
        Case "["
        
        Case "("
        
        Case ")"
        
        Case "#"
        
        Case Chr$(8)             ' embedded backspace
          CurX = CurX - 1
          term_validatecurX
          inEscape = False
        
        Case "7"                 ' save cursor
          'Save cursor position
          SavecurX = CurX
          SavecurY = CurY
          inEscape = False
        
        Case "8"                 ' restore cursor
          'restore cursor position
          CurX = SavecurX
          CurY = SavecurY
          inEscape = False
        
        Case "c"                 ' look at VSIreset()
        
        Case "D"                 ' cursor down
          CurY = CurY + 1
          term_validatecurY
          inEscape = False
        
        Case "E"                 ' next line
          CurY = CurY + 1
          CurX = 0
          term_validatecurY
          term_validatecurX
          inEscape = False
        
        Case "H"                 ' set tab
          tab_table(tabno) = CurY
          tabno = tabno + 1
          inEscape = False
        
        Case "I"                 ' look at bp_ESC_I()
          inEscape = False
        
        Case "M"                 ' cursor up
          CurY = CurY - 1
          term_validatecurY
                
        Case "Z"                 ' send ident
          inEscape = False
        
        Case Else
            inEscape = False
            Exit Sub
      End Select
    End If

    EscString = EscString & C
    If IsCharAlpha(CH) = 0 Then
        ' Not a character ...
        If Len(EscString) > 15 Then
            inEscape = False
        End If
        Exit Sub
    End If


    Select Case C

        Case "A"

            ' A ==> move cursor up
            
            EscString = Mid$(EscString, 2)

            yDiff = Val(term_escapeParseArg(EscString))
            If yDiff = 0 Then
                yDiff = 1
            End If

            CurY = CurY - yDiff
            term_validatecurY
        
        Case "B"

            ' B ==> move cursor down
            
            EscString = Mid$(EscString, 2)

            yDiff = Val(term_escapeParseArg(EscString))
            If yDiff = 0 Then
                yDiff = 1
            End If

            CurY = CurY + yDiff
            term_validatecurY

        Case "C"
            ' C ==> move cursor right

            EscString = Mid$(EscString, 2)

            xDiff = Val(term_escapeParseArg(EscString))
            If xDiff = 0 Then
                xDiff = 1
            End If

            CurX = CurX + xDiff
            term_validatecurX
        
        Case "D"
            ' D ==> move cursor left

            EscString = Mid$(EscString, 2)

            xDiff = Val(term_escapeParseArg(EscString))
            If xDiff = 0 Then
                xDiff = 1
            End If
            CurX = CurX - xDiff
            term_validatecurX
        
        Case "H"

            'Goto cursor position indicated by escape sequence

            EscString = Mid$(EscString, 2)

            CurY = Val(term_escapeParseArg(EscString)) - 1
            term_validatecurY

            CurX = Val(EscString) - 1
            term_validatecurX

        Case "J"

            'Erase display

            Select Case Val(Mid$(EscString, 2))

                Case 0
                    If CurX = 0 And CurY = 0 Then
                        Call term_eraseSCREEN
                    Else
                        Call term_eraseEOS
                    End If

                Case 1
                    Call term_eraseBOS

                Case 2
                    Call term_eraseSCREEN

            End Select

        Case "K"

            'Erase line
            Select Case Val(Mid$(EscString, 2))
                Case 0
                    'erase to end of line
                    Call term_eraseEOL
                Case 1
                    'erase to end of line
                    Call term_eraseBOL
                Case 2
                    Call term_eraseLINE
            End Select

        Case "f"

            'Goto cursor position indicated by escape sequence

            EscString = Mid$(EscString, 2)

            CurY = Val(term_escapeParseArg(EscString)) - 1
            term_validatecurY

            CurX = Val(EscString) - 1
            term_validatecurX
        
        Case "g"
            ' clear tabs
            
            Dim tY As Integer
            For tY = 0 To 19
              tab_table(tY) = 0
            Next tY
        
        Case "h"

            'restore cursor position
            CurX = SavecurX
            CurY = SavecurY

        Case "i"
            ' print though mode
        
        Case "l"
            'Save cursor position
            SavecurX = CurX
            SavecurY = CurY

        Case "m"

            'Change text attributes, screen colors
            
            EscString = Mid$(EscString, 2)
            Do
                Call term_setattr(Chr$(Val(term_escapeParseArg(EscString))))
            Loop While EscString <> ""

        Case "r"
            
            'Set scrollable region
            EscString = Mid$(EscString, 2)

            lprcScroll.Top = (Val(term_escapeParseArg(EscString)) - 1) * charHeight
            lprcClip = lprcScroll
        
        Case "s"
            'Save cursor position
            SavecurX = CurX
            SavecurY = CurY

        Case "u"

            'restore cursor position
            CurX = SavecurX
            CurY = SavecurY


        Case Else

          'If frmPreview.Tracevt100 Then Debug.Print EscString

    End Select

    inEscape = False
    EscString = ""

End Sub

'Private Sub term_scroll_up()
'
'    Dim i As Integer
'    Dim S As Integer
'
'    oldForeColor = GetTextColor(frmPreview.hdc)
'    oldBackColor = GetBkColor(frmPreview.hdc)
'
'
'    ret = SetBkColor(frmPreview.hdc, RGB(0, 0, 0))
'    ret = SetTextColor(frmPreview.hdc, RGB(0, 0, 0))
'
'
'    If frmPreview.WindowState <> 1 Then
'         ret = ScrollDC(frmPreview.hdc, 0, -charHeight, lprcScroll, lprcClip, hRgnUpdate, lprcUpdate)
'         ret = TextOut(frmPreview.hdc, 0, CurY * charHeight, Block_Line, CharsPerLine)
'
'    End If
'    ret = SetBkColor(frmPreview.hdc, oldBackColor)
'    ret = SetTextColor(frmPreview.hdc, oldForeColor)
'
'
'    frmPreview.refresh
'
'    'Update the redisplay buffer (only update the scrollable region)
'    'Might consider making this a circular array so only one line
'    'needs to be written per scroll, rather than relinking the array
'    'S = (lprcScroll.Top \ charheight + 1)
'
'End Sub

Public Sub term_write(CH As Byte)


    Dim intIsBold As Integer
    
    If isBold = True Then
        intIsBold = 1
    Else
        intIsBold = 0
    End If

    If CurY < 25 Then previewMatrix(CurX, CurY) = Chr$(CH) & Trim(str(intIsBold)) & Trim(str(MyBackColor)) & Trim(str(MyForeColor))
    
    If frmPreview.WindowState <> 1 Then
        oldForeColor = GetTextColor(frmPreview.hdc)
        ret = SetTextColor(frmPreview.hdc, GetBkColor(frmPreview.hdc))
        
        ret = TextOut(frmPreview.hdc, CurX * charWidth, CurY * charHeight, Chr$(219), 1)
        
        term_setattr Chr(MyBackColor)
        term_setattr Chr(MyForeColor)
        
        'Ret = SetTextColor(frmPreview.hdc, oldForeColor)
        
        ret = TextOut(frmPreview.hdc, CurX * charWidth, CurY * charHeight, Chr$(CH), 1)
    End If

    If Not doNotAdvance Then
    If Not (CurX = LastChar) Then
        CurX = CurX + 1
    Else
        term_process_char (10)
        'term_process_char (13)
        'strAnsi = strAnsi & vbCrLf
    End If
    End If
    

End Sub


Public Function term_init()
    doNotAdvance = False
    lastKeyAlpha = False
    CurState = False
    ret = SetBkMode(frmPreview.hdc, OPAQUE)
    frmPreview.ForeColor = RGB(0, 128, 0)
    frmPreview.BackColor = QBColor(0)
    ret = SetBkColor(frmPreview.hdc, frmPreview.BackColor)
    ret = SetTextColor(frmPreview.hdc, frmPreview.ForeColor)

    TermTextColor = GetTextColor(frmPreview.hdc)
    TermBkColor = GetBkColor(frmPreview.hdc)


    'Initialize repaint buffer
    'Norm_Attr = String$(CharsPerLine, "0")
    Blank_Line = Space$(CharsPerLine)
    Block_Line = String$(CharsPerLine, Chr$(219))
    'term_eraseBUFFER
    'tAttrBold = Norm_Attr

End Function
Private Sub term_validatecurX()
   If (CurX < 0) Then
        CurX = 0
   ElseIf CurX > LastChar Then
        CurX = LastChar
   End If
End Sub

Private Sub term_validatecurY()
   If (CurY < 0) Then
        CurY = 0
   ElseIf CurY > LastLine Then
        CurY = LastLine
   End If
End Sub

Private Function term_escapeParseArg(S As String) As String
'
'   PopArg takes the next argument (digits up to a ;) and
'   returns it.  It also removes the arg and the ; from
'   the "s"

    Dim i As Integer

    i = InStr(S, ";")
    If i = 0 Then
        term_escapeParseArg = S
        S = ""
    Else
        term_escapeParseArg = Left$(S, i - 1)
        S = Mid$(S, i + 1)
    End If

End Function

Public Sub term_eraseSCREEN()

    'Assume that they want to repaint using the latest background color
    
    
    'term_eraseBUFFER
    frmPreview.Cls
    CurX = 0
    CurY = 0

End Sub

Private Sub term_eraseEOL()
'
'   Erase to End of Line
'
    If frmPreview.WindowState <> 1 Then
        ret = TextOut(frmPreview.hdc, CurX * charWidth, CurY * charHeight, Space$(CharsPerLine - CurX), CharsPerLine - CurX)
    End If

    'Update screen buffer
'    Mid$(ScrImage(CurY + 1), CurX + 1, CharsPerLine - CurX) = Space$(CharsPerLine - CurX)
'    Mid$(ScrAttr(CurY + 1), CurX + 1, CharsPerLine - CurX) = String$(CharsPerLine - CurX, "0")
'    Mid$(ScrAttrBold(CurY + 1), CurX + 1, CharsPerLine - CurX) = String$(CharsPerLine - CurX, "0")

End Sub

Private Sub term_eraseEOS()
'
'   Erase to end of screen
'
    'Dim y As Integer

    Call term_eraseEOL
    If (CurY <> LastLine) Then

        If frmPreview.WindowState <> 1 Then
            ret = TextOut(frmPreview.hdc, 0, (CurY + 1) * charHeight, Space$((LastLine - CurY) * CharsPerLine), (LastLine - CurY) * CharsPerLine)
        End If

'        For Y = CurY + 2 To LinesPerPage
'            ScrImage(Y) = Blank_Line
'            ScrAttr(Y) = Norm_Attr
'            ScrAttrBold(Y) = Norm_Attr
'        Next Y

     End If
End Sub

Private Sub term_eraseLINE()

'   Erase Line

    If frmPreview.WindowState <> 1 Then
        ret = TextOut(frmPreview.hdc, 0, CurY * charHeight, Blank_Line, CharsPerLine)
    End If

'    ScrImage(CurY + 1) = Blank_Line
'    ScrAttr(CurY + 1) = Norm_Attr
'    ScrAttrBold(CurY + 1) = Norm_Attr

End Sub

Private Sub term_eraseBOL()
'------------------------------------------------------------------------
'   term_eraseBOL
'   erase from beginning of current line
'------------------------------------------------------------------------
    Dim ret As Integer

    If frmPreview.WindowState <> 1 Then
       ' Ret = PatBlt(frmPreview.hdc, 0, CurY * charheight, curX * charWidth, charheight, BLACKNESS)
        ret = TextOut(frmPreview.hdc, 0, CurY * charHeight, Blank_Line, CharsPerLine)
        
    End If

'    Mid$(ScrImage(CurY + 1), 1, CurX + 1) = Space$(CurX + 1)
'    Mid$(ScrAttr(CurY + 1), 1, CurX + 1) = String$(CurX + 1, "0")
'    Mid$(ScrAttrBold(CurY + 1), 1, CurX + 1) = String$(CurX + 1, "0")
End Sub

Private Sub term_eraseBOS()
'------------------------------------------------------------------------
'   term_eraseBOS
'   erase all lines from beginning of screen to and including current
'------------------------------------------------------------------------
    'Dim y As Integer

    'Erase the current line first
    Call term_eraseBOL

    'Erase everything up to current line
    If (CurY > 0) Then
        If frmPreview.WindowState <> 1 Then
            ret = TextOut(frmPreview.hdc, 0, 0, Space$(CharsPerLine * CurY + CurX), CharsPerLine * CurY + CurX)
            
        End If

        ' reset screen buffer contents
'        For Y = 1 To CurY
'           ScrImage(Y) = Blank_Line
'           ScrAttr(Y) = Norm_Attr
'           ScrAttrBold(Y) = Norm_Attr
'        Next Y
    End If
End Sub

Public Sub term_setattr(CH As String)
    Select Case Asc(CH)

            Case 0 'normal
                'Ret = SetTextColor(frmPreview.hdc, RGB(255, 255, 255))
                'Ret = SetBkColor(frmPreview.hdc, TermBkColor)
                isBold = False
                                        
            Case 1 '"Bold" text - remember for later
                isBold = True
                
            Case 7  '   Reverse Video
                ret = SetTextColor(frmPreview.hdc, TermBkColor)
                ret = SetBkColor(frmPreview.hdc, TermTextColor)
        
            Case 8  '   Cancel (Invisible)
                'Attr_BitMap = Attr_BitMap And ATTR_INVISIBLE
                ret = SetTextColor(frmPreview.hdc, TermBkColor)
                ret = SetBkColor(frmPreview.hdc, TermBkColor)

'-----------------------------------------------------------------------------

            Case 30  '  Black Foreground
                MyForeColor = 30
                If isBold = True Then
                    ret = SetTextColor(frmPreview.hdc, RGB(128, 128, 128))
                Else
                    ret = SetTextColor(frmPreview.hdc, RGB(0, 0, 0))
                End If
                
            Case 31  '  Red Foreground
                MyForeColor = 31
                If isBold = True Then
                    ret = SetTextColor(frmPreview.hdc, RGB(255, 0, 0))
                Else
                    ret = SetTextColor(frmPreview.hdc, RGB(128, 0, 0))
                End If
                
            Case 32  '  Green Foreground
                MyForeColor = 32
                If isBold = True Then
                    ret = SetTextColor(frmPreview.hdc, RGB(0, 255, 0))
                Else
                    ret = SetTextColor(frmPreview.hdc, RGB(0, 128, 0))
                End If
                
            Case 33  '  Orange Foreground
                MyForeColor = 33
                If isBold = True Then
                    ret = SetTextColor(frmPreview.hdc, RGB(255, 255, 0))
                Else
                    ret = SetTextColor(frmPreview.hdc, RGB(128, 128, 0))
                End If
                
            Case 34  '  Blue Foreground
                MyForeColor = 34
                If isBold = True Then
                    ret = SetTextColor(frmPreview.hdc, RGB(0, 0, 255))
                Else
                    ret = SetTextColor(frmPreview.hdc, RGB(0, 0, 128))
                End If
                
            Case 35  '  Pink Foreground
                MyForeColor = 35
                If isBold = True Then
                    ret = SetTextColor(frmPreview.hdc, RGB(255, 0, 255))
                Else
                    ret = SetTextColor(frmPreview.hdc, RGB(128, 0, 128))
                End If
                
            Case 36  '  Cyan Foreground
                MyForeColor = 36
                If isBold = True Then
                    ret = SetTextColor(frmPreview.hdc, RGB(0, 255, 255))
                Else
                    ret = SetTextColor(frmPreview.hdc, RGB(0, 128, 128))
                End If
                
            Case 37  '  Gray Foreground
                MyForeColor = 37
                If isBold = True Then
                    ret = SetTextColor(frmPreview.hdc, RGB(255, 255, 255))
                Else
                    ret = SetTextColor(frmPreview.hdc, RGB(192, 192, 192))
                End If
                
'------------------------------------------------------------
'BACKGROUND COLORS
'------------------------------------------------------------
 
            Case 40 '   Black Background
                MyBackColor = 40
                ret = SetBkColor(frmPreview.hdc, QBColor(0))

            Case 41 '   Red Background
                MyBackColor = 41
                ret = SetBkColor(frmPreview.hdc, QBColor(4))

            Case 42 '   Green Background
                MyBackColor = 42
                ret = SetBkColor(frmPreview.hdc, QBColor(2))

            Case 43 '   Orange Background
                MyBackColor = 43
                ret = SetBkColor(frmPreview.hdc, RGB(128, 128, 0))

            Case 44 '   Blue Background
                MyBackColor = 44
                ret = SetBkColor(frmPreview.hdc, QBColor(1))

            Case 45 '   Magenta Background
                MyBackColor = 45
                ret = SetBkColor(frmPreview.hdc, QBColor(5))

            Case 46 '   Cyan Background
                MyBackColor = 46
                ret = SetBkColor(frmPreview.hdc, QBColor(3))

            Case 47 '   White Background
                MyBackColor = 47
                ret = SetBkColor(frmPreview.hdc, RGB(168, 168, 168))

    End Select
    
    
End Sub

Public Sub term_reset_matrix()
    Dim i, j
    For i = 0 To 79
        For j = 0 To 24
            previewMatrix(i, j) = " 04032"
        Next j
    Next i
End Sub

Public Sub term_DriveCursor()
frmPreview.refresh
    If CurState = False Then
        Call term_Caretshow
    Else
        Call term_Carethide
    End If
frmPreview.refresh
End Sub

Public Sub term_Carethide()

        If CurState = True Then

        If frmPreview.WindowState <> 1 Then
            ret = PatBlt(frmPreview.hdc, CurX * charWidth, CurY * charHeight, charWidth, charHeight, DSTINVERT)
            'Debug.Print "hid"
        End If
        CurState = False
        
        End If
    
End Sub

Public Sub term_Caretshow()

    '------------------------------------------------------------------------
    '   term_CaretShow
    '
    '   display the inverted block cursor on the screen.
    '   currently uses PatBlt.
    '------------------------------------------------------------------------
    Dim ret As Integer
    
   
    If CurState = False Then
    
    If frmPreview.WindowState <> 1 Then
       ret = PatBlt(frmPreview.hdc, CurX * charWidth, CurY * charHeight, charWidth, charHeight, DSTINVERT)
       'Debug.Print "shown"
    End If

    CurState = True
    
    End If

End Sub

'Public Sub term_CaretControl(TurnOff As Boolean)
'Static SaveState As Boolean

'    If TurnOff = True Then
'        SaveState = CurState
'        frmPreview.CursorTimer.Enabled = False
'        term_Carethide
'    Else
'        If SaveState = True Then
'            term_Caretshow
'            frmPreview.CursorTimer.Enabled = True
'        End If
'    End If
    
'End Sub

