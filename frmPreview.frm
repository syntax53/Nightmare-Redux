VERSION 5.00
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Preview Window"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Timer CursorTimer 
      Interval        =   500
      Left            =   960
      Top             =   1200
   End
   Begin VB.PictureBox inputCheck 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   2640
      Width           =   255
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit
'Dim control_on As Boolean

Private Sub CursorTimer_Timer()
term_DriveCursor
End Sub

Private Sub Form_Load()
On Error Resume Next
    PreviewLoaded = True
    term_reset_matrix
    'modANSIStuff.strAnsi = "[1;32m"
    modANSIStuff.MyBackColor = 40
    modANSIStuff.MyForeColor = 32
    inputCheck.Top = Screen.Height
    inputCheck.Height = 0
    inputCheck.Width = 0
    'Get the pixel metrics of the current font
    Me.FontUnderline = False
    Me.FontItalic = False
    Me.FontBold = False
    
    Me.ScaleMode = 3
    modANSIStuff.charHeight = Me.TextHeight("M")
    modANSIStuff.charWidth = Me.TextWidth("M")

    'Set up the vt100 screen
    Me.ScaleMode = 1
    Me.Height = (Me.Height - Me.ScaleHeight) + modANSIStuff.LinesPerPage * Me.TextHeight("M")
    Me.Width = (Me.Width - Me.ScaleWidth) + modANSIStuff.CharsPerLine * Me.TextWidth("M")


    'Set the user scale of the display
    Me.ScaleMode = 0
    Me.ScaleWidth = modANSIStuff.LinesPerPage
    Me.ScaleWidth = modANSIStuff.CharsPerLine
    Me.Scale (0, 0)-(modANSIStuff.LastChar, modANSIStuff.LastLine)

    frmPreview.Top = ReadINI("Windows", "PreviewTop")
    frmPreview.Left = ReadINI("Windows", "PreviewLeft")
    frmPreview.Width = ReadINI("Windows", "PreviewWidth")
    frmPreview.Height = ReadINI("Windows", "PreviewHeight")
    
    term_init
    
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'frmPreview.CursorTimer.Enabled = False
'term_Carethide
'frmPreview.Refresh
'
'Dim percentdown
'Dim percentright
'
'percentdown = y / frmPreview.ScaleHeight
'percentright = x / frmPreview.ScaleWidth
'
'CurX = Int(percentright * modANSIStuff.CharsPerLine)
'CurY = Int(percentdown * modANSIStuff.LinesPerPage)
'term_Caretshow
'frmPreview.CursorTimer.Enabled = True
'frmPreview.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState = vbMinimized Then GoTo SkipWL:
    Call WriteINI("Windows", "PreviewTop", frmPreview.Top)
    Call WriteINI("Windows", "PreviewLeft", frmPreview.Left)
    Call WriteINI("Windows", "PreviewWidth", frmPreview.Width)
    Call WriteINI("Windows", "PreviewHeight", frmPreview.Height)
SkipWL:
    PreviewLoaded = False
    Unload Me
End Sub

Public Sub inputCheck_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
'  frmPreview.CursorTimer.Enabled = False
'  term_Carethide
'  frmPreview.Refresh
'
'  Dim CH As String
'  Dim addme
'  addme = False
'  CH = Chr$(0)
'
'  'Translate keycodes to VT100 escape sequences
'  'DoEvents
'  Select Case KeyCode
'    'Case vbKeyControl
'    '  control_on = True
'    'Case vbKeyEnd
'    '  CH = Chr$(27) + "[K"
'    'Case vbKeyHome
'    '  CH = Chr$(27) + "[H"
'    Case vbKeyLeft
'      If modANSIStuff.CurX > 0 Then
'        modANSIStuff.CurX = modANSIStuff.CurX - 1
'        'update_colors
'      End If
'    Case vbKeyUp
'      If modANSIStuff.CurY > 0 Then
'        modANSIStuff.CurY = modANSIStuff.CurY - 1
'        'update_colors
'      End If
'    Case vbKeyRight
'      If modANSIStuff.CurX < modANSIStuff.LastChar Then
'        modANSIStuff.CurX = modANSIStuff.CurX + 1
'        'update_colors
'      End If
'    Case vbKeyDown
'      If modANSIStuff.CurY < modANSIStuff.LastLine Then
'        modANSIStuff.CurY = modANSIStuff.CurY + 1
'        'update_colors
'      End If
'
'    Case vbKeyDelete
'
'      modANSIStuff.doNotAdvance = True
'      term_write Asc(" ")
'      modANSIStuff.doNotAdvance = False
'
'    'Case vbKeyF1
'    '  CH = Chr$(27) + "OP"
'    'Case vbKeyF2
'    '  CH = Chr$(27) + "OQ"
'    'Case vbKeyF3
'    '  CH = Chr$(27) + "OR"
'    'Case vbKeyF4
'    '  CH = Chr$(27) + "OS"
'    Case Else
'      If control_on And KeyCode > 63 Then
'        CH = Chr$(KeyCode - 64)
'        addme = True
'      End If
'
'  End Select
'
'  If CH > Chr$(0) Then
'    For i = 1 To Len(CH)
'        term_process_char (Asc(Mid(CH, i, 1)))
'        'modANSIStuff.strAnsi = modANSIStuff.strAnsi & Mid(CH, i, 1)
'    Next i
'  End If
'term_Caretshow
'frmPreview.CursorTimer.Enabled = True
'frmPreview.Refresh
End Sub

Private Sub inputCheck_KeyPress(KeyAscii As Integer)
KeyAscii = 0

'frmPreview.CursorTimer.Enabled = False
'term_Carethide
'frmPreview.Refresh
'Dim CH As String
'
'        CH = Chr$(KeyAscii)
'        If control_on Then
'          If KeyAscii > 63 Then
'            CH = Chr$(KeyAscii - 64)
'          Else
'            CH = Chr$(0)
'          End If
'        End If
'
'        If CH > Chr$(0) Then
'              'If CH <> Chr(13) Then modANSIStuff.lastKeyAlpha = True
'              For i = 1 To Len(CH)
'                term_process_char (Asc(Mid(CH, i, 1)))
'                'If Asc(CH) <> 8 Then modANSIStuff.strAnsi = modANSIStuff.strAnsi & Mid(CH, i, 1)
'              Next i
'        End If
'
'term_Caretshow
'frmPreview.CursorTimer.Enabled = True
'frmPreview.Refresh


End Sub

Private Sub inputCheck_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
'    Select Case KeyCode
'      Case vbKeyControl
'        control_on = False
'    End Select


End Sub

'Private Sub update_colors()
        'modANSIStuff.MyBackColor = Int(Mid(previewMatrix(modANSIStuff.CurX, modANSIStuff.CurY), 3, 2))
        'modANSIStuff.MyForeColor = Int(Mid(previewMatrix(modANSIStuff.CurX, modANSIStuff.CurY), 5, 2))
        'modANSIStuff.isBold = Int(Mid(previewMatrix(modANSIStuff.CurX, modANSIStuff.CurY), 2, 1))
'End Sub
