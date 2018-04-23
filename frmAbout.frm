VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   5475
   ClientLeft      =   240
   ClientTop       =   630
   ClientWidth     =   5910
   ForeColor       =   &H00000000&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   5910
   Begin VB.CommandButton cmdChangeLog 
      Caption         =   "[ changelog ]"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3600
      TabIndex        =   12
      Top             =   5055
      Width           =   1035
   End
   Begin VB.TextBox txtScroll 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF8080&
      Height          =   2475
      HideSelection   =   0   'False
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "frmAbout.frx":08CA
      Top             =   1020
      Width           =   5355
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   5040
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Thanks and Props"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5655
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFC0C0&
         Height          =   435
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   2400
         Width           =   5235
      End
   End
   Begin VB.Frame frmLinks 
      BackColor       =   &H00000000&
      Caption         =   "WebLinks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1275
      Left            =   120
      TabIndex        =   4
      Top             =   3660
      Width           =   5655
      Begin VB.Label lblLinks 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "MajorMUD - Realm of Legends"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   3060
         MousePointer    =   2  'Cross
         TabIndex        =   8
         Top             =   900
         Width           =   2475
      End
      Begin VB.Label lblLinks 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "MUDiNFO.NET"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   4140
         MousePointer    =   2  'Cross
         TabIndex        =   7
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label lblLinks 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Nightmare Redux Source Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         MousePointer    =   2  'Cross
         TabIndex        =   6
         Top             =   900
         Width           =   2595
      End
      Begin VB.Label lblLinks 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Nightmare Redux Forums"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         MousePointer    =   2  'Cross
         TabIndex        =   5
         Top             =   300
         Width           =   2595
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFC0C0&
         Height          =   435
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   2400
         Width           =   5235
      End
   End
   Begin VB.Label lblSynEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "syntax53"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1860
      MousePointer    =   2  'Cross
      TabIndex        =   10
      Top             =   5085
      Width           =   915
   End
   Begin VB.Label lblb2yb 
      BackColor       =   &H00000000&
      Caption         =   "Brought to you by: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5100
      Width           =   1875
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Nightmare Redux v##"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

'Dim iFileNum As Integer
'Dim lLineCount As Long
'Dim lLineHeight As Long
Dim objTooltip As clsToolTip

Private Sub cmdChangeLog_Click()
Load frmHelpChangeLog
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim rc As RECT, i As Integer

Set objTooltip = New clsToolTip
With objTooltip
    .DelayTime = 25
    .VisibleTime = 10000
    .BkColor = &HC0FFFF
    .txtColor = &H0
    .Style = 1 'ttStyleBalloon
End With

lblVersion.Caption = sMenuCaption

'lLineCount = 51
'lLineHeight = TextHeight("TEST") 'Get the height of text in file
'txtScroll.Height = lLineHeight * lLineCount
'picScroll.Left = 0
'picScroll.Visible = True
'tmrScroll.Enabled = True

Me.Top = 200
Me.Left = 200



For i = 0 To 3 'number of lbl links
    rc.Left = lblLinks(i).Left \ Screen.TwipsPerPixelX
    rc.Top = lblLinks(i).Top \ Screen.TwipsPerPixelY
    rc.Bottom = (lblLinks(i).Top + lblLinks(i).Height) \ Screen.TwipsPerPixelX
    rc.Right = (lblLinks(i).Left + lblLinks(i).Width) \ Screen.TwipsPerPixelY
    Select Case i
        Case 0:
            objTooltip.SetToolTipItem frmLinks.hwnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, "http://www.mudinfo.net/viewforum.php?f=44", False
        Case 1:
            objTooltip.SetToolTipItem frmLinks.hwnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, "https://github.com/syntax53/Nightmare-Redux", False
        Case 2:
            objTooltip.SetToolTipItem frmLinks.hwnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, "http://www.mudinfo.net/", False
        Case 3:
            objTooltip.SetToolTipItem frmLinks.hwnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, "http://www.majormud.com/", False
    End Select
Next

rc.Left = lblSynEmail.Left \ Screen.TwipsPerPixelX
rc.Top = lblSynEmail.Top \ Screen.TwipsPerPixelY
rc.Bottom = (lblSynEmail.Top + lblSynEmail.Height) \ Screen.TwipsPerPixelX
rc.Right = (lblSynEmail.Left + lblSynEmail.Width) \ Screen.TwipsPerPixelY
objTooltip.SetToolTipItem Me.hwnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, "syntax53@mudinfo.net", False

Me.Show
Me.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set objTooltip = Nothing
End Sub

Private Sub lblLinks_Click(Index As Integer)

Select Case Index
    Case 0: Call ShellExecute(0&, "open", "http://www.mudinfo.net/viewforum.php?f=44", vbNullString, vbNullString, vbNormalFocus)
    Case 1: Call ShellExecute(0&, "open", "https://github.com/syntax53/Nightmare-Redux", vbNullString, vbNullString, vbNormalFocus)
    Case 2: Call ShellExecute(0&, "open", "http://www.mudinfo.net/", vbNullString, vbNullString, vbNormalFocus)
    Case 3: Call ShellExecute(0&, "open", "http://www.majormud.com/", vbNullString, vbNullString, vbNormalFocus)
End Select

End Sub

Private Sub lblSynEmail_Click()
    Call ShellExecute(0&, "open", "mailto:syntax53@mudinfo.net &subject=Nightmare Redux", vbNullString, vbNullString, vbNormalFocus)
End Sub

'Private Sub tmrScroll_Timer()
    'scroll txtScroll
'    If txtScroll.Top + txtScroll.Height < picScroll.Top Then 'picScroll.Top
'        txtScroll.Top = picScroll.Height
'    Else
'        txtScroll.Top = txtScroll.Top - 25
'    End If
'End Sub
Private Sub txtScroll_GotFocus()
    cmdOK.SetFocus
    'Don't let the text box get focus, althought the text
    'box is locked it looks bad to see a cursor in the
    'text box as it scrolls up
End Sub
