VERSION 5.00
Begin VB.Form frmHelpTextblocks 
   Caption         =   "Textblock Tutorial"
   ClientHeight    =   7815
   ClientLeft      =   4425
   ClientTop       =   2055
   ClientWidth     =   7815
   Icon            =   "frmHelpTextblocks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   7815
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelpTextblocks.frx":08CA
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "frmHelpTextblocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Private Sub Form_Load()
On Error Resume Next
Me.Top = ReadINI("Windows", "HelpTextblocksTop")
Me.Left = ReadINI("Windows", "HelpTextblocksLeft")
Me.Width = ReadINI("Windows", "HelpTextblocksWidth")
Me.Height = ReadINI("Windows", "HelpTextblocksHeight")
Me.Show
Me.SetFocus
End Sub

Private Sub Form_Resize()
 
    Dim lUseWidth As Long
    Dim lUseHeight As Long
 
    Const MINWIDTH As Long = 3000
    Const MINHEIGHT As Long = 3000
 
    'Copy the current width and height to our variables
    lUseWidth = Me.Width
    lUseHeight = Me.Height
 
    'Set a minimum limit on the lUseWidth and lUseHeight variables
    If lUseWidth < MINWIDTH Then lUseWidth = MINWIDTH
    If lUseHeight < MINHEIGHT Then lUseHeight = MINHEIGHT
 
    'Set the size of the textbox using the values in lUseWidth and lUseHeight
    With Text1
            .Move .Left, .Top, lUseWidth - 125, lUseHeight - TITLEBAR_OFFSET - 425
    End With
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
        If Me.WindowState = vbMinimized Then Exit Sub
        Call WriteINI("Windows", "HelpTextblocksTop", frmHelpTextblocks.Top)
        Call WriteINI("Windows", "HelpTextblocksLeft", frmHelpTextblocks.Left)
        Call WriteINI("Windows", "HelpTextblocksHeight", frmHelpTextblocks.Height)
        Call WriteINI("Windows", "HelpTextblocksWidth", frmHelpTextblocks.Width)
End Sub
