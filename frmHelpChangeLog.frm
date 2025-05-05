VERSION 5.00
Begin VB.Form frmHelpChangeLog 
   Caption         =   "ChangeLog"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "frmHelpChangeLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   9360
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelpChangeLog.frx":08CA
      Top             =   0
      Width           =   9315
   End
End
Attribute VB_Name = "frmHelpChangeLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Private Sub Form_Load()
On Error Resume Next
Me.Top = 100
Me.Left = 100
Me.Height = 7000
Me.Width = 10000
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

