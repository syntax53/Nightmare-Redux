VERSION 5.00
Begin VB.Form frmHelpMessages 
   Caption         =   "Messages Help"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   Icon            =   "frmHelpMessages.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   7440
   Begin VB.TextBox Text1 
      Height          =   7935
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelpMessages.frx":08CA
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frmHelpMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Private Sub Form_Load()
On Error Resume Next
Me.Top = ReadINI("Windows", "HelpMessagesTop")
Me.Left = ReadINI("Windows", "HelpMessagesLeft")
Me.Width = ReadINI("Windows", "HelpMessagesWidth")
Me.Height = ReadINI("Windows", "HelpMessagesHeight")
Me.Show
Me.SetFocus
If ReadINI("Windows", "HelpMessagesMaxed") = "1" Then Me.WindowState = vbMaximized
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
        
        If Me.WindowState = vbMaximized Then
            Call WriteINI("Windows", "HelpMessagesMaxed", 1)
        Else
            Call WriteINI("Windows", "HelpMessagesMaxed", 0)
            Call WriteINI("Windows", "HelpMessagesTop", frmHelpMessages.Top)
            Call WriteINI("Windows", "HelpMessagesLeft", frmHelpMessages.Left)
            Call WriteINI("Windows", "HelpMessagesHeight", frmHelpMessages.Height)
            Call WriteINI("Windows", "HelpMessagesWidth", frmHelpMessages.Width)
        End If
End Sub

