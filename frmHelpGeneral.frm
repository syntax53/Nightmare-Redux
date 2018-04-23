VERSION 5.00
Begin VB.Form frmHelpGeneral 
   Caption         =   "General Info"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   Icon            =   "frmHelpGeneral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   7425
   Begin VB.TextBox Text1 
      Height          =   7935
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelpGeneral.frx":08CA
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frmHelpGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Private Sub Form_Load()
On Error Resume Next
Me.Top = ReadINI("Windows", "HelpGeneralTop")
Me.Left = ReadINI("Windows", "HelpGeneralLeft")
Me.Width = ReadINI("Windows", "HelpGeneralWidth")
Me.Height = ReadINI("Windows", "HelpGeneralHeight")
Me.Show
Me.SetFocus
If ReadINI("Windows", "HelpGeneralMaxed") = "1" Then Me.WindowState = vbMaximized
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
            Call WriteINI("Windows", "HelpGeneralMaxed", "1")
        Else
            Call WriteINI("Windows", "HelpGeneralMaxed", "0")
            Call WriteINI("Windows", "HelpGeneralTop", frmHelpGeneral.Top)
            Call WriteINI("Windows", "HelpGeneralLeft", frmHelpGeneral.Left)
            Call WriteINI("Windows", "HelpGeneralHeight", frmHelpGeneral.Height)
            Call WriteINI("Windows", "HelpGeneralWidth", frmHelpGeneral.Width)
        End If
End Sub


