VERSION 5.00
Begin VB.Form frmHelpRooms 
   Caption         =   "Rooms Help"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   Icon            =   "frmHelpRooms.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   5985
   Begin VB.TextBox Text1 
      Height          =   7575
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelpRooms.frx":08CA
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmHelpRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Private Sub Form_Load()
On Error Resume Next
Me.Top = ReadINI("Windows", "HelpRoomsTop")
Me.Left = ReadINI("Windows", "HelpRoomsLeft")
Me.Width = ReadINI("Windows", "HelpRoomsWidth")
Me.Height = ReadINI("Windows", "HelpRoomsHeight")
Me.Show
Me.SetFocus
If ReadINI("Windows", "HelpRoomsMaxed") = "1" Then Me.WindowState = vbMaximized
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
        Call WriteINI("Windows", "HelpRoomsMaxed", 1)
    Else
        Call WriteINI("Windows", "HelpRoomsMaxed", 0)
        Call WriteINI("Windows", "HelpRoomsTop", frmHelpRooms.Top)
        Call WriteINI("Windows", "HelpRoomsLeft", frmHelpRooms.Left)
        Call WriteINI("Windows", "HelpRoomsHeight", frmHelpRooms.Height)
        Call WriteINI("Windows", "HelpRoomsWidth", frmHelpRooms.Width)
    End If
End Sub

