VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMapFavorites 
   Caption         =   " Map Favorites"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   4215
   Begin MSComctlLib.ListView lvResults 
      Height          =   3075
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5424
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmMapFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Sub Form_Load()
frmResults.lvResults.ColumnHeaders.Clear
frmResults.lvResults.ColumnHeaders.add 1, "Location", "Location", 3500
Me.Top = frmMain.Top
Me.Left = frmMain.Left
Me.Width = 4335
Me.Height = 2500
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
If Me.Width < 2000 Then Exit Sub
If Me.Height < 2000 Then Exit Sub

lvResults.Width = Me.Width - 130
lvResults.Height = Me.Height - TITLEBAR_OFFSET - 400
lvResults.ColumnHeaders(1).Width = lvResults.Width - 500
End Sub

Private Sub lvResults_DblClick()
If lvResults.SelectedItem Is Nothing Then Exit Sub

Call frmMain.GotoLocation(lvResults.SelectedItem)

End Sub
