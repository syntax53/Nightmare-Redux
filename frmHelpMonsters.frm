VERSION 5.00
Begin VB.Form frmHelpMonsters 
   Caption         =   "Monsters Help"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   Icon            =   "frmHelpMonsters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   7440
   Begin VB.CommandButton cmdGI2 
      Caption         =   "Group/Index p2"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   60
      Width           =   1635
   End
   Begin VB.CommandButton cmdBL 
      Caption         =   "Boss/NPC List"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   60
      Width           =   1695
   End
   Begin VB.CommandButton cmdGI1 
      Caption         =   "Group/Index p1"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1635
   End
   Begin VB.TextBox txtGI1 
      Height          =   7575
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "frmHelpMonsters.frx":08CA
      Top             =   360
      Width           =   7455
   End
   Begin VB.TextBox txtBL 
      Height          =   7575
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmHelpMonsters.frx":54B5
      Top             =   360
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.TextBox txtGI2 
      Height          =   7575
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmHelpMonsters.frx":9845
      Top             =   360
      Visible         =   0   'False
      Width           =   7455
   End
End
Attribute VB_Name = "frmHelpMonsters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Private Sub cmdBL_Click()
txtBL.Visible = True
txtGI1.Visible = False
txtGI2.Visible = False
End Sub

Private Sub cmdGI1_Click()
txtGI1.Visible = True
txtGI2.Visible = False
txtBL.Visible = False
End Sub

Private Sub cmdGI2_Click()
txtGI2.Visible = True
txtGI1.Visible = False
txtBL.Visible = False
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = ReadINI("Windows", "HelpMonstersTop")
Me.Left = ReadINI("Windows", "HelpMonstersLeft")
Me.Width = ReadINI("Windows", "HelpMonstersWidth")
Me.Height = ReadINI("Windows", "HelpMonstersHeight")
Me.Show
Me.SetFocus
If ReadINI("Windows", "HelpMonstersMaxed") = "1" Then Me.WindowState = vbMaximized
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


With txtGI1
    .Move .Left, .Top, lUseWidth - 125, lUseHeight - TITLEBAR_OFFSET - 760
End With


With txtGI2
    .Move .Left, .Top, lUseWidth - 125, lUseHeight - TITLEBAR_OFFSET - 760
End With


With txtBL
    .Move .Left, .Top, lUseWidth - 125, lUseHeight - TITLEBAR_OFFSET - 760
End With

    
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
        If Me.WindowState = vbMinimized Then Exit Sub
        
        If Me.WindowState = vbMaximized Then
            Call WriteINI("Windows", "HelpMonstersMaxed", 1)
        Else
            Call WriteINI("Windows", "HelpMonstersMaxed", 0)
            Call WriteINI("Windows", "HelpMonstersTop", frmHelpMonsters.Top)
            Call WriteINI("Windows", "HelpMonstersLeft", frmHelpMonsters.Left)
            Call WriteINI("Windows", "HelpMonstersHeight", frmHelpMonsters.Height)
            Call WriteINI("Windows", "HelpMonstersWidth", frmHelpMonsters.Width)
        End If
End Sub

