VERSION 5.00
Begin VB.Form frmUserSelectItem 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Duplicate Record Match"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frmUserSelectItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2595
      Left            =   60
      TabIndex        =   0
      Top             =   -60
      Width           =   4515
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   2100
         Width           =   1575
      End
      Begin VB.ListBox lstItems 
         Height          =   1230
         ItemData        =   "frmUserSelectItem.frx":08CA
         Left            =   120
         List            =   "frmUserSelectItem.frx":08CC
         TabIndex        =   1
         Top             =   720
         Width           =   4275
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Please choose which record to use.  You can double click to view the record."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4275
      End
   End
End
Attribute VB_Name = "frmUserSelectItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'0 = items
'1 = spells
Public FormJump As Integer

Private Sub cmdOK_Click()
Me.Tag = "1"
Me.Hide
End Sub

Private Sub lstItems_DblClick()
Select Case FormJump
    Case 0: 'items
        Call frmItem.GotoItem(lstItems.ItemData(lstItems.ListIndex))
    Case 1: 'spells
        Call frmSpell.GotoSpell(lstItems.ItemData(lstItems.ListIndex))
End Select

End Sub
