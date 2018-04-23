VERSION 5.00
Begin VB.Form frmUserPaste 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paste Window"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7545
   ControlBox      =   0   'False
   Icon            =   "frmUserPaste.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clea&r"
      Height          =   315
      Left            =   5100
      TabIndex        =   3
      Top             =   0
      Width           =   1155
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "&Paste from Clipboard"
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   0
      Width           =   1995
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   6360
      TabIndex        =   4
      Top             =   0
      Width           =   1155
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Co&ntinue"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   1995
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   0
      MaxLength       =   10000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   7515
   End
End
Attribute VB_Name = "frmUserPaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdClear_Click()
txtText.Text = ""
DoEvents
End Sub

Private Sub cmdContinue_Click()
txtText.SetFocus
Me.Tag = "1"
Me.Hide
End Sub

Private Sub cmdPaste_Click()
Dim nYesNo As Integer

If Not Clipboard.GetText = "" Then
    If Not txtText.Text = "" Then
        nYesNo = MsgBox("Clear paste area first?", vbYesNo + vbDefaultButton1 + vbQuestion, "Clear?")
    Else
        nYesNo = vbYes
    End If
    
    If nYesNo = vbYes Then
        txtText.Text = Clipboard.GetText
    Else
        txtText.Text = Clipboard.GetText & vbCrLf & vbCrLf & txtText.Text
    End If
End If

End Sub

