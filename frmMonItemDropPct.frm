VERSION 5.00
Begin VB.Form frmMonsterItemDropPct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monster Item Drop Percentage Modifier"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2055
   ScaleWidth      =   7935
   Begin VB.ComboBox cmbOnlyIfModifier 
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
      ItemData        =   "frmMonItemDropPct.frx":0000
      Left            =   60
      List            =   "frmMonItemDropPct.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1200
      Width           =   1035
   End
   Begin VB.ComboBox cmbOption 
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
      ItemData        =   "frmMonItemDropPct.frx":002E
      Left            =   60
      List            =   "frmMonItemDropPct.frx":0038
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   11
      Top             =   1020
      Width           =   1155
   End
   Begin VB.CheckBox chkLogOnly 
      Caption         =   "Log Only - No Changes"
      Height          =   255
      Left            =   5820
      TabIndex        =   10
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1995
   End
   Begin VB.CheckBox chkItemTypes 
      Caption         =   "Scrolls"
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   9
      Top             =   780
      Width           =   915
   End
   Begin VB.CheckBox chkItemTypes 
      Caption         =   "Containers"
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.CheckBox chkItemTypes 
      Caption         =   "Keys"
      Height          =   255
      Index           =   2
      Left            =   5880
      TabIndex        =   6
      Top             =   480
      Width           =   795
   End
   Begin VB.CheckBox chkItemTypes 
      Caption         =   "Armour "
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   5
      Top             =   780
      Width           =   915
   End
   Begin VB.CheckBox chkItemTypes 
      Caption         =   "Weapons"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtPctLimit 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "20"
      Top             =   1200
      Width           =   795
   End
   Begin VB.TextBox txtDropPct 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2580
      TabIndex        =   0
      Text            =   "20"
      Top             =   360
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Value"
      Height          =   195
      Index           =   5
      Left            =   2580
      TabIndex        =   17
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Choose option..."
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   1680
      Width           =   2115
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Items with abil-119 Del@maint will be excluded"
      Height          =   435
      Index           =   3
      Left            =   3480
      TabIndex        =   12
      Top             =   1140
      Width           =   2115
   End
   Begin VB.Label lblExtra 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   1620
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Do it for these item types:"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Condition: Only if drop % is..."
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   900
      Width           =   2055
   End
End
Attribute VB_Name = "frmMonsterItemDropPct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Private Sub cmbOnlyIfModifier_Click()
lblExtra.Caption = ""
Select Case cmbOnlyIfModifier.ListIndex
    Case 0: '=
    Case 1: ' <=
        lblExtra.Caption = "(AND greater than 0)"
    Case 2: ' >=
    Case 3: ' !=
End Select
End Sub

Private Sub cmdGo_Click()
Dim fso As FileSystemObject, fil As String, ts As TextStream
On Error GoTo error:

If Val(txtPctLimit.Text) > 100 Then txtPctLimit.Text = 100
If Val(txtPctLimit.Text) < 1 Then txtPctLimit.Text = 0

Set fso = CreateObject("Scripting.FileSystemObject")
If Right(App.Path, 1) = "\" Then
    fil = App.Path & "NMR-Log_MonItemPct.txt"
Else
    fil = App.Path & "\NMR-Log_MonItemPct.txt"
End If
If fso.FileExists(fil) = True Then fso.DeleteFile fil, True
Set ts = fso.OpenTextFile(fil, ForWriting, True)

ts.WriteLine ("Monster Item Drop Percentage Modifier " & Date & " @ " & Time)
If chkLogOnly.Value = 1 Then ts.WriteLine ("** LOGGING ONLY, NO CHANGES EXECUTED **")
ts.WriteBlankLines (1)

out:
On Error Resume Next
ts.WriteBlankLines (1)
ts.WriteLine ("Complete - " & Date & " @ " & Time)
ts.Close
Exit Sub
error:
Call HandleError("cmdGo_Click")
Resume out: End Sub

Private Sub Form_Load()
On Error Resume Next
Dim nStatus As Integer

Me.Top = ReadINI("Windows", "MonItemPctTop")
Me.Left = ReadINI("Windows", "MonItemPctLeft")
cmbOption.ListIndex = 0
cmbOnlyIfModifier.ListIndex = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
        If Me.WindowState = vbMinimized Then Exit Sub
        Call WriteINI("Windows", "MonItemPctTop", frmMonsterItemDropPct.Top)
        Call WriteINI("Windows", "MonItemPctLeft", frmMonsterItemDropPct.Left)
End Sub

Private Sub txtDropPct_Change()
If Val(txtDropPct.Text) > 100 Then txtDropPct.Text = 100
If Val(txtDropPct.Text) < -100 Then txtDropPct.Text = -100
End Sub

Private Sub txtPctLimit_Change()
If Val(txtPctLimit.Text) > 100 Then txtPctLimit.Text = 100
If Val(txtPctLimit.Text) < 1 Then txtPctLimit.Text = 0
End Sub

Private Sub txtPctLimit_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtDropPct_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
