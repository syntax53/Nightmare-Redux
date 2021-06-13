VERSION 5.00
Begin VB.Form frmGang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gang Editor"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmGang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2100
   ScaleWidth      =   7470
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   1020
      TabIndex        =   0
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txtRollTimes 
      Height          =   285
      Left            =   1380
      TabIndex        =   15
      Top             =   1680
      Width           =   795
   End
   Begin VB.TextBox txtRollOverExp 
      Height          =   285
      Left            =   1380
      TabIndex        =   14
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtLeader 
      Height          =   285
      Left            =   5100
      MaxLength       =   30
      TabIndex        =   19
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   5100
      TabIndex        =   20
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtMembers 
      Height          =   285
      Left            =   5100
      TabIndex        =   21
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtExp 
      Height          =   285
      Left            =   1380
      TabIndex        =   13
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtDisplayName 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   5100
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdDiscard 
      Caption         =   "Dis&card"
      Height          =   375
      Left            =   6540
      TabIndex        =   7
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   5700
      TabIndex        =   6
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "&First"
      Height          =   375
      Left            =   2100
      TabIndex        =   2
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   60
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "MM/DD/YYYY"
      Height          =   195
      Left            =   6180
      TabIndex        =   22
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Times Rolled"
      Height          =   255
      Index           =   7
      Left            =   60
      TabIndex        =   24
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Created"
      Height          =   255
      Index           =   6
      Left            =   3780
      TabIndex        =   18
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Roll Over Exp"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Members"
      Height          =   255
      Index           =   4
      Left            =   3780
      TabIndex        =   23
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Leader"
      Height          =   255
      Index           =   3
      Left            =   3780
      TabIndex        =   16
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Base Exp"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Display Name"
      Height          =   255
      Index           =   1
      Left            =   3780
      TabIndex        =   10
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "KeyName"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmGang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim Loaded As Boolean

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim nStatus As Integer

Me.Top = ReadINI("Windows", "GangTop")
Me.Left = ReadINI("Windows", "GangLeft")

nStatus = BTRCALL(BGETFIRST, GangPosBlock, GangDatabuf, Len(GangDatabuf), ByVal GangKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadGang, BGETFIRST, Gang, Error: " & BtrieveErrorCode(nStatus)
    Loaded = False
Else
    Loaded = True
    DispGangInfo GangDatabuf.buf
End If

End Sub
Private Sub cmdDiscard_Click()
'Dim nStatus As Integer
'
'GangKey.BBSName = txtBBSName.Text & String(Len(GangKey.BBSName) - Len(txtBBSName.Text), vbNullChar)
'GangKey.ShopNumber = 8
'
'nStatus = BTRCALL(BGETGREATER, GangPosBlock, GangDatabuf, Len(GangDatabuf), GangKey, KEY_BUF_LEN, 0)
'If Not nStatus = 0 Then
'    MsgBox "cmdGoto_Click(), BGETEQUAL, Gang, Error: " & BtrieveErrorCode(nStatus)
'Else
'    DispGangInfo GangDatabuf.buf
'End If
Dim nStatus As Integer

nStatus = BTRCALL(BGETFIRST, GangPosBlock, GangDatabuf, Len(GangDatabuf), ByVal GangKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdFirst_Click, BGETFIRST, Gang, Error: " & BtrieveErrorCode(nStatus)
Else
    DispGangInfo GangDatabuf.buf
End If

End Sub

Private Sub cmdSave_Click()
If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
SaveGang
End Sub


Private Sub DispGangInfo(row() As Byte)
On Error GoTo Error:
Loaded = True
RowToStruct row, GangFldMap, Gangrec, LenB(Gangrec)
    
txtName.Text = Gangrec.KeyName
txtDisplayName.Text = Gangrec.DisplayName
txtMembers.Text = Gangrec.Members
txtDate.Text = DOSDate2Date(Gangrec.DateCreated)
txtLeader.Text = Gangrec.Leader
txtExp.Text = SLong2ULong(Gangrec.Exp)
txtRollOverExp.Text = SLong2ULong(Gangrec.RollOver)
txtRollTimes = SLong2ULong(Gangrec.RollTimes)

Exit Sub
Error:
Call HandleError
MsgBox "Warning, record was not completely displayed." & vbCrLf _
    & "Previous records stats may still be in memory.  Select 'Disable DB Writing'" & vbCrLf _
    & "from the file menu and then reload the editor.", vbExclamation
End Sub


Private Sub SaveGang()
On Error GoTo Error:
Dim nStatus As Integer, temp As Long
'DoEvents
Gangrec.DisplayName = txtDisplayName.Text & String(20 - Len(txtDisplayName.Text), Chr(0))
Gangrec.Members = Val(txtMembers.Text)
Gangrec.Leader = txtLeader.Text & String(30 - Len(txtLeader.Text), Chr(0))
Gangrec.KeyName = UCase(txtName.Text) & String(20 - Len(txtName.Text), Chr(0))
Gangrec.Exp = ULong2SLong(Val(txtExp.Text))
Gangrec.RollOver = ULong2SLong(Val(txtRollOverExp.Text))
Gangrec.RollTimes = ULong2SLong(Val(txtRollTimes.Text))

temp = Date2DOSDate(txtDate.Text)
If Not temp = -1 Then
    Gangrec.DateCreated = UInt2SInt(temp)
End If

nStatus = UpdateGang
If Not nStatus = 0 Then
    MsgBox "cmd_Save, Error: " & BtrieveErrorCode(nStatus)
Else
    DispGangInfo GangDatabuf.buf
End If

Exit Sub
Error:
Call HandleError
End Sub


'Private Sub cmdGoto_Click()
'On Error GoTo error:
'Dim nStatus As Integer, sGoto As String, x As Integer
'If Loaded = True Then SaveGang
'
'sGoto = UCase(txtGoto.Text) & String(20 - Len(txtGoto.Text), Chr(0))
''sGoto = sGoto & txtGoto.Text & String(20 - Len(txtGoto.Text), Chr(0))
'
'nStatus = BTRCALL(BGETEQUAL, GangPosBlock, GangDatabuf, Len(GangDatabuf), sGoto, KEY_BUF_LEN, 0)
'If Not nStatus = 0 Then
'    MsgBox "cmdGoto_Click(), BGETEQUAL, Gang, Error: " & BtrieveErrorCode(nStatus)
'Else
'    DispGangInfo GangDatabuf.buf
'End If
'
'Exit Sub
'error:
'Call HandleError
'End Sub

Private Sub cmdFirst_Click()
Dim nStatus As Integer
SaveGang
nStatus = BTRCALL(BGETFIRST, GangPosBlock, GangDatabuf, Len(GangDatabuf), ByVal GangKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdFirst_Click, BGETFIRST, Gang, Error: " & BtrieveErrorCode(nStatus)
Else
    DispGangInfo GangDatabuf.buf
End If
End Sub

Private Sub cmdNext_Click()
Dim nStatus As Integer
SaveGang
nStatus = BTRCALL(BGETNEXT, GangPosBlock, GangDatabuf, Len(GangDatabuf), ByVal GangKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdNext_Click, BGETNEXT, Gang, Error: " & BtrieveErrorCode(nStatus)
Else
    DispGangInfo GangDatabuf.buf
End If
End Sub

Private Sub cmdPrevious_Click()
Dim nStatus As Integer
SaveGang
nStatus = BTRCALL(BGETPREVIOUS, GangPosBlock, GangDatabuf, Len(GangDatabuf), ByVal GangKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdPrevious_Click, BGETPREVIOUS, Gang, Error: " & BtrieveErrorCode(nStatus)
Else
    DispGangInfo GangDatabuf.buf
End If
End Sub

Private Sub cmdLast_Click()
Dim nStatus As Integer
SaveGang
nStatus = BTRCALL(BGETLAST, GangPosBlock, GangDatabuf, Len(GangDatabuf), ByVal GangKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdLast_Click, BGETLAST, Gang, Error: " & BtrieveErrorCode(nStatus)
Else
    DispGangInfo GangDatabuf.buf
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
        If Loaded = True Then SaveGang
        If Me.WindowState = vbMinimized Then Exit Sub
        Call WriteINI("Windows", "GangTop", frmGang.Top)
        Call WriteINI("Windows", "GangLeft", frmGang.Left)
End Sub

Private Sub cmdDelete_Click()
Dim nStatus As Integer
Dim nDelete As Integer
If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
SaveGang
nDelete = MsgBox("Delete this record from database?", vbYesNo, "Delete Record?")
If nDelete = 6 Then
    nStatus = BTRCALL(BDELETE, GangPosBlock, GangDatabuf, Len(GangDatabuf), ByVal GangKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "cmdDelete, BDELETE, Error: " & BtrieveErrorCode(nStatus)
    Else
        Form_Load
    End If
End If
End Sub

Private Sub cmdInsert_Click()
On Error GoTo Error:
Dim nStatus As Integer
Dim GangName As String, ShopNumber As Integer, temp As String
If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
If Loaded = True Then SaveGang

retry:
GangName = InputBox("Enter GangName (max 20 chars):", "Gang Insert", GangName)
If GangName = "" Then Exit Sub
If Len(GangName) > 20 Then GoTo retry:

    Gangrec.KeyName = UCase(GangName) & String(20 - Len(GangName), Chr(0))
    Gangrec.DisplayName = GangName & String(20 - Len(GangName), Chr(0))
    
    GangStructToRow GangDatabuf.buf
    nStatus = BTRCALL(BINSERT, GangPosBlock, GangDatabuf, Len(GangDatabuf), ByVal GangKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "cmdInsert, BINSERT, Error: " & BtrieveErrorCode(nStatus)
    Else
        DispGangInfo GangDatabuf.buf
    End If
Exit Sub
Error:
Call HandleError
End Sub


Private Sub txtDate_GotFocus()
Call SelectAll(txtDate)

End Sub

Private Sub txtDisplayName_GotFocus()
Call SelectAll(txtDisplayName)

End Sub

Private Sub txtExp_GotFocus()
Call SelectAll(txtExp)

End Sub

Private Sub txtLeader_GotFocus()
Call SelectAll(txtLeader)

End Sub

Private Sub txtMembers_GotFocus()
Call SelectAll(txtMembers)

End Sub

Private Sub txtName_GotFocus()
Call SelectAll(txtName)

End Sub

Private Sub txtRollOverExp_GotFocus()
Call SelectAll(txtRollOverExp)

End Sub

Private Sub txtRollTimes_GotFocus()
Call SelectAll(txtRollTimes)

End Sub
