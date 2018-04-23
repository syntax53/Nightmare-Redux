VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExpCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exp Calculator"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   Icon            =   "frmExpCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   4215
   Begin VB.CommandButton cmdNote 
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3300
      TabIndex        =   13
      Top             =   4440
      Width           =   795
   End
   Begin VB.TextBox txtEndLVL 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3600
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "255"
      Top             =   720
      Width           =   555
   End
   Begin VB.TextBox txtStartLVL 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "2"
      Top             =   720
      Width           =   555
   End
   Begin VB.ComboBox cmbRace 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   300
      Width           =   1515
   End
   Begin VB.ComboBox cmbClass 
      Height          =   315
      Left            =   60
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   300
      Width           =   1635
   End
   Begin VB.TextBox txtCalcEXPTable 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdCalcExp 
      Caption         =   "&Calc."
      Height          =   555
      Left            =   3420
      TabIndex        =   4
      Top             =   60
      Width           =   735
   End
   Begin MSComctlLib.ListView lvCalcExp 
      Height          =   3315
      Left            =   60
      TabIndex        =   10
      Top             =   1080
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5847
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3375
      TabIndex        =   12
      Top             =   780
      Width           =   195
   End
   Begin VB.Label Label4 
      Caption         =   "LVL Range:"
      Height          =   195
      Left            =   1800
      TabIndex        =   7
      Top             =   780
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Thanks to Locke Cole for the exp calc function"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   60
      TabIndex        =   11
      Top             =   4440
      Width           =   2880
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Race"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   60
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1635
   End
   Begin VB.Label Label39 
      Caption         =   "Exp Table %:"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   780
      Width           =   1035
   End
End
Attribute VB_Name = "frmExpCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Declare Function CalcExpNeeded Lib "lltmmudxp" (ByVal Level As Integer, ByVal Chart As Integer) As Currency

Public Sub CalcBy(nClass As Long, nRace As Long, nLevel As Long)
Dim x As Integer, bSuccess As Boolean
On Error GoTo Error:

bSuccess = True
For x = 0 To cmbClass.ListCount - 1
    If cmbClass.ItemData(x) = nClass Then
        cmbClass.ListIndex = x
        Exit For
    End If
Next x
If x = cmbClass.ListCount Then bSuccess = False

For x = 0 To cmbRace.ListCount - 1
    If cmbRace.ItemData(x) = nRace Then
        cmbRace.ListIndex = x
        Exit For
    End If
Next x
If x = cmbRace.ListCount Then bSuccess = False

txtStartLVL.Text = nLevel - 1
txtEndLVL.Text = nLevel + 50

If bSuccess Then
    Call cmdCalcExp_Click
    On Error Resume Next
    Set lvCalcExp.SelectedItem = lvCalcExp.ListItems(2)
    lvCalcExp.SelectedItem.EnsureVisible
    lvCalcExp.SetFocus
End If
out:
Exit Sub
Error:
Call HandleError("CalcBy")
Resume out:
End Sub

Private Sub cmdNote_Click()
MsgBox "Clicking on a line will copy the experience value for that level to your clipboard.", vbInformation
End Sub

Private Sub Form_Load()
Dim nStatus As Integer
On Error GoTo Error:

Me.Left = ReadINI("Windows", "ExpCalcLeft")
Me.Top = ReadINI("Windows", "ExpCalcTop")

cmbClass.clear
nStatus = BTRCALL(BGETFIRST, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
If nStatus = 0 Then
    Do While nStatus = 0
        Call ClassRowToStruct(Classdatabuf.buf)
        
        cmbClass.AddItem ClipNull(Classrec.Name)
        cmbClass.ItemData(cmbClass.NewIndex) = Classrec.Number
        
        nStatus = BTRCALL(BGETNEXT, ClassPosBlock, Classdatabuf, Len(Classdatabuf), ByVal ClassKeyBuffer, KEY_BUF_LEN, 0)
    Loop
Else
    MsgBox "Error getting first class: " & BtrieveErrorCode(nStatus)
End If
cmbClass.AddItem "custom", 0
cmbClass.ListIndex = 0

cmbRace.clear
nStatus = BTRCALL(BGETFIRST, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
If nStatus = 0 Then
    Do While nStatus = 0
        Call RaceRowToStruct(Racedatabuf.buf)
        
        cmbRace.AddItem ClipNull(Racerec.Name)
        cmbRace.ItemData(cmbRace.NewIndex) = Racerec.Number
        
        nStatus = BTRCALL(BGETNEXT, RacePosBlock, Racedatabuf, Len(Racedatabuf), ByVal RaceKeyBuffer, KEY_BUF_LEN, 0)
    Loop
Else
    MsgBox "Error getting first Race: " & BtrieveErrorCode(nStatus)
End If
cmbRace.AddItem "custom", 0
cmbRace.ListIndex = 0

txtStartLVL.Text = ReadINI("Options", "ExpCalcStartLevel")
txtEndLVL.Text = ReadINI("Options", "ExpCalcEndLevel")
If Val(txtEndLVL.Text) < 10 Then txtEndLVL.Text = 255

Exit Sub
Error:
Call HandleError
Resume Next
End Sub

Private Sub CalcExp()
Dim nClassExp As Integer, nRaceExp As Integer, nStatus As Integer, nRecord As Integer

On Error GoTo Error:

If cmbClass.ListIndex > 0 Then
    nRecord = cmbClass.ItemData(cmbClass.ListIndex)
    nStatus = BTRCALL(BGETEQUAL, ClassPosBlock, Classdatabuf, Len(Classdatabuf), nRecord, KEY_BUF_LEN, 0)
    Call ClassRowToStruct(Classdatabuf.buf)
    If nStatus = 0 Then
        nClassExp = Classrec.Exp + 100
    Else
        MsgBox "Error getting class: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If

If cmbRace.ListIndex > 0 Then
    nRecord = cmbRace.ItemData(cmbRace.ListIndex)
    nStatus = BTRCALL(BGETEQUAL, RacePosBlock, Racedatabuf, Len(Racedatabuf), nRecord, KEY_BUF_LEN, 0)
    Call RaceRowToStruct(Racedatabuf.buf)
    If nStatus = 0 Then
        nRaceExp = Racerec.ExpChart
    Else
        MsgBox "Error getting Race: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
End If

txtCalcEXPTable.Text = nClassExp + nRaceExp

Exit Sub
Error:
Call HandleError
End Sub

Private Sub cmbClass_Click()
Call CalcExp
End Sub

Private Sub cmbRace_Click()
Call CalcExp
End Sub

Private Sub cmdCalcExp_Click()
Dim sExp As String, nExp As Currency, x As Long
Dim oLI As ListItem, nLastExp As Currency

On Error GoTo Error:

lvCalcExp.ListItems.clear
lvCalcExp.ColumnHeaders.clear
lvCalcExp.ColumnHeaders.add , , "LVL", 500
lvCalcExp.ColumnHeaders.add , , "Experience", 1600
lvCalcExp.ColumnHeaders.add , , "Needed", 1400

If Val(txtStartLVL.Text) < 2 Then
    txtStartLVL.Text = 2
ElseIf Val(txtStartLVL.Text) > 500 Then
    txtStartLVL.Text = 500
End If

If Val(txtEndLVL.Text) < 10 Then
    txtEndLVL.Text = 10
ElseIf Val(txtEndLVL.Text) > 999 Then
    txtEndLVL.Text = 999
End If

For x = Val(txtStartLVL.Text) To Val(txtEndLVL.Text)
    nExp = CalcExpNeeded(x, CInt(txtCalcEXPTable.Text))
    sExp = CStr(nExp * 10000)

    Set oLI = lvCalcExp.ListItems.add()
    oLI.Text = x
    oLI.SubItems(1) = PutCommas(sExp)
    oLI.SubItems(2) = PutCommas(Val(sExp) - nLastExp)

    nLastExp = Val(sExp)
    Set oLI = Nothing
Next

Exit Sub

Error:
Call HandleError

End Sub



Private Sub Form_Unload(Cancel As Integer)
Call WriteINI("Windows", "ExpCalcLeft", Me.Left)
Call WriteINI("Windows", "ExpCalcTop", Me.Top)
Call WriteINI("Options", "ExpCalcStartLevel", Val(txtStartLVL.Text))
Call WriteINI("Options", "ExpCalcEndLevel", Val(txtEndLVL.Text))
End Sub

Private Sub lvCalcExp_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim sArr() As String, x As Integer, sExp As String
On Error GoTo Error:

If InStr(1, Item.ListSubItems(1).Text, ",") > 0 Then
    sArr() = Split(Item.ListSubItems(1).Text, ",")
Else
    ReDim sArr(0)
    sArr(0) = Item.ListSubItems(1).Text
End If

For x = 0 To UBound(sArr())
    sExp = sExp & sArr(x)
Next x

Clipboard.clear
Clipboard.SetText sExp

out:
Exit Sub
Error:
Call HandleError("lvCalcExp_ItemClick")
Resume out:
End Sub

Private Sub txtCalcEXPTable_GotFocus()
Call SelectAll(txtCalcEXPTable)
End Sub

Private Sub txtCalcEXPTable_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtCalcEXPTable_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sStr As String, nPos As Integer, nSel As Integer

On Error GoTo Error:

nPos = txtCalcEXPTable.SelStart
sStr = txtCalcEXPTable.Text
nSel = txtCalcEXPTable.SelLength

cmbClass.ListIndex = 0
cmbRace.ListIndex = 0
txtCalcEXPTable.Text = sStr
txtCalcEXPTable.SelStart = nPos
txtCalcEXPTable.SelLength = nSel

Exit Sub

Error:
Call HandleError

End Sub

Private Sub txtEndLVL_GotFocus()
Call SelectAll(txtEndLVL)
End Sub

Private Sub txtEndLVL_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub

Private Sub txtStartLVL_GotFocus()
Call SelectAll(txtStartLVL)
End Sub

Private Sub txtStartLVL_KeyPress(KeyAscii As Integer)
KeyAscii = NumberKeysOnly(KeyAscii)
End Sub
