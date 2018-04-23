VERSION 5.00
Begin VB.Form frmBank 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bankbook Editor"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "frmBank.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1635
   ScaleWidth      =   6765
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Next"
      Height          =   315
      Index           =   1
      Left            =   5280
      TabIndex        =   21
      Top             =   1140
      Width           =   1335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find BBS Name"
      Height          =   315
      Index           =   0
      Left            =   3780
      TabIndex        =   20
      Top             =   1140
      Width           =   1515
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last"
      Height          =   315
      Left            =   2460
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   795
   End
   Begin VB.TextBox txtGotoShopNumber 
      Height          =   315
      Left            =   5640
      TabIndex        =   18
      Top             =   1140
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAmount 
      Height          =   315
      Left            =   1080
      TabIndex        =   14
      Top             =   1140
      Width           =   1815
   End
   Begin VB.TextBox txtShopNumber 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   3780
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   660
      Width           =   735
   End
   Begin VB.TextBox txtShopName 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   660
      Width           =   2055
   End
   Begin VB.CommandButton cmdDiscard 
      Caption         =   "Dis&card"
      Height          =   315
      Left            =   5820
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   5040
      TabIndex        =   6
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "&Goto"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6420
      TabIndex        =   19
      Top             =   1140
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtGotoBBSName 
      Height          =   315
      Left            =   4020
      TabIndex        =   16
      Top             =   1140
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "&First"
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   675
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   315
      Left            =   900
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   315
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtBBSName 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   660
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Shop #"
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   1140
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label label 
      Caption         =   "BBS Name"
      Height          =   255
      Index           =   1
      Left            =   3060
      TabIndex        =   15
      Top             =   1140
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Amount"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1140
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Shop #"
      Height          =   255
      Left            =   3060
      TabIndex        =   10
      Top             =   660
      Width           =   615
   End
   Begin VB.Label label 
      Caption         =   "BBS Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   660
      Width           =   855
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim sFind As String
Dim bLoaded As Boolean

Private Sub cmdFind_Click(Index As Integer)
Dim nStatus As Integer, sTemp As String
On Error GoTo error:

If bLoaded Then Call SaveBank

If Index = 0 Then
    sTemp = InputBox("Enter BBS Name to Find", "Find BBS Name", sFind)
    If sTemp = "" Then Exit Sub
    sFind = sTemp
    
    nStatus = BTRCALL(BGETFIRST, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "cmdFirst_Click, BGETFIRST, Bank, Error: " & BtrieveErrorCode(nStatus)
        bLoaded = False
        Exit Sub
    End If
Else
    If sFind = "" Then
        sTemp = InputBox("Enter BBS Name to Find", "Find BBS Name", sFind)
        If sTemp = "" Then Exit Sub
        sFind = sTemp
    End If
    
    nStatus = BTRCALL(BGETNEXT, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "BGETNEXT, Bank, Error: " & BtrieveErrorCode(nStatus)
        bLoaded = False
        Exit Sub
    End If
End If

Do While nStatus = 0
    Call BankRowToStruct(BankDatabuf.buf)
    
    If InStr(1, LCase(Bankrec.BBSName), LCase(sFind)) > 0 Then
        'MsgBox "Found.", vbInformation
        Call DispBankInfo(BankDatabuf.buf)
        Exit Sub
    End If
    
    nStatus = BTRCALL(BGETNEXT, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
Loop

If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "BGETNEXT, Bank, Error: " & BtrieveErrorCode(nStatus)
    bLoaded = False
    Exit Sub
End If

nStatus = BTRCALL(BGETLAST, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "BGETLAST, Bank, Error: " & BtrieveErrorCode(nStatus)
    bLoaded = False
    Exit Sub
End If

MsgBox "Not Found.", vbInformation
Call DispBankInfo(BankDatabuf.buf)

out:
Exit Sub
error:
Call HandleError("cmdFind_Click")
Resume out:

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim nStatus As Integer

Me.Top = ReadINI("Windows", "BankTop")
Me.Left = ReadINI("Windows", "BankLeft")

nStatus = BTRCALL(BGETFIRST, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadBank, BGETFIRST, Bank, Error: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    bLoaded = True
    DispBankInfo BankDatabuf.buf
End If
Me.Show
Me.SetFocus
cmdNext.SetFocus

End Sub
Private Sub cmdDiscard_Click()
'Dim nStatus As Integer
'
'BankKey.BBSName = txtBBSName.Text & String(Len(BankKey.BBSName) - Len(txtBBSName.Text), vbNullChar)
'BankKey.ShopNumber = 8
'
'nStatus = BTRCALL(BGETGREATER, BankPosBlock, BankDatabuf, Len(BankDatabuf), BankKey, KEY_BUF_LEN, 0)
'If Not nStatus = 0 Then
'    MsgBox "cmdGoto_Click(), BGETEQUAL, Bank, Error: " & BtrieveErrorCode(nStatus)
'Else
'    DispBankInfo BankDatabuf.buf
'End If
Dim nStatus As Integer

nStatus = BTRCALL(BGETFIRST, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdFirst_Click, BGETFIRST, Bank, Error: " & BtrieveErrorCode(nStatus)
Else
    DispBankInfo BankDatabuf.buf
End If
End Sub

Private Sub cmdSave_Click()
If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
Call SaveBank
End Sub


Private Sub DispBankInfo(row() As Byte)
On Error GoTo error:
bLoaded = True
RowToStruct row, BankFldMap, Bankrec, LenB(Bankrec)

txtBBSName.Text = Bankrec.BBSName
txtShopNumber.Text = SLong2ULong(Bankrec.ShopNumber)
txtShopName.Text = GetShopName(SLong2ULong(Bankrec.ShopNumber))
txtAmount.Text = SLong2ULong(Bankrec.Cash)

Exit Sub
error:
Call HandleError
MsgBox "Warning, record was not completely displayed." & vbCrLf _
    & "Previous records stats may still be in memory.  Select 'Disable DB Writing'" & vbCrLf _
    & "from the file menu and then reload the editor.", vbExclamation
End Sub


Private Sub SaveBank()
On Error GoTo error:
Dim nStatus As Integer
'DoEvents
Bankrec.Cash = ULong2SLong(Val(txtAmount.Text))

nStatus = UpdateBank
If Not nStatus = 0 Then
    MsgBox "cmd_Save, Error: " & BtrieveErrorCode(nStatus)
Else
    DispBankInfo BankDatabuf.buf
End If
Exit Sub
error:
Call HandleError
End Sub


Private Sub cmdGoto_Click()
On Error GoTo error:
Dim nStatus As Integer, nGoto As Long, temp As String
If bLoaded = True Then SaveBank

BankKey.BBSName = txtGotoBBSName.Text & String(30 - Len(txtGotoBBSName.Text), " ")
BankKey.ShopNumber = Val(txtGotoShopNumber.Text)

BankKeyStructToRow BankKeyDataBuf.buf
'BankStructToRow BankDatabuf.buf

nStatus = BTRCALL(BGETEQUAL, BankPosBlock, BankDatabuf, Len(BankDatabuf), BankKeyDataBuf, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdGoto_Click(), BGETEQUAL, Bank, Error: " & BtrieveErrorCode(nStatus)
Else
    DispBankInfo BankDatabuf.buf
End If

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdFirst_Click()
Dim nStatus As Integer
If bLoaded Then SaveBank
nStatus = BTRCALL(BGETFIRST, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdFirst_Click, BGETFIRST, Bank, Error: " & BtrieveErrorCode(nStatus)
Else
    DispBankInfo BankDatabuf.buf
End If
End Sub

Private Sub cmdNext_Click()
Dim nStatus As Integer
If bLoaded Then SaveBank
nStatus = BTRCALL(BGETNEXT, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdNext_Click, BGETNEXT, Bank, Error: " & BtrieveErrorCode(nStatus)
Else
    DispBankInfo BankDatabuf.buf
End If
End Sub

Private Sub cmdPrevious_Click()
Dim nStatus As Integer
If bLoaded Then SaveBank
nStatus = BTRCALL(BGETPREVIOUS, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdPrevious_Click, BGETPREVIOUS, Bank, Error: " & BtrieveErrorCode(nStatus)
Else
    DispBankInfo BankDatabuf.buf
End If
End Sub

Private Sub cmdLast_Click()
Dim nStatus As Integer
If bLoaded Then SaveBank
nStatus = BTRCALL(BGETLAST, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdLast_Click, BGETLAST, Bank, Error: " & BtrieveErrorCode(nStatus)
Else
    DispBankInfo BankDatabuf.buf
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
        If bLoaded = True Then SaveBank
        If Me.WindowState = vbMinimized Then Exit Sub
        Call WriteINI("Windows", "BankTop", frmBank.Top)
        Call WriteINI("Windows", "BankLeft", frmBank.Left)
End Sub

Private Sub cmdDelete_Click()
Dim nStatus As Integer
Dim nDelete As Integer
If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
If bLoaded Then SaveBank
nDelete = MsgBox("Delete this record from database?", vbYesNo, "Delete Record?")
If nDelete = 6 Then
    nStatus = BTRCALL(BDELETE, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "cmdDelete, BDELETE, Error: " & BtrieveErrorCode(nStatus)
    Else
        Form_Load
    End If
End If
End Sub

Private Sub cmdInsert_Click()
On Error GoTo error:
Dim nStatus As Integer
Dim BBSName As String * 30, ShopNumber As Integer, temp As String
If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
If bLoaded = True Then SaveBank

BBSName = InputBox("Enter *BBS Name*" & vbCrLf & vbCrLf & "NOTE: Enter it EXACTLY, Case Sensitive!", "Bank Insert", "")
If BBSName = String(30, " ") Then Exit Sub

temp = InputBox("Enter Shop (bank) Number", "Bank Insert", "8")
If temp = "" Then Exit Sub

BBSName = BBSName & String(Len(Bankrec.BBSName) - Len(BBSName), vbNullChar)
ShopNumber = ULong2SLong(Val(temp))

    Bankrec.BBSName = BBSName
    Bankrec.ShopNumber = ShopNumber
    Bankrec.Cash = 1
    
    BankStructToRow BankDatabuf.buf
    nStatus = BTRCALL(BINSERT, BankPosBlock, BankDatabuf, Len(BankDatabuf), ByVal BankKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "cmdInsert, BINSERT, Error: " & BtrieveErrorCode(nStatus)
    Else
        DispBankInfo BankDatabuf.buf
    End If
Exit Sub
error:
Call HandleError
End Sub


Private Sub txtAmount_GotFocus()
Call SelectAll(txtAmount)

End Sub
