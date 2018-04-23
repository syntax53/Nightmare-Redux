VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message Editor"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2655
   ScaleWidth      =   6540
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear"
      Height          =   315
      Left            =   3840
      TabIndex        =   12
      Top             =   480
      Width           =   795
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Ne&xt"
      Height          =   315
      Index           =   1
      Left            =   2580
      TabIndex        =   11
      Top             =   480
      Width           =   795
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   315
      Index           =   0
      Left            =   1740
      TabIndex        =   10
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Last"
      Height          =   315
      Left            =   5760
      TabIndex        =   7
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   4200
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdDiscard 
      Caption         =   "Dis&card"
      Height          =   315
      Left            =   5760
      TabIndex        =   14
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   5040
      TabIndex        =   13
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "&Goto"
      Default         =   -1  'True
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   795
   End
   Begin VB.TextBox txtGoto 
      Height          =   315
      Left            =   960
      MaxLength       =   5
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First"
      Height          =   315
      Left            =   5040
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   315
      Left            =   2580
      TabIndex        =   3
      Top             =   0
      Width           =   795
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Height          =   315
      Left            =   1740
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   315
      Left            =   3480
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtMessageLine3 
      Height          =   315
      Left            =   60
      MaxLength       =   74
      TabIndex        =   20
      Top             =   2280
      Width           =   6375
   End
   Begin VB.TextBox txtMessageLine2 
      Height          =   315
      Left            =   60
      MaxLength       =   74
      TabIndex        =   18
      Top             =   1680
      Width           =   6375
   End
   Begin VB.TextBox txtMessageLine1 
      Height          =   315
      Left            =   60
      MaxLength       =   74
      TabIndex        =   16
      Top             =   1080
      Width           =   6375
   End
   Begin VB.TextBox txtNumber 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   615
   End
   Begin VB.Label label 
      Caption         =   "Line 3"
      Height          =   315
      Index           =   133
      Left            =   75
      TabIndex        =   19
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label label 
      Caption         =   "Line 2"
      Height          =   315
      Index           =   131
      Left            =   75
      TabIndex        =   17
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label label 
      Caption         =   "Line 1"
      Height          =   315
      Index           =   37
      Left            =   75
      TabIndex        =   15
      Top             =   840
      Width           =   855
   End
   Begin VB.Label label 
      Caption         =   "Number"
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
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim bStopSearch As Boolean
Dim sFindText As String
Dim bLoaded As Boolean

Public Sub StopSearch()
    bStopSearch = True
End Sub

Private Sub cmdClearAll_Click()
txtMessageLine1.Text = ""
txtMessageLine2.Text = ""
txtMessageLine3.Text = ""
End Sub

Private Sub cmdFind_Click(Index As Integer)
On Error GoTo error:
Dim nStatus As Integer, x As Integer, decrypted As String, sTemp As String, nRecord As Long

bStopSearch = False
If Index = 0 Then 'find
    sTemp = InputBox("Enter String to search for", "Enter search string", sFindText)
    If sTemp = "" Then Exit Sub
        
    sFindText = sTemp
    
    nStatus = BTRCALL(BGETFIRST, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "GETFIRST Message, Error: " & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
    
Else 'find next
    If sFindText = "" Then
        sTemp = InputBox("Enter String to search for", "Enter search string", sFindText)
        If sTemp = "" Then Exit Sub
        sFindText = sTemp
    End If
    
    nRecord = Val(txtNumber.Text)
    nStatus = BTRCALL(BGETEQUAL, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), nRecord, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "Error getting current record." & BtrieveErrorCode(nStatus)
        Exit Sub
    End If
    
    nStatus = BTRCALL(BGETNEXT, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        If nStatus = 9 Then
            MsgBox "You are at the last record."
            Exit Sub
        Else
            MsgBox "Couldn't get next record -- Error: " & BtrieveErrorCode(nStatus)
            Exit Sub
        End If
    End If
End If

frmProgressBar.sCaption = "Message Search"
frmProgressBar.lblCaption.Caption = "Searching ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.ProgressBar.Value = 0
Set frmProgressBar.FormOwner = Me

nStatus = BTRCALL(BSTAT, MessagePosBlock, DBStatDatabuf, Len(Messagedatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    Call frmProgressBar.SetRange(3000)
    frmProgressBar.ProgressBar.Value = 1
Else
    DBStatRowToStruct DBStatDatabuf.buf
    Call frmProgressBar.SetRange(DBStat.nRecords)
End If

frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & strDatSuffix_TEXT
frmProgressBar.lblPanel(1).Caption = Messagerec.Number
frmProgressBar.Show
frmMain.Enabled = False
DoEvents


Do While nStatus = 0
    If bStopSearch Then GoTo canceled:
    
    MessageRowToStruct Messagedatabuf.buf
    frmProgressBar.IncreaseProgress
    frmProgressBar.lblPanel(1).Caption = Messagerec.Number
    
    If InStr(1, LCase(ClipNull(Messagerec.MessageLine1)), LCase(sFindText)) > 0 Then GoTo found:
    If InStr(1, LCase(ClipNull(Messagerec.MessageLine2)), LCase(sFindText)) > 0 Then GoTo found:
    If InStr(1, LCase(ClipNull(Messagerec.MessageLine3)), LCase(sFindText)) > 0 Then GoTo found:
    
    nStatus = BTRCALL(BGETNEXT, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
    DoEvents
Loop

GoTo notfound:

found:
'MsgBox "Found.", vbInformation
Unload frmProgressBar
DoEvents
frmMain.Enabled = True
frmMain.SetFocus
Call DispMessageInfo(Messagedatabuf.buf)
Exit Sub

notfound:
MsgBox "String not found.", vbInformation
canceled:
Unload frmProgressBar
DoEvents
frmMain.Enabled = True
frmMain.SetFocus
nRecord = Val(txtNumber.Text)
nStatus = BTRCALL(BGETEQUAL, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), nRecord, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error getting current record." & BtrieveErrorCode(nStatus)
End If

Exit Sub
error:
Call HandleError
On Error Resume Next
Unload frmProgressBar
DoEvents
frmMain.Enabled = True

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim nStatus As Integer

Me.Top = ReadINI("Windows", "MessageTop")
Me.Left = ReadINI("Windows", "MessageLeft")

nStatus = BTRCALL(BGETFIRST, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadMessage, BGETFIRST, Message, Error: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    bLoaded = True
    DispMessageInfo Messagedatabuf.buf
End If

Me.Show
Me.SetFocus
txtGoto.SetFocus

End Sub
Private Sub cmdDiscard_Click()
Dim nStatus As Integer, nGoto As Long
nGoto = txtNumber.Text
nStatus = BTRCALL(BGETEQUAL, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), nGoto, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdGoto_Click(), BGETEQUAL, Message, Error: " & BtrieveErrorCode(nStatus)
Else
    DispMessageInfo Messagedatabuf.buf
End If
End Sub

Private Sub cmdSave_Click()
Dim nStatus As Integer, nGoto As Long

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
SaveMessage

nGoto = Val(txtNumber.Text)

nStatus = BTRCALL(BGETEQUAL, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), nGoto, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdGoto_Click(), BGETEQUAL, Message, Error: " & BtrieveErrorCode(nStatus)
Else
    DispMessageInfo Messagedatabuf.buf
End If
End Sub



Private Sub DispMessageInfo(row() As Byte)
On Error GoTo error:
bLoaded = True

RowToStruct row, MessageFldMap, Messagerec, LenB(Messagerec)
    
Me.Caption = "Message Editor -- " & Messagerec.Number

txtNumber.Text = Messagerec.Number
txtMessageLine1.Text = Messagerec.MessageLine1
txtMessageLine2.Text = Messagerec.MessageLine2
txtMessageLine3.Text = Messagerec.MessageLine3

Exit Sub
error:
Call HandleError
MsgBox "Warning, record was not completely displayed." & vbCrLf _
    & "Previous records stats may still be in memory.  Select 'Disable DB Writing'" & vbCrLf _
    & "from the file menu and then reload the editor.", vbExclamation
End Sub

Private Sub SaveMessage()
On Error GoTo error:
Dim nStatus As Integer, nGoto As Long

nGoto = Val(txtNumber.Text)
nStatus = BTRCALL(BGETEQUAL, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), nGoto, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Save Error, BGETEQUAL, Message, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

MessageRowToStruct Messagedatabuf.buf

'DoEvents
Messagerec.MessageLine1 = RTrim(txtMessageLine1.Text) & Chr(0)
Messagerec.MessageLine2 = RTrim(txtMessageLine2.Text) & Chr(0)
Messagerec.MessageLine3 = RTrim(txtMessageLine3.Text) & Chr(0)

UpdateMessageRecord

Exit Sub
error:
Call HandleError

End Sub
Private Sub UpdateMessageRecord()
Dim nStatus As Integer

nStatus = UpdateMessage
If Not nStatus = 0 Then
    MsgBox "Update Message Error: " & BtrieveErrorCode(nStatus)
Else
    DispMessageInfo Messagedatabuf.buf
End If
End Sub
Public Sub GotoMSG(ByVal nMSG As Long)
On Error GoTo error:
Dim nStatus As Integer
If bLoaded Then SaveMessage

Me.Show
Me.SetFocus

nStatus = BTRCALL(BGETEQUAL, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), nMSG, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdGoto_Click(), BGETEQUAL, Message, Error: " & BtrieveErrorCode(nStatus)
Else
    DispMessageInfo Messagedatabuf.buf
End If

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdGoto_Click()
On Error GoTo error:
Dim nStatus As Integer, nGoto As Long
SaveMessage
nGoto = Val(txtGoto.Text)

nStatus = BTRCALL(BGETEQUAL, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), nGoto, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdGoto_Click(), BGETEQUAL, Message, Error: " & BtrieveErrorCode(nStatus)
Else
    DispMessageInfo Messagedatabuf.buf
End If

Exit Sub
error:
Call HandleError
End Sub


Private Sub cmdFirst_Click()
Dim nStatus As Integer
If bLoaded = True Then SaveMessage
nStatus = BTRCALL(BGETFIRST, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdFirst_Click, BGETFIRST, Message, Error: " & BtrieveErrorCode(nStatus)
Else
    DispMessageInfo Messagedatabuf.buf
End If
End Sub

Private Sub cmdNext_Click()
Dim nStatus As Integer
If bLoaded = True Then SaveMessage

nStatus = BTRCALL(BGETNEXT, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdNext_Click, BGETNEXT, Message, Error: " & BtrieveErrorCode(nStatus)
Else
    DispMessageInfo Messagedatabuf.buf
End If
End Sub

Private Sub cmdPrevious_Click()
Dim nStatus As Integer
If bLoaded = True Then SaveMessage

nStatus = BTRCALL(BGETPREVIOUS, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdPrevious_Click, BGETPREVIOUS, Message, Error: " & BtrieveErrorCode(nStatus)
Else
    DispMessageInfo Messagedatabuf.buf
End If
End Sub

Private Sub cmdLast_Click()
Dim nStatus As Integer
If bLoaded = True Then SaveMessage

nStatus = BTRCALL(BGETLAST, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdLast_Click, BGETLAST, Message, Error: " & BtrieveErrorCode(nStatus)
Else
    DispMessageInfo Messagedatabuf.buf
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
        If bLoaded = True Then SaveMessage
        If Me.WindowState = vbMinimized Then Exit Sub
        Call WriteINI("Windows", "MessageTop", Me.Top)
        Call WriteINI("Windows", "MessageLeft", Me.Left)
End Sub

Private Sub cmdDelete_Click()
Dim nStatus As Integer
Dim nDelete As Integer

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
SaveMessage

nDelete = MsgBox("Delete this record from database?", vbYesNo, "Delete Record?")
If nDelete <> 6 Then Exit Sub

    nStatus = BTRCALL(BDELETE, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "cmdDelete, BDELETE, Error: " & BtrieveErrorCode(nStatus)
    Else
        Form_Load
    End If

End Sub

Private Sub cmdInsert_Click()
On Error GoTo error:
Dim nStatus As Integer
Dim nNewMessageNumber As Variant

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
If bLoaded = True Then SaveMessage

nNewMessageNumber = InputBox("New Message Number:" & vbCrLf & vbCrLf & "Enter 0 for the next highest number.", "Insert", Val(txtNumber.Text) + 1)
If nNewMessageNumber = "" Then Exit Sub

    Messagerec.Number = Val(nNewMessageNumber)
    
    MessageStructToRow Messagedatabuf.buf
    
    nStatus = BTRCALL(BINSERT, MessagePosBlock, Messagedatabuf, Len(Messagedatabuf), ByVal MessageKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "cmdInsert, BINSERT, Error: " & BtrieveErrorCode(nStatus)
    Else
        DispMessageInfo Messagedatabuf.buf
    End If

Exit Sub
error:
Call HandleError
End Sub


Private Sub txtGoto_GotFocus()
Call SelectAll(txtGoto)

End Sub

Private Sub txtNumber_GotFocus()
Call SelectAll(txtNumber)

End Sub
