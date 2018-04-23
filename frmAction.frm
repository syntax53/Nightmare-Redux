VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Begin VB.Form frmAction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Action Editor"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "frmAction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   6285
   Begin VB.TextBox txtOffset 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   6540
      Width           =   975
   End
   Begin VB.CheckBox chkAutoSave 
      Caption         =   "Auto-Save"
      Height          =   195
      Left            =   3300
      TabIndex        =   29
      Top             =   60
      Value           =   1  'Checked
      Width           =   1155
   End
   Begin VB.CommandButton cmdDiscard 
      Caption         =   "Dis&card"
      Height          =   315
      Left            =   5340
      TabIndex        =   5
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   4500
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox txtGoto 
      Height          =   285
      Left            =   60
      MaxLength       =   28
      TabIndex        =   0
      Top             =   60
      Width           =   1515
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   0
      Width           =   795
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtFloorItemToRoom 
      Height          =   285
      Left            =   1680
      MaxLength       =   74
      TabIndex        =   27
      Top             =   6000
      Width           =   4455
   End
   Begin VB.TextBox txtFloorItemToUser 
      Height          =   285
      Left            =   1680
      MaxLength       =   74
      TabIndex        =   25
      Top             =   5460
      Width           =   4455
   End
   Begin VB.TextBox txtInventoryToRoom 
      Height          =   285
      Left            =   1680
      MaxLength       =   74
      TabIndex        =   23
      Top             =   4920
      Width           =   4455
   End
   Begin VB.TextBox txtInventoryToUser 
      Height          =   285
      Left            =   1680
      MaxLength       =   74
      TabIndex        =   21
      Top             =   4380
      Width           =   4455
   End
   Begin VB.TextBox txtMonsterToRoom 
      Height          =   285
      Left            =   1680
      MaxLength       =   74
      TabIndex        =   19
      Top             =   3840
      Width           =   4455
   End
   Begin VB.TextBox txtMonsterToUser 
      Height          =   285
      Left            =   1680
      MaxLength       =   74
      TabIndex        =   17
      Top             =   3300
      Width           =   4455
   End
   Begin VB.TextBox txtUserToRoom 
      Height          =   285
      Left            =   1680
      MaxLength       =   74
      TabIndex        =   15
      Top             =   2760
      Width           =   4455
   End
   Begin VB.TextBox txtUserToOtherUser 
      Height          =   285
      Left            =   1680
      MaxLength       =   74
      TabIndex        =   13
      Top             =   2220
      Width           =   4455
   End
   Begin VB.TextBox txtUserToUser 
      Height          =   285
      Left            =   1680
      MaxLength       =   74
      TabIndex        =   11
      Top             =   1680
      Width           =   4455
   End
   Begin VB.TextBox txtSingleToRoom 
      Height          =   285
      Left            =   1680
      MaxLength       =   74
      TabIndex        =   9
      Top             =   1140
      Width           =   4455
   End
   Begin VB.TextBox txtSingleToUser 
      Height          =   285
      Left            =   1680
      MaxLength       =   74
      TabIndex        =   7
      Top             =   600
      Width           =   4455
   End
   Begin MSComctlLib.ListView lvDatabase 
      Height          =   5895
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtActionCommand 
      Height          =   315
      Left            =   1200
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Offset (Debug)"
      Height          =   255
      Left            =   1680
      TabIndex        =   31
      Top             =   6300
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Floor Item to Room   (EX: %s peed on the %s!)"
      Height          =   255
      Left            =   1680
      TabIndex        =   26
      Top             =   5760
      Width           =   4455
   End
   Begin VB.Label Label10 
      Caption         =   "Floor Item to User   (EX: You pee on the %s!)"
      Height          =   255
      Left            =   1680
      TabIndex        =   24
      Top             =   5220
      Width           =   4455
   End
   Begin VB.Label Label9 
      Caption         =   "Inventory to Room   (EX: %s peed on his %s!)"
      Height          =   255
      Left            =   1680
      TabIndex        =   22
      Top             =   4680
      Width           =   4455
   End
   Begin VB.Label Label8 
      Caption         =   "Inventory to User   (EX: You pee on your %s!)"
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   4140
      Width           =   4455
   End
   Begin VB.Label Label7 
      Caption         =   "Monster to Room   (EX: %s peed on the %s!)"
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   3600
      Width           =   4455
   End
   Begin VB.Label Label6 
      Caption         =   "Monster to User   (EX: You pee on the %s!)"
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   3060
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "User to Room   (EX: %s peed on %s!)"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "User to Other User   (EX: %s peed on you!)"
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   1980
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "User to User   (EX: You pee on %s!)"
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Single to Room   (EX: %s peed on everyone!)"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   900
      Width           =   4455
   End
   Begin VB.Label labelxx 
      Caption         =   "Single to User   (EX: You pee on everyone!)"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frmAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim bLoaded As Boolean

Private Sub Form_Load()
On Error Resume Next

Me.Top = ReadINI("Windows", "ActionTop")
Me.Left = ReadINI("Windows", "ActionLeft")

bLoaded = False
lvDatabase.ListItems.clear
Call LoadActions

Me.Show
Me.SetFocus
txtGoto.SetFocus
If ReadINI("Windows", "ActionMaxed") = "1" Then Me.WindowState = vbMaximized
End Sub

Private Sub LoadActions()
On Error GoTo error:
Dim nStatus As Integer

lvDatabase.ColumnHeaders.clear
lvDatabase.ColumnHeaders.add 1, "Action", "Action", 1000, lvwColumnLeft

nStatus = BTRCALL(BGETFIRST, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "LoadAction, BGETFIRST, Action, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

Do While nStatus = 0
    ActionRowToStruct ActionDatabuf.buf

    Call AddActionToLV(Actionrec.Name)

    nStatus = BTRCALL(BGETNEXT, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0)
Loop
If Not nStatus = 0 And Not nStatus = 9 Then
    MsgBox "LoadActions, Error: " & BtrieveErrorCode(nStatus)
End If

lvDatabase.refresh
SortListView lvDatabase, 1, ldtString, True
If lvDatabase.ListItems.Count >= 1 Then Call lvDatabase_ItemClick(lvDatabase.ListItems(1))
bLoaded = True

Exit Sub
error:
Call HandleError
End Sub

Private Sub AddActionToLV(ByVal sAction As String)
Dim nStatus As Integer, oLI As ListItem, sGoto As String * 30
On Error GoTo error:

If Not sAction = Actionrec.Name Then
    sGoto = sAction & Chr(0)
    nStatus = BTRCALL(BGETEQUAL, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal sGoto, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then MsgBox "Error getting action '" & sAction & "': " & BtrieveErrorCode(nStatus)
    bLoaded = False
    Exit Sub
End If

Set oLI = lvDatabase.ListItems.add()
oLI.Text = Actionrec.Name

Set oLI = Nothing
Exit Sub
error:
Call HandleError
Set oLI = Nothing
End Sub

Private Sub cmdDiscard_Click()
Dim nStatus As Integer, sGoto As String * 30

sGoto = txtActionCommand & Chr(0)
nStatus = BTRCALL(BGETEQUAL, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal sGoto, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdGoto_Click(), BGETEQUAL, Action, Error: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    DispActionInfo ActionDatabuf.buf
    bLoaded = True
End If
End Sub

Private Sub cmdSave_Click()
Dim nStatus As Integer, sGoto As String * 30

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

Call SaveAction
sGoto = txtActionCommand & Chr(0)
nStatus = BTRCALL(BGETEQUAL, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal sGoto, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "cmdGoto_Click(), BGETEQUAL, Action, Error: " & BtrieveErrorCode(nStatus)
Else
    DispActionInfo ActionDatabuf.buf
End If
End Sub



Private Sub DispActionInfo(row() As Byte)
On Error GoTo error:
bLoaded = True

RowToStruct row, ActionFldMap, Actionrec, LenB(Actionrec)

Me.Caption = "Action Editor -- " & ClipNull(Actionrec.Name)

txtActionCommand.Text = Actionrec.Name
txtSingleToUser.Text = Actionrec.SingleToUser
txtSingleToRoom.Text = Actionrec.SingleToRoom
txtUserToUser.Text = Actionrec.UserToUser
txtUserToOtherUser.Text = Actionrec.UserToOtherUser
txtUserToRoom.Text = Actionrec.UserToRoom
txtMonsterToUser.Text = Actionrec.MonsterToUser
txtMonsterToRoom.Text = Actionrec.MonsterToRoom
txtInventoryToUser.Text = Actionrec.InventoryToUser
txtInventoryToRoom.Text = Actionrec.InventoryToRoom
txtFloorItemToUser.Text = Actionrec.FloorItemToUser
txtFloorItemToRoom.Text = Actionrec.FloorItemToRoom

UpdateOffsetValue

Exit Sub
error:
Call HandleError
MsgBox "Warning, record was not completely displayed." & vbCrLf _
    & "Previous records stats may still be in memory.  Select 'Disable DB Writing'" & vbCrLf _
    & "from the file menu and then reload the editor.", vbExclamation
End Sub


Private Sub SaveAction()
On Error GoTo error:
Dim nStatus As Integer, sGoto As String * 30

sGoto = txtActionCommand.Text & Chr(0)

nStatus = BTRCALL(BGETEQUAL, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal sGoto, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Save Error, BGETEQUAL, Action, Error: " & BtrieveErrorCode(nStatus)
    Exit Sub
End If

ActionRowToStruct ActionDatabuf.buf

'DoEvents
Actionrec.SingleToUser = RTrim(txtSingleToUser.Text) & Chr(0)
Actionrec.SingleToRoom = RTrim(txtSingleToRoom.Text) & Chr(0)
Actionrec.UserToUser = RTrim(txtUserToUser.Text) & Chr(0)
Actionrec.UserToOtherUser = RTrim(txtUserToOtherUser.Text) & Chr(0)
Actionrec.UserToRoom = RTrim(txtUserToRoom.Text) & Chr(0)
Actionrec.MonsterToUser = RTrim(txtMonsterToUser.Text) & Chr(0)
Actionrec.MonsterToRoom = RTrim(txtMonsterToRoom.Text) & Chr(0)
Actionrec.InventoryToUser = RTrim(txtInventoryToUser.Text) & Chr(0)
Actionrec.InventoryToRoom = RTrim(txtInventoryToRoom.Text) & Chr(0)
Actionrec.FloorItemToUser = RTrim(txtFloorItemToUser.Text) & Chr(0)
Actionrec.FloorItemToRoom = RTrim(txtFloorItemToRoom.Text) & Chr(0)

UpdateOffsetValue
Actionrec.Offset = txtOffset.Text

UpdateActionRecord

Exit Sub
error:
Call HandleError
End Sub

Private Sub UpdateOffsetValue()

Dim tmpOffset As Byte

tmpOffset = 0

If Len(txtSingleToUser.Text) > 0 Or Len(txtSingleToUser.Text) > 0 Then
    tmpOffset = tmpOffset + 1
End If
If Len(txtUserToUser.Text) > 0 Or Len(txtUserToOtherUser.Text) > 0 Or Len(txtUserToRoom.Text) > 0 Then
    tmpOffset = tmpOffset + 2
End If
If Len(txtMonsterToUser.Text) > 0 Or Len(txtMonsterToRoom.Text) > 0 Then
    tmpOffset = tmpOffset + 4
End If
If Len(txtInventoryToUser.Text) > 0 Or Len(txtInventoryToRoom.Text) > 0 Then
    tmpOffset = tmpOffset + 8
End If
If Len(txtFloorItemToUser.Text) > 0 Or Len(txtFloorItemToRoom.Text) > 0 Then
    tmpOffset = tmpOffset + 16
End If

txtOffset.Text = tmpOffset

Exit Sub
error:
Call HandleError
End Sub

Private Sub UpdateActionRecord()
Dim nStatus As Integer

nStatus = UpdateAction
If Not nStatus = 0 Then
    MsgBox "Update ActiOn Error: " & BtrieveErrorCode(nStatus)
Else
    DispActionInfo ActionDatabuf.buf
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bLoaded = True Then SaveAction
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState = vbMaximized Then
        Call WriteINI("Windows", "ActionMaxed", 1)
    Else
        Call WriteINI("Windows", "ActionMaxed", 0)
        Call WriteINI("Windows", "ActionTop", Me.Top)
        Call WriteINI("Windows", "ActionLeft", Me.Left)
    End If
End Sub

Private Sub cmdDelete_Click()
On Error GoTo error:
Dim nStatus As Integer
Dim nDelete As Integer, temp As Long, strTemp As String * 30

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

nDelete = MsgBox("Delete this record from database?", vbYesNo, "Delete Record?")
If Not nDelete = vbYes Then Exit Sub

If bLoaded Then Call SaveAction
    
temp = lvDatabase.SelectedItem.Index

strTemp = txtActionCommand.Text & Chr(0)
nStatus = BTRCALL(BGETEQUAL, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal strTemp, KEY_BUF_LEN, 0)
If nStatus = 0 Then
    nStatus = BTRCALL(BDELETE, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "cmdDelete, BDELETE, Error: " & BtrieveErrorCode(nStatus)
    Else
        lvDatabase.ListItems.Remove temp
        bLoaded = False
        If lvDatabase.ListItems.Count >= 1 Then
            If temp > 1 Then temp = temp - 1 Else temp = 1
            Set lvDatabase.SelectedItem = lvDatabase.ListItems(temp)
            lvDatabase.SelectedItem.EnsureVisible
            Call lvDatabase_ItemClick(lvDatabase.SelectedItem)
        Else
            Call Form_Unload(1)
            Call Form_Load
        End If
    End If
Else
    MsgBox "Couldn't get record, Error: " & BtrieveErrorCode(nStatus)
End If

Exit Sub
error:
Call HandleError
End Sub

Private Sub cmdInsert_Click()
On Error GoTo error:
Dim nStatus As Integer, sGoto As String, oLI As ListItem
Dim nNewActionName As String
If bLoaded = True Then SaveAction

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
nNewActionName = InputBox("New Action Name:", "Insert", "")
If nNewActionName = "" Then Exit Sub
    
    Actionrec.Name = nNewActionName & Chr(0)
    
    ActionStructToRow ActionDatabuf.buf
    
    nStatus = BTRCALL(BINSERT, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal ActionKeyBuffer, KEY_BUF_LEN, 0)
    If Not nStatus = 0 Then
        MsgBox "cmdInsert, BINSERT, Error: " & BtrieveErrorCode(nStatus)
    Else
        ActionRowToStruct ActionDatabuf.buf
    
        Call AddActionToLV(Actionrec.Name)
        
        DispActionInfo ActionDatabuf.buf
        
        SortListView lvDatabase, 1, ldtString, True
        
        Set oLI = lvDatabase.FindItem(Actionrec.Name, lvwText, , 0)
        If Not oLI Is Nothing Then
            Set lvDatabase.SelectedItem = oLI
            lvDatabase.SelectedItem.EnsureVisible
            Call lvDatabase_ItemClick(lvDatabase.SelectedItem)
            Set oLI = Nothing
        Else
            Set lvDatabase.SelectedItem = lvDatabase.ListItems(lvDatabase.ListItems.Count)
            lvDatabase.SelectedItem.EnsureVisible
            Call lvDatabase_ItemClick(lvDatabase.SelectedItem)
        End If
    
'        '''''
'
'        sGoto = nNewActionName & Chr(0)
'        nStatus = BTRCALL(BGETEQUAL, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal sGoto, KEY_BUF_LEN, 0)
'        If Not nStatus = 0 Then
'            MsgBox "cmdGoto_Click(), BGETEQUAL, Action, Error: " & BtrieveErrorCode(nStatus)
'        Else
'            DispActionInfo ActionDatabuf.buf
'        End If
    End If
    
Exit Sub
error:
Call HandleError
Set oLI = Nothing
End Sub

Private Sub lvDatabase_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
SortListView lvDatabase, ColumnHeader.Index, ldtString, lvDatabase.SortOrder
End Sub

Private Sub lvDatabase_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim strTemp As String * 30, nStatus As Integer

If bLoaded = True And chkAutoSave.Value = 1 Then Call SaveAction

strTemp = Item.Text & Chr(0)

nStatus = BTRCALL(BGETEQUAL, ActionPosBlock, ActionDatabuf, Len(ActionDatabuf), ByVal strTemp, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Error on BGETEQUAL: " & BtrieveErrorCode(nStatus)
    bLoaded = False
Else
    DispActionInfo ActionDatabuf.buf
    bLoaded = True
End If
End Sub

Private Sub txtFloorItemToRoom_GotFocus()
Call SelectAll(txtFloorItemToRoom)

End Sub

Private Sub txtFloorItemToUser_GotFocus()
Call SelectAll(txtFloorItemToUser)

End Sub

Private Sub txtGoto_GotFocus()
Call SelectAll(txtGoto)

End Sub

Private Sub txtGoto_KeyUp(KeyCode As Integer, Shift As Integer)
Dim x As Long, SearchStart As Long

If txtGoto.Text = "" Then Exit Sub
If lvDatabase.ListItems.Count < 1 Then Exit Sub

SearchStart = 1

If KeyCode = vbKeyUp Then Exit Sub
If KeyCode = vbKeyDown Then lvDatabase.SetFocus
If KeyCode = vbKeyLeft Then Exit Sub
If KeyCode = vbKeyRight Then SearchStart = lvDatabase.SelectedItem.Index + 1
If KeyCode = vbKeyControl Then Exit Sub 'control
If KeyCode = 18 Then Exit Sub 'alt
If KeyCode = vbKeyTab Then Exit Sub 'tab
If KeyCode = vbKeyShift Then Exit Sub

For x = SearchStart To lvDatabase.ListItems.Count
    If Not InStr(1, LCase(lvDatabase.ListItems(x)), LCase(txtGoto.Text)) = 0 Then
        Set lvDatabase.SelectedItem = lvDatabase.ListItems(x)
        lvDatabase.SelectedItem.EnsureVisible
        Call lvDatabase_ItemClick(lvDatabase.SelectedItem)
        Exit For
    End If
Next x

End Sub

Private Sub txtInventoryToRoom_GotFocus()
Call SelectAll(txtInventoryToRoom)

End Sub

Private Sub txtInventoryToUser_GotFocus()
Call SelectAll(txtInventoryToUser)
End Sub

Private Sub txtMonsterToRoom_GotFocus()
Call SelectAll(txtMonsterToRoom)

End Sub

Private Sub txtMonsterToUser_GotFocus()
Call SelectAll(txtMonsterToUser)

End Sub

Private Sub txtSingleToRoom_GotFocus()
Call SelectAll(txtSingleToRoom)

End Sub

Private Sub txtSingleToUser_GotFocus()
Call SelectAll(txtSingleToUser)
End Sub

Private Sub txtUserToOtherUser_GotFocus()
Call SelectAll(txtUserToOtherUser)

End Sub

Private Sub txtUserToRoom_GotFocus()
Call SelectAll(txtUserToRoom)

End Sub

Private Sub txtUserToUser_GotFocus()
Call SelectAll(txtUserToUser)

End Sub
