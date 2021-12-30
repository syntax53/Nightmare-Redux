VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMonsterNPC_List 
   Caption         =   "Monster NPC/Room List"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7395
   Icon            =   "frmMonsterNPC_List.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   7395
   Begin exlimiter.EL EL1 
      Left            =   6480
      Top             =   4500
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin VB.CommandButton cmdBuild 
      Caption         =   "&Build New List"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1395
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save List"
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   6480
      TabIndex        =   2
      Top             =   60
      Width           =   855
   End
   Begin MSComctlLib.ListView lvNPC_List 
      Height          =   4935
      Left            =   60
      TabIndex        =   3
      Top             =   420
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMonsterNPC_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim bCancel As Boolean


Private Sub Form_Load()
On Error Resume Next

With EL1
    .FormInQuestion = Me
    .MINWIDTH = 500 + (TITLEBAR_OFFSET / 10)
    .MINHEIGHT = 200
    .EnableLimiter = True
End With

Me.Width = 7515
Me.Height = 5925

Call AddColumnHeaders

End Sub
Private Sub cmdBuild_Click()
On Error GoTo error:
Dim nStatus As Integer
bCancel = False
lvNPC_List.ListItems.clear

nStatus = BTRCALL(BGETLAST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Could not get last Room record, Error: " & BtrieveErrorCode(nStatus), vbOKOnly, "Creating Monster NPC/Room List"
    Exit Sub
End If
Call RoomRowToStruct(Roomdatabuf.buf)

frmProgressBar.sCaption = "Building Monster NPC/Room List"
frmProgressBar.lblCaption.Caption = "Scanning Rooms ..."
frmProgressBar.cmdCancel.Enabled = True
frmProgressBar.lblPanel(0).Caption = "w" & strDatCallLetters & "mp002.dat"
frmProgressBar.lblPanel(1).Caption = ""
Call frmProgressBar.SetRange(Roomrec.RoomNumber)
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

Me.WindowState = vbMinimized
Me.Hide
'LockWindowUpdate Me.hWnd
Call CreateList

If lvNPC_List.ListItems.Count > 0 Then
    SortListView lvNPC_List, 2, ldtString, True
End If

Unload frmProgressBar
frmMain.Enabled = True
'LockWindowUpdate 0&
Me.WindowState = vbNormal
Me.Show
Me.SetFocus

Exit Sub
error:
Call HandleError("cmdBuild_Click")
Unload frmProgressBar
frmMain.Enabled = True
'LockWindowUpdate 0&
Me.WindowState = vbNormal
Me.Show
End Sub

Private Sub CreateList()
Dim nStatus As Integer, oLI As ListItem ', x As Integer
Dim nNPC As Long, nMapNum As Long, nRoomNum As Long

nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Could not get first Room record, Error: " & BtrieveErrorCode(nStatus), vbOKOnly, "Creating Monster NPC/Room List"
    Exit Sub
End If

'For x = 1 To 500
'    GetNextExDataBuf.buf(x) = &H0
'Next x

'GetNextExRec.HeaderDataBufLength = 35
Call SetBinaryValue(35, GetNextExDataBuf.buf(), 1, 2)
'GetNextExRec.HeaderBeginCode = "UC"
Call SetBinaryString("UC", GetNextExDataBuf.buf(), 3, 2)
'GetNextExRec.HeaderMaxReject = 20000
Call SetBinaryValue(65535, GetNextExDataBuf.buf(), 5, 2)
'GetNextExRec.HeaderNumFilterTerms = 1
Call SetBinaryValue(1, GetNextExDataBuf.buf(), 7, 2)

'GetNextExRec.LogicFieldDataType = FLD_INTEGER
Call SetBinaryValue(FLD_INTEGER, GetNextExDataBuf.buf(), 9, 1)
'GetNextExRec.LogicFieldLength = 4
Call SetBinaryValue(4, GetNextExDataBuf.buf(), 10, 2)
'GetNextExRec.LogicFieldOffset = 1480
Call SetBinaryValue(1480, GetNextExDataBuf.buf(), 12, 2)
'GetNextExRec.LogicComparisonCode = 2 '>
Call SetBinaryValue(2, GetNextExDataBuf.buf(), 14, 1)
'GetNextExRec.LogicAndOr = 0
Call SetBinaryValue(0, GetNextExDataBuf.buf(), 15, 1)
'GetNextExRec.Logic2ndOffsetOrConstant = 0
Call SetBinaryValue(0, GetNextExDataBuf.buf(), 16, 4)

'GetNextExRec.HeaderNumRecordsReturned = 1
Call SetBinaryValue(1, GetNextExDataBuf.buf(), 20, 2)
'GetNextExRec.HeaderNumFieldsExtracted = 3
Call SetBinaryValue(3, GetNextExDataBuf.buf(), 22, 2)

'GetNextExRec.Field1Length = 4
Call SetBinaryValue(4, GetNextExDataBuf.buf(), 24, 2)
'GetNextExRec.Field1Offset = 0 'mapnum
Call SetBinaryValue(0, GetNextExDataBuf.buf(), 26, 2)
'GetNextExRec.Field2Length = 4
Call SetBinaryValue(4, GetNextExDataBuf.buf(), 28, 2)
'GetNextExRec.Field2Offset = 4 'roomnum
Call SetBinaryValue(4, GetNextExDataBuf.buf(), 30, 2)
'GetNextExRec.Field3Length = 4
Call SetBinaryValue(4, GetNextExDataBuf.buf(), 32, 2)
'GetNextExRec.Field3Offset = 1480 'npc
Call SetBinaryValue(1480, GetNextExDataBuf.buf(), 34, 2)

nStatus = BTRCALL(BGETNEXTEXTENDED, RoomPosBlock, GetNextExDataBuf, Len(GetNextExDataBuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Could not get first Room record, Error: " & BtrieveErrorCode(nStatus), vbOKOnly, "Creating Monster NPC/Room List"
    Exit Sub
End If

Do While nStatus = 0 And bCancel = False
    frmProgressBar.lblPanel(1).Caption = nRoomNum
    frmProgressBar.ProgressBar.Value = nRoomNum
    Call frmProgressBar.IncreaseProgress
    
    nMapNum = GetBinaryValue(GetNextExDataBuf.buf(), 9, 4)
    nRoomNum = GetBinaryValue(GetNextExDataBuf.buf(), 13, 4)
    nNPC = GetBinaryValue(GetNextExDataBuf.buf(), 17, 4)

    If nNPC > 0 Then
        Set oLI = lvNPC_List.ListItems.add()
        oLI.Text = nMapNum & "/" & nRoomNum
        oLI.ListSubItems.add (1), "Monster", GetMonsterName(nNPC) & " (" & nNPC & ")"
        Set oLI = Nothing
    End If

    'header
    Call SetBinaryValue(35, GetNextExDataBuf.buf(), 1, 2)
    Call SetBinaryString("EG", GetNextExDataBuf.buf(), 3, 2)
    Call SetBinaryValue(65535, GetNextExDataBuf.buf(), 5, 2)
    Call SetBinaryValue(1, GetNextExDataBuf.buf(), 7, 2)
    'filter
    Call SetBinaryValue(FLD_INTEGER, GetNextExDataBuf.buf(), 9, 1)
    Call SetBinaryValue(4, GetNextExDataBuf.buf(), 10, 2)
    Call SetBinaryValue(1480, GetNextExDataBuf.buf(), 12, 2)
    Call SetBinaryValue(2, GetNextExDataBuf.buf(), 14, 1)
    Call SetBinaryValue(0, GetNextExDataBuf.buf(), 15, 1)
    Call SetBinaryValue(0, GetNextExDataBuf.buf(), 16, 4)
    'record info
    Call SetBinaryValue(1, GetNextExDataBuf.buf(), 20, 2)
    Call SetBinaryValue(3, GetNextExDataBuf.buf(), 22, 2)
    'field info
    Call SetBinaryValue(4, GetNextExDataBuf.buf(), 24, 2)
    Call SetBinaryValue(0, GetNextExDataBuf.buf(), 26, 2)
    Call SetBinaryValue(4, GetNextExDataBuf.buf(), 28, 2)
    Call SetBinaryValue(4, GetNextExDataBuf.buf(), 30, 2)
    Call SetBinaryValue(4, GetNextExDataBuf.buf(), 32, 2)
    Call SetBinaryValue(1480, GetNextExDataBuf.buf(), 34, 2)
        
    nStatus = BTRCALL(BGETNEXTEXTENDED, RoomPosBlock, GetNextExDataBuf, Len(GetNextExDataBuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

If Not nStatus = 9 And Not nStatus = 0 Then
    MsgBox "Abnormal Exit: " & BtrieveErrorCode(nStatus), vbOKOnly + vbExclamation, "Creating Monster NPC/Room List"
End If

Set oLI = Nothing
End Sub
Public Sub ToggleStopBuild()
    bCancel = True
End Sub

Private Sub AddColumnHeaders()

lvNPC_List.ColumnHeaders.clear
lvNPC_List.ColumnHeaders.add 1, "Room", "Map/Room", 1200, lvwColumnLeft
lvNPC_List.ColumnHeaders.add 2, "Monster", "Monster", 10000, lvwColumnLeft
'lvNPC_List.ColumnHeaders.add 3, "Name", "Name", 2400, lvwColumnCenter
'lvNPC_List.ColumnHeaders.add 4, "MapRoom", "Map/Room Monster Regens In", 6000, lvwColumnLeft

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo error:
Dim oLI As ListItem, str As String, x As Integer, nMaxGrpLen As Integer
Dim fso As FileSystemObject, nYesNo As Integer, sFile As TextStream

CommonDialog1.Filter = "TXT Files (*.txt)|*.txt"
CommonDialog1.DialogTitle = "Enter New File Name"
CommonDialog1.FileName = "NMR-MonsterNPC_List.txt"

On Error GoTo canceled:
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then GoTo canceled:

On Error GoTo error:

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(CommonDialog1.FileName) Then
    nYesNo = MsgBox("File Exists, Overwrite?", vbYesNo, "Overwrite?")
    If nYesNo = vbYes Then
        fso.DeleteFile (CommonDialog1.FileName), True
    Else
        GoTo canceled:
    End If
End If

Set sFile = fso.OpenTextFile(CommonDialog1.FileName, ForWriting, True)
sFile.WriteLine ("Monster NPC List -- " & Date & " @ " & Time)
sFile.WriteBlankLines (1)

nMaxGrpLen = 10
For Each oLI In lvNPC_List.ListItems
    
    If Len(oLI.Text) > nMaxGrpLen Then nMaxGrpLen = Len(oLI.Text)
    
    str = oLI.Text & " " & String(nMaxGrpLen - Len(oLI.Text), ".") & " "
    str = str & oLI.SubItems(1)
    
    sFile.WriteLine (str)
Next

sFile.Close

canceled:
Set fso = Nothing
Set sFile = Nothing
Set oLI = Nothing
Exit Sub
error:
Call HandleError
Set fso = Nothing
Set sFile = Nothing
Set oLI = Nothing
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub

cmdClose.Left = Me.Width - 1000
lvNPC_List.Width = Me.Width - 230
lvNPC_List.Height = Me.Height - TITLEBAR_OFFSET - 860

If Not lvNPC_List.ColumnHeaders.Count = 0 Then
    lvNPC_List.ColumnHeaders(2).Width = lvNPC_List.Width - 1700
End If

End Sub

Private Sub lvNPC_List_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
SortListView lvNPC_List, ColumnHeader.Index, ldtString, lvNPC_List.SortOrder
'If ColumnHeader.Index = 1 Then
'    SortListView lvNPC_List, ColumnHeader.Index, ldtNumber, lvNPC_List.SortOrder
'Else
'    SortListView lvNPC_List, ColumnHeader.Index, ldtString, lvNPC_List.SortOrder
'End If
End Sub


Public Sub CopyLine()

If lvNPC_List.SelectedItem Is Nothing Then Exit Sub

Clipboard.clear
Clipboard.SetText lvNPC_List.SelectedItem.Text & " -- " & lvNPC_List.SelectedItem.SubItems(1)

End Sub

Private Sub lvNPC_List_DblClick()
Dim typExits As RoomExitType
On Error GoTo error:

'Call frmMonster.GotoMonster(Val(lvNPC_List.SelectedItem.Tag))
typExits = ExtractMapRoom(lvNPC_List.SelectedItem.Text)
If typExits.Map > 0 And typExits.Room > 0 Then
    Call frmRoom.GotoRoom(typExits.Map, typExits.Room)
End If


Exit Sub
error:
Call HandleError("lvNPC_List_DblClick")

End Sub

Private Sub lvNPC_List_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu frmMain.mnuMonsterNPCListRightClick
End If
End Sub
