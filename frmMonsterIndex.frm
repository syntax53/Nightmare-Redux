VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20D5284F-7B23-4F0A-B8B1-6C9D18B64F1C}#1.0#0"; "exlimiter.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMonsterIndex 
   Caption         =   "Monster Group / Index List"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7395
   Icon            =   "frmMonsterIndex.frx":0000
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
   Begin MSComctlLib.ListView lvMonsterIndex 
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
      Left            =   6300
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "(You can have this list automaticly created under settings for use in the map views.)"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   60
      Width           =   3315
   End
End
Attribute VB_Name = "frmMonsterIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

'Dim LimitedItem() As Boolean
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

If Not UBound(MGIL(), 1) = 1 Then Call JustCreateList

End Sub
Private Sub cmdBuild_Click()
On Error GoTo error:

bCancel = False
Me.WindowState = vbMinimized
Me.Hide

frmProgressBar.sCaption = "Building Monster Group/Index List"
frmProgressBar.lblCaption.Caption = "Scanning Monsters ..."
frmProgressBar.cmdCancel.Enabled = True
Call frmProgressBar.SetRange(CalcTotalRecords)
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

lvMonsterIndex.ListItems.clear
Erase MGIL()
ReDim MGIL(39, 9999)
Call ScanMonsters
If bCancel Then GoTo ReEnable:

frmProgressBar.lblCaption.Caption = "Creating List ..."
Call frmProgressBar.SetRange(39)
frmProgressBar.lblPanel(0).Caption = ""
frmProgressBar.lblPanel(1).Caption = ""
DoEvents

'LockWindowUpdate Me.hWnd
Call CreateList
If bCancel Then GoTo ReEnable:

If lvMonsterIndex.ListItems.Count = 0 Then GoTo ReEnable:

ReEnable:
Unload frmProgressBar
frmMain.Enabled = True
'LockWindowUpdate 0&
Me.WindowState = vbNormal
Me.Show
Me.SetFocus

Exit Sub
error:
Call HandleError
Unload frmProgressBar
frmMain.Enabled = True
'LockWindowUpdate 0&
Me.WindowState = vbNormal
End Sub

Private Sub JustCreateList()
On Error GoTo error:

bCancel = False
lvMonsterIndex.ListItems.clear

frmProgressBar.sCaption = "Building Monster Group/Index List"
frmProgressBar.lblCaption.Caption = "Previous List Found, Building Table ..."
frmProgressBar.cmdCancel.Enabled = True
Call frmProgressBar.SetRange(39)
frmProgressBar.Show
frmMain.Enabled = False
DoEvents

LockWindowUpdate lvMonsterIndex.hwnd

frmProgressBar.lblPanel(0).Caption = ""
frmProgressBar.lblPanel(1).Caption = ""
Call CreateList
If bCancel Then GoTo ReEnable:

If lvMonsterIndex.ListItems.Count = 0 Then GoTo ReEnable:

ReEnable:
Unload frmProgressBar
frmMain.Enabled = True
LockWindowUpdate 0&
'Me.SetFocus

Exit Sub
error:
Call HandleError
Unload frmProgressBar
frmMain.Enabled = True
LockWindowUpdate 0&
End Sub

Private Sub ScanMonsters()
On Error GoTo error:
Dim nStatus As Integer, x As Integer

nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    MsgBox "Could not get first monster record, Error: " & BtrieveErrorCode(nStatus), vbOKOnly, "Creating Monster Group/Index List"
    Exit Sub
End If

Do While nStatus = 0 And Not bCancel
    MonsterRowToStruct Monsterdatabuf.buf
    
    frmProgressBar.lblPanel(1).Caption = Monsterrec.Number
    Call frmProgressBar.IncreaseProgress
    
    If Left(Monsterrec.Name, 3) = "sdf" Then GoTo Skip:

    If Monsterrec.Index < 0 Then GoTo Skip:
    
    If UBound(MGIL(), 2) < Monsterrec.Index Then ReDim Preserve MGIL(UBound(MGIL(), 1), Monsterrec.Index)
    'If UBound(MGIL(), 3) < Monsterrec.Number Then ReDim Preserve MGIL(UBound(MGIL(), 1), UBound(MGIL(), 2), Monsterrec.Number)
    
    For x = 0 To 10 '20
        If MGIL(Monsterrec.Group, Monsterrec.Index).nNumber(x) = 0 Then
            MGIL(Monsterrec.Group, Monsterrec.Index).nNumber(x) = Monsterrec.Number
            'MGIL(Monsterrec.Group, Monsterrec.Index).sName(x) = ClipNull(Monsterrec.Name)
            GoTo Skip:
        End If
    Next
    
Skip:
    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
    If Not bUseCPU Then DoEvents
Loop

Exit Sub

error:
Call HandleError

End Sub

'Private Sub ScanRooms()
'Dim nStatus As Integer, nRec As Long, x As Integer
'
'nStatus = BTRCALL(BGETFIRST, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
'If Not nStatus = 0 Then
'    MsgBox "Rooms: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
'    Exit Sub
'End If
'
'Do While nStatus = 0 And Not bCancel
'    RoomRowToStruct Roomdatabuf.buf
'
'    nRec = nRec + 1
'    frmProgressBar.lblPanel(1).Caption = nRec
'    Call frmProgressBar.IncreaseProgress
'
'    If Roomrec.MinIndex < 0 Or Roomrec.MaxIndex < 0 Then GoTo Skip:
'
'    If UBound(MGIL(), 2) < Roomrec.MinIndex Then ReDim Preserve MGIL(UBound(MGIL(), 1), Roomrec.MinIndex, UBound(MGIL(), 3))
'    If UBound(MGIL(), 2) < Roomrec.MaxIndex Then ReDim Preserve MGIL(UBound(MGIL(), 1), Roomrec.MaxIndex, UBound(MGIL(), 3))
'    If UBound(MGIL(), 3) < Roomrec.MaxIndex Then ReDim Preserve MGIL(UBound(MGIL(), 1), Roomrec.MaxIndex, UBound(MGIL(), 3))
'
''    For x = Roomrec.MinIndex To Roomrec.MaxIndex
''        MGIL(Roomrec.MonsterType, x).Used = True
''    Next x
'
'Skip:
'    nStatus = BTRCALL(BGETNEXT, RoomPosBlock, Roomdatabuf, Len(Roomdatabuf), ByVal RoomKeyBuffer, KEY_BUF_LEN, 0)
'    If Not bUseCPU Then DoEvents
'Loop
'
'End Sub

Private Sub CreateList()
Dim oLI As ListItem, x As Integer, y As Integer, z As Integer, sMonsters As String

On Error GoTo error:

For x = LBound(MGIL(), 1) To UBound(MGIL(), 1) 'group
    Call frmProgressBar.IncreaseProgress
    For y = LBound(MGIL(), 2) To UBound(MGIL(), 2) 'index
        
        sMonsters = ""
        
        For z = 0 To 10 '20 'monsters within
            If Not MGIL(x, y).nNumber(z) = 0 Then
                If Not sMonsters = "" Then sMonsters = sMonsters & ", "
                '& "(" & z & ")"
                sMonsters = sMonsters & GetMonsterName(MGIL(x, y).nNumber(z)) & "(" & MGIL(x, y).nNumber(z) & ")"
                If z = 20 Then sMonsters = sMonsters & " + More"
            End If
        Next z
        
        If Not sMonsters = "" Then
            Set oLI = lvMonsterIndex.ListItems.add()
            oLI.Text = GetMonGroupName(x) & "/" & y
            oLI.ListSubItems.add (1), "GroupIndex", sMonsters
        End If
        If Not bUseCPU Then DoEvents
        If bCancel Then Exit Sub
    Next y
Next x

'nStatus = BTRCALL(BGETFIRST, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
'If Not nStatus = 0 Then
'    MsgBox "Monsters: Could not get first record, Error: " & BtrieveErrorCode(nStatus)
'    Exit Sub
'End If

'Do While nStatus = 0 And Not bCancel
'    MonsterRowToStruct Monsterdatabuf.buf
'
'    frmProgressBar.lblPanel(1).Caption = Monsterrec.Number
'    Call frmProgressBar.IncreaseProgress
'
'    'no name, skip
'    sName = ClipNull(Monsterrec.Name)
'    If sName = "" Then GoTo Skip:
'
'    If UBound(MGIL(), 2) < Monsterrec.Index Then ReDim Preserve MGIL(UBound(MGIL(), 1), Monsterrec.Index, UBound(MGIL(), 3))
'
'    'add it
'    Set oLI = lvMonsterIndex.ListItems.add()
'    oLI.Text = Monsterrec.Number
'
'    oLI.ListSubItems.add (1), "Name", sName
'    oLI.ListSubItems.add (2), "GroupIndex", GetMonGroupName(Monsterrec.Group) & "/" & Monsterrec.Index
'    'oLI.ListSubItems.add (3), "MapRoom", MGIL(Monsterrec.Group, Monsterrec.Index)
'
'Skip:
'    nStatus = BTRCALL(BGETNEXT, MonsterPosBlock, Monsterdatabuf, Len(Monsterdatabuf), ByVal MonsterKeyBuffer, KEY_BUF_LEN, 0)
'    If Not bUseCPU Then DoEvents
'Loop

Set oLI = Nothing

out:
Exit Sub
error:
Call HandleError("CreateList")
Resume out:
End Sub
Public Sub ToggleStopBuild()
bCancel = True
End Sub

Private Sub AddColumnHeaders()

lvMonsterIndex.ColumnHeaders.clear
lvMonsterIndex.ColumnHeaders.add 1, "GroupIndex", "Group/Index", 1200, lvwColumnLeft
lvMonsterIndex.ColumnHeaders.add 2, "Monsters", "Monsters", 10000, lvwColumnLeft
'lvMonsterIndex.ColumnHeaders.add 3, "Name", "Name", 2400, lvwColumnCenter
'lvMonsterIndex.ColumnHeaders.add 4, "MapRoom", "Map/Room Monster Regens In", 6000, lvwColumnLeft

End Sub
Private Function CalcTotalRecords() As Long
On Error GoTo error:
Dim nStatus As Integer

CalcTotalRecords = 0

'nStatus = BTRCALL(BSTAT, RoomPosBlock, DBStatDatabuf, Len(Roomdatabuf), 0, KEY_BUF_LEN, 0)
'If Not nStatus = 0 Then
'    CalcTotalRecords = CalcTotalRecords + 30000
'Else
'    DBStatRowToStruct DBStatDatabuf.buf
'    CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
'End If

nStatus = BTRCALL(BSTAT, MonsterPosBlock, DBStatDatabuf, Len(Monsterdatabuf), 0, KEY_BUF_LEN, 0)
If Not nStatus = 0 Then
    CalcTotalRecords = CalcTotalRecords + 3000
Else
    DBStatRowToStruct DBStatDatabuf.buf
    CalcTotalRecords = CalcTotalRecords + DBStat.nRecords
End If


If CalcTotalRecords <= 0 Then CalcTotalRecords = 1
'If CalcTotalRecords > 32767 Then CalcTotalRecords = 32767

Exit Function

error:
Call HandleError
End Function

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo error:
Dim oLI As ListItem, str As String, x As Integer, nMaxGrpLen As Integer
Dim fso As FileSystemObject, nYesNo As Integer, sFile As TextStream

CommonDialog1.Filter = "TXT Files (*.txt)|*.txt"
CommonDialog1.DialogTitle = "Enter New File Name"
CommonDialog1.FileName = "NMR-MonsterIndex.txt"

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
sFile.WriteLine ("Monster Group/Index List -- " & Date & " @ " & Time)
sFile.WriteBlankLines (1)

nMaxGrpLen = 15
For Each oLI In lvMonsterIndex.ListItems
    
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
lvMonsterIndex.Width = Me.Width - 230
lvMonsterIndex.Height = Me.Height - TITLEBAR_OFFSET - 860

If Not lvMonsterIndex.ColumnHeaders.Count = 0 Then
    lvMonsterIndex.ColumnHeaders(2).Width = lvMonsterIndex.Width - 1700
End If

End Sub

Private Sub lvMonsterIndex_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'SortListView lvMonsterIndex, ColumnHeader.Index, ldtString, lvMonsterIndex.SortOrder
'If ColumnHeader.Index = 1 Then
'    SortListView lvMonsterIndex, ColumnHeader.Index, ldtNumber, lvMonsterIndex.SortOrder
'Else
'    SortListView lvMonsterIndex, ColumnHeader.Index, ldtString, lvMonsterIndex.SortOrder
'End If
End Sub


Public Sub CopyLine()

If lvMonsterIndex.SelectedItem Is Nothing Then Exit Sub

Clipboard.clear
Clipboard.SetText lvMonsterIndex.SelectedItem.Text & " -- " & lvMonsterIndex.SelectedItem.SubItems(1)

End Sub

Private Sub lvMonsterIndex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu frmMain.mnuMonsterIndexRightClick
End If
End Sub
