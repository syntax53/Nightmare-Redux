VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMonsterIndexChanger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monster Group/Index Changer"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9990
   Icon            =   "frmMonsterIndexChanger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   9990
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   8580
      TabIndex        =   29
      Top             =   3120
      Width           =   1155
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton cmdCopyClip 
         Caption         =   "Copy"
         Height          =   375
         Left            =   5880
         TabIndex        =   32
         Top             =   3120
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add -->"
         Height          =   315
         Left            =   5040
         TabIndex        =   31
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove Line"
         Height          =   315
         Left            =   8100
         TabIndex        =   30
         Top             =   60
         Width           =   1635
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start"
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
         Left            =   6780
         TabIndex        =   25
         Top             =   3120
         Width           =   1635
      End
      Begin VB.Frame fraRL 
         Caption         =   "List of Indexes From Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3555
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton cmdHideRL 
            Caption         =   "X"
            Height          =   375
            Left            =   4080
            TabIndex        =   19
            Top             =   3000
            Width           =   435
         End
         Begin VB.ListBox lstRecordList 
            Height          =   2595
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   4455
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Click the lines in the box to have the ranges pre-entered to the right."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   2880
            Width           =   3855
         End
      End
      Begin VB.Frame framFile 
         Caption         =   "Select Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3555
         Left            =   60
         TabIndex        =   11
         Top             =   0
         Width           =   4695
         Begin VB.CommandButton cmdWhy 
            Caption         =   "Why?"
            Height          =   375
            Left            =   3780
            TabIndex        =   33
            Top             =   3060
            Width           =   795
         End
         Begin VB.DirListBox Dir1 
            Height          =   2340
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   2295
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   2295
         End
         Begin VB.FileListBox filFileList 
            Height          =   2625
            Left            =   2460
            Pattern         =   "*.mdb"
            TabIndex        =   26
            ToolTipText     =   "Double Click to Open"
            Top             =   240
            Width           =   2115
         End
         Begin VB.CommandButton cmdListRecords 
            Caption         =   "List Indexes by Rooms"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   3060
            Width           =   3555
         End
         Begin VB.CommandButton cmdListRecords 
            Caption         =   "List Indexes by Monsters"
            Height          =   375
            Index           =   0
            Left            =   1020
            TabIndex        =   20
            Top             =   2760
            Visible         =   0   'False
            Width           =   2175
         End
      End
      Begin VB.TextBox txtStart 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   0
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtEnd 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   0
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtStart 
         Height          =   315
         Index           =   1
         Left            =   5040
         TabIndex        =   8
         Text            =   "1"
         Top             =   2220
         Width           =   735
      End
      Begin VB.TextBox txtEnd 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   1
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   2220
         Width           =   735
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C&lear"
         Height          =   315
         Left            =   6780
         TabIndex        =   6
         Top             =   60
         Width           =   1095
      End
      Begin VB.ListBox lstChange 
         Height          =   2595
         Left            =   6780
         TabIndex        =   5
         Top             =   420
         Width           =   2955
      End
      Begin VB.CommandButton cmdLog 
         Caption         =   "Log"
         Height          =   375
         Left            =   5040
         TabIndex        =   1
         Top             =   3120
         Width           =   735
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   3660
         Visible         =   0   'False
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.ComboBox cmbGroup 
         Height          =   315
         Index           =   1
         ItemData        =   "frmMonsterIndexChanger.frx":08CA
         Left            =   5040
         List            =   "frmMonsterIndexChanger.frx":0946
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1680
         Width           =   1575
      End
      Begin VB.ComboBox cmbGroup 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   0
         ItemData        =   "frmMonsterIndexChanger.frx":0AB5
         Left            =   5040
         List            =   "frmMonsterIndexChanger.frx":0B31
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   420
         Width           =   1575
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   4860
         X2              =   4860
         Y1              =   60
         Y2              =   3600
      End
      Begin VB.Label Label4 
         Caption         =   "Start"
         Height          =   195
         Left            =   5040
         TabIndex        =   18
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "End"
         Height          =   195
         Left            =   5880
         TabIndex        =   17
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Change This ..."
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
         Left            =   5040
         TabIndex        =   16
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label Label7 
         Caption         =   "Start"
         Height          =   195
         Left            =   5040
         TabIndex        =   15
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "End"
         Height          =   195
         Left            =   5880
         TabIndex        =   14
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "To This ..."
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
         Left            =   5040
         TabIndex        =   13
         Top             =   1440
         Width           =   900
      End
   End
   Begin MSComctlLib.StatusBar stsStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   4050
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15002
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMonsterIndexChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim DB As Database
Dim tabInfo As Recordset
Dim tabMonsters As Recordset
Dim tabRooms As Recordset

Dim nScale As Integer
Dim nScaleCount As Long

Dim bCancelProcess As Boolean
Dim DataSource As String
Dim fso As FileSystemObject
Dim ts As TextStream
Dim ChangeList() As Long
Dim nNextRange As Long


Private Sub cmdCopyClip_Click()
Dim sClip As String, x As Long

sClip = sClip & String(Len(sClip), "-") & vbCrLf
For x = 0 To lstChange.ListCount - 1
    sClip = sClip & lstChange.List(x) & vbCrLf
Next x

If Not sClip = "" Then
    Clipboard.clear
    Clipboard.SetText sClip
End If

End Sub

Private Sub cmdAdd_Click()
On Error GoTo error:
Dim x As Long, y As Long, sLine As String
Dim sGroup(1) As String, nIndex(1) As Long, z As Integer

If txtEnd(1).Text = "INVALID" Then Exit Sub
If txtStart(1).Text < 0 Then Exit Sub
If txtStart(0).Text < 0 Then Exit Sub
'If txtStart(1).Text = txtStart(0).Text Then Exit Sub

y = txtStart(1).Text

If lstChange.ListCount = 0 Then GoTo skip_verify:

For x = 0 To lstChange.ListCount - 1
    sLine = lstChange.List(x)
    
    z = InStr(1, sLine, " -> ")
    sGroup(0) = Left(sLine, InStr(1, sLine, "/") - 1)
'    sGroup(1) = Mid(sLine, z + 4, InStr(z, sLine, "/") - 4 - z)
    nIndex(0) = Val(Mid(sLine, InStr(1, sLine, "/") + 1, z - InStr(1, sLine, "/")))
'    nIndex(1) = Val(Mid(sLine, InStr(z + 4, sLine, "/") + 1))
    
    If nIndex(0) >= Val(txtStart(0).Text) And nIndex(0) <= Val(txtEnd(0).Text) And cmbGroup(0).Text = sGroup(0) Then
        MsgBox "Adding this range would include 'change from' indexes you've already added.", vbExclamation
        bCancelProcess = True
        Exit Sub
    End If

'    If nIndex(0) >= Val(txtStart(1).Text) And nIndex(0) <= Val(txtEnd(1).Text) Then
'        MsgBox "Adding this range would include 'change from' record numbers you've already added." _
'            & vbCrLf & "(Record #" & nIndex(0) & " is set to be changed to " & nIndex(1) & ")", vbExclamation
'        bCancelProcess = True
'        Exit Sub
'    ElseIf nIndex(1) >= Val(txtStart(1).Text) And nIndex(1) <= Val(txtEnd(1).Text) Then
'        MsgBox "Adding this range would include 'change to' record numbers you've already added." _
'            & vbCrLf & "(Record #" & nIndex(0) & " is set to be changed to " & nIndex(1) & ")", vbExclamation
'        bCancelProcess = True
'        Exit Sub
'    End If
Next x

skip_verify:
For x = txtStart(0).Text To txtEnd(0).Text
    lstChange.AddItem cmbGroup(0).Text & "/" & x & " -> " & cmbGroup(1).Text & "/" & y
    y = y + 1
Next

nNextRange = Val(txtEnd(1).Text) + 1
txtStart(1).Text = nNextRange

Call CalcRange

If lstRecordList.Visible Then
    If lstRecordList.ListIndex >= 0 Then
        If lstRecordList.ListIndex + 1 < lstRecordList.ListCount Then
            lstRecordList.ListIndex = lstRecordList.ListIndex + 1
        End If
    End If
End If

Exit Sub

error:
Call HandleError
Me.Enabled = True
End Sub

Private Sub cmdLog_Click()
Dim sFile As String
On Error GoTo error:

If Right(Dir1.Path, 1) = "\" Then
    sFile = Dir1.Path & "NMR-Log_IndexChange.txt"
Else
    sFile = Dir1.Path & "\NMR-Log_IndexChange.txt"
End If

If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(sFile) = False Then
    MsgBox sFile & " was not found.", vbInformation
    Exit Sub
End If

Call ShellExecute(0&, "open", sFile, vbNullString, vbNullString, vbNormalFocus)

out:
Exit Sub
error:
Call HandleError("cmdLog_Click")
Resume out:
End Sub

Private Sub cmdRemove_Click()
Dim nTemp As Long
If lstChange.ListIndex < 0 Then Exit Sub
nTemp = lstChange.ListIndex
lstChange.RemoveItem (nTemp)
If nTemp > 0 Then lstChange.ListIndex = nTemp - 1
End Sub


Private Sub cmdClear_Click()
Dim x As Integer
x = MsgBox("Are you sure?", vbYesNo + vbDefaultButton2 + vbQuestion)
If x = vbYes Then lstChange.clear
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdWhy_Click()
MsgBox "The indexes need to be changed by the rooms' restrictions because the rooms have ranges for the min/max index.  You can't say " _
    & """change monsters with index 1 to 10 and monsters with index 2 to 20"" because if a room has a range " _
    & "min-max index of 1-2 you can't change that to now suit the new index numbers (index 1 would become 10, index 2 would become 20, but the rooms' range would now be 10-20 and that may include other monsters).", vbInformation
End Sub

Private Sub Dir1_Change()
filFileList.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub filFileList_DblClick()
Dim fso As FileSystemObject

If filFileList.FileName = "" Then
    MsgBox "You must select a file first!", vbInformation + vbOKOnly
    Exit Sub
End If

Set fso = CreateObject("Scripting.FileSystemObject")

DataSource = filFileList.FileName
If Right(Dir1.Path, 1) = "/" Then
    DataSource = Dir1.Path & DataSource
Else
    DataSource = Dir1.Path & "/" & DataSource
End If


If fso.FileExists(DataSource) = True Then
    Call ShellExecute(0&, "open", DataSource, vbNullString, vbNullString, vbNormalFocus)
Else
    MsgBox DataSource & " was not found.", vbInformation
End If

Set fso = Nothing

End Sub


Private Sub cmdClose_Click()
Dim nYesNo As Integer
If cmdClose.Caption = "&Cancel" Then
    nYesNo = MsgBox("Are you sure you want to cancel?", vbQuestion + vbYesNo)
    If nYesNo = vbNo Then Exit Sub
    bCancelProcess = True
    DoEvents
    Exit Sub
End If

Unload Me
End Sub

Private Sub cmdStart_Click()
On Error GoTo error:
Dim x As Integer, sFile As String, y As Integer, z As Integer, sGroup(1) As String, nIndex(3) As Long
Dim StartTime As Long, nTotalTime As Double, sTotalTime As String, sLine As String
Dim nYesNo As Integer

'If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub

If lstChange.ListCount = 0 Then Exit Sub

bCancelProcess = False

'set the datasource
If filFileList.FileName = "" Then
    MsgBox "You must select an export file first.", vbInformation + vbOKOnly
    Exit Sub
End If

DataSource = filFileList.FileName
If Right(Dir1.Path, 1) = "\" Then
    DataSource = Dir1.Path & DataSource
    sFile = Dir1.Path & "NMR-Log_IndexChange.txt"
Else
    DataSource = Dir1.Path & "\" & DataSource
    sFile = Dir1.Path & "\NMR-Log_IndexChange.txt"
End If

nYesNo = MsgBox("Are you sure you want to change the record numbers in" & vbCrLf & DataSource, vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Action")
If nYesNo = vbNo Then Exit Sub

If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(DataSource) = False Then
    MsgBox DataSource & " was not found."
    Set fso = Nothing
    Exit Sub
End If

'disable stuff
cmdStart.Enabled = False
cmdClose.Caption = "&Cancel"
fraMain.Enabled = False

'set starting time
StartTime = Timer

'build change list array
ReDim ChangeList(lstChange.ListCount - 1, 1 To 4)
For x = 0 To lstChange.ListCount - 1
    sGroup(0) = ""
    sGroup(1) = ""
    nIndex(0) = 0
    nIndex(1) = 0
    nIndex(2) = 0
    nIndex(3) = 0
    
    sLine = lstChange.List(x)
    
    z = InStr(1, sLine, " -> ")
    sGroup(0) = Left(sLine, InStr(1, sLine, "/") - 1)
    sGroup(1) = Mid(sLine, z + 4, InStr(z, sLine, "/") - 4 - z)
    nIndex(0) = Val(Mid(sLine, InStr(1, sLine, "/") + 1, z - InStr(1, sLine, "/")))
    nIndex(1) = Val(Mid(sLine, InStr(z + 4, sLine, "/") + 1))

    For z = 0 To cmbGroup(0).ListCount - 1
        If sGroup(0) = cmbGroup(0).List(z) Then
            nIndex(2) = z
        End If
        If sGroup(1) = cmbGroup(0).List(z) Then
            nIndex(3) = z
        End If
    Next z
    
    ChangeList(x, 1) = nIndex(2) 'group from
    ChangeList(x, 2) = nIndex(0) 'index from
    ChangeList(x, 3) = nIndex(3) 'group to
    ChangeList(x, 4) = nIndex(1) 'index to
Next x

If bCancelProcess Then GoTo out:

'start log file
If fso.FileExists(sFile) Then
    fso.DeleteFile sFile, True
End If

Set ts = fso.OpenTextFile(sFile, ForWriting, True)

ts.WriteLine ("Record Number Change Job Started " & Date & " @ " & Time)
ts.WriteBlankLines (1)

ts.WriteLine ("ChangeList:")

ts.WriteLine ("-----------------------------------------------------")
For x = 0 To lstChange.ListCount - 1
    ts.WriteLine (lstChange.List(x))
Next x
ts.WriteBlankLines (1)

'GoTo out:

Set DB = OpenDatabase(DataSource)
nYesNo = OpenTables
If nYesNo < 0 Then GoTo out:

'set up progress bar
Call SetRange(CalcTotalRecords + lstChange.ListCount)
ProgressBar.Visible = True

bCancelProcess = False
DoEvents

Call ScanMonsters
If bCancelProcess Then GoTo out:
Call ScanRooms
If bCancelProcess Then GoTo out:

ts.Close

If bCancelProcess Then GoTo out:

ProgressBar.Value = ProgressBar.Max
DoEvents

nTotalTime = Timer - StartTime
sTotalTime = CStr(Round(CDbl(nTotalTime / 60), 2))
sTotalTime = Left(sTotalTime, InStr(1, sTotalTime, ".") + 2)

nYesNo = MsgBox("Change Complete, view log?" & vbCrLf & vbCrLf & "Total time: " & sTotalTime & " minutes.", vbInformation + vbYesNo, "Complete.")
If nYesNo = vbYes Then
    If fso.FileExists(sFile) = True Then
        Call ShellExecute(0&, "open", sFile, vbNullString, vbNullString, vbNormalFocus)
    Else
        MsgBox sFile & " was not found.", vbInformation
    End If
End If

out:
On Error Resume Next
ts.Close
Call CloseAll
stsStatusBar.Panels(1).Text = ""
stsStatusBar.Panels(2).Text = ""
ProgressBar.Visible = False
cmdStart.Enabled = True
cmdClose.Caption = "&Close"
fraMain.Enabled = True
Erase ChangeList()
Set ts = Nothing
Set fso = Nothing

Exit Sub
error:
Call HandleError("cmdStart_Click")
Resume out:

End Sub

Private Sub ScanMonsters()
Dim nStatus As Integer, x As Long, y As Long, nRec As Long, bUpdate As Boolean

'-------------------------------
'       MonsterS - SCAN
'-------------------------------
If tabMonsters.RecordCount = 0 Then
    ts.WriteLine vbCrLf & "Monsters -- No records to scan." & vbCrLf
    Exit Sub
End If

stsStatusBar.Panels(1).Text = "Monsters"

nRec = 0
tabMonsters.MoveFirst
Do Until tabMonsters.EOF Or bCancelProcess
    nRec = nRec + 1
    stsStatusBar.Panels(2).Text = nRec
    bUpdate = False
    
    For x = 0 To UBound(ChangeList(), 1)
        If tabMonsters.Fields("Group") = ChangeList(x, 1) Then
            If ChangeList(x, 2) = tabMonsters.Fields("Index") Then
                tabMonsters.Edit
                ts.WriteLine ("Monster " & ClipNull(tabMonsters.Fields("Name")) & " (" & tabMonsters.Fields("Number") & "): " _
                    & GetMonGroupName(ChangeList(x, 1)) & "/" & tabMonsters.Fields("Index") _
                    & " --> " & GetMonGroupName(ChangeList(x, 3)) & "/" & ChangeList(x, 4))
                tabMonsters.Fields("Group") = ChangeList(x, 3)
                tabMonsters.Fields("Index") = ChangeList(x, 4)
                bUpdate = True
                Exit For
            End If
        End If
    Next x
    
    If bUpdate Then tabMonsters.Update
    tabMonsters.MoveNext
    Call IncreaseProgressBar
    If Not bUseCPU Then DoEvents
Loop

End Sub

Private Sub ScanRooms()
Dim nStatus As Integer, x As Long, y As Long, nRec As Long, bUpdate As Boolean

'-------------------------------
'       ROOMS - SCAN
'-------------------------------
If tabRooms.RecordCount = 0 Then
    ts.WriteLine vbCrLf & "Rooms -- No records to scan." & vbCrLf
    Exit Sub
End If

stsStatusBar.Panels(1).Text = "Rooms"

nRec = 0
tabRooms.MoveFirst
Do Until tabRooms.EOF Or bCancelProcess
    nRec = nRec + 1
    stsStatusBar.Panels(2).Text = nRec
    bUpdate = False
    
    For x = 0 To UBound(ChangeList(), 1)
        If tabRooms.Fields("Mon Type") = ChangeList(x, 1) Then
            If ChangeList(x, 2) >= tabRooms.Fields("Min Index") And ChangeList(x, 2) <= tabRooms.Fields("Max Index") Then
                tabRooms.Edit
                y = tabRooms.Fields("Max Index") - tabRooms.Fields("Min Index")
                ts.WriteLine ("Room " & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number") _
                    & " (" & ClipNull(tabRooms.Fields("Name")) & "): " & GetMonGroupName(ChangeList(x, 1)) & "/" & tabRooms.Fields("Min Index") & "-" & tabRooms.Fields("Max Index") _
                    & " --> " & GetMonGroupName(ChangeList(x, 3)) & "/" & ChangeList(x, 4) & "-" & ChangeList(x, 4) + y)
                tabRooms.Fields("Mon Type") = ChangeList(x, 3)
                tabRooms.Fields("Min Index") = ChangeList(x, 4)
                tabRooms.Fields("Max Index") = ChangeList(x, 4) + y
                bUpdate = True
                Exit For
            End If
        End If
    Next x
    
    If bUpdate Then tabRooms.Update
    tabRooms.MoveNext
    Call IncreaseProgressBar
    If Not bUseCPU Then DoEvents
Loop

End Sub

Private Sub cmdListRecords_Click(Index As Integer)
Dim nInt As Integer, tabWorkingTable As Recordset, sText As String
Dim x As Long, y As Long, nTmpList() As String, sArr() As String, z As Integer
Dim nLastMax As Long, nLastMin As Long, nLastGroup As Integer, sRooms As String
On Error GoTo error:

If filFileList.FileName = "" Then
    MsgBox "You must select an export file first.", vbInformation + vbOKOnly
    Exit Sub
End If

DataSource = filFileList.FileName
If Right(Dir1.Path, 1) = "\" Then
    DataSource = Dir1.Path & DataSource
Else
    DataSource = Dir1.Path & "\" & DataSource
End If

Set DB = OpenDatabase(DataSource)
nInt = OpenTables
If nInt < 0 Then GoTo out:

Select Case Index
    Case 0: 'Monsters
        Set tabWorkingTable = tabMonsters
        If tabWorkingTable.RecordCount < 1 Then GoTo no_records:
        tabWorkingTable.Index = "pkMonsters"
    Case 1: 'Rooms
        Set tabWorkingTable = tabRooms
        If tabWorkingTable.RecordCount < 1 Then GoTo no_records:
        tabWorkingTable.Index = "idxRooms"
End Select

lstRecordList.clear
tabWorkingTable.MoveFirst

Call SetRange(tabWorkingTable.RecordCount)
ProgressBar.Visible = True

cmdStart.Enabled = False
cmdClose.Caption = "&Cancel"
fraMain.Enabled = False
'frmMain.Enabled = False

Me.MousePointer = vbHourglass
frmMain.MousePointer = vbHourglass

Call LockWindowUpdate(lstRecordList.hwnd)
'fraRL.Visible = False
DoEvents

Erase nTmpList()
ReDim nTmpList(39, 9999)

stsStatusBar.Panels(1).Text = tabWorkingTable.Name
bCancelProcess = False
Do While tabWorkingTable.EOF = False And bCancelProcess = False
    
    If Index = 0 Then 'monsters
        If tabMonsters.Fields("Group") > 39 Then
            MsgBox "Group Invalid on Monster #" & tabMonsters.Fields("Number") & " (" & ClipNull(tabMonsters.Fields("Name")) & ")", vbExclamation
            GoTo nextrec:
        End If
        If tabMonsters.Fields("Index") > 9999 Then ReDim Preserve nTmpList(39, tabMonsters.Fields("Index"))
        
        If Len(nTmpList(tabMonsters.Fields("Group"), tabMonsters.Fields("Index"))) < 30 Then
            If Len(nTmpList(tabMonsters.Fields("Group"), tabMonsters.Fields("Index"))) > 1 Then _
                nTmpList(tabMonsters.Fields("Group"), tabMonsters.Fields("Index")) = _
                    nTmpList(tabMonsters.Fields("Group"), tabMonsters.Fields("Index")) & ", "
            nTmpList(tabMonsters.Fields("Group"), tabMonsters.Fields("Index")) = _
                nTmpList(tabMonsters.Fields("Group"), tabMonsters.Fields("Index")) _
                & ClipNull(tabMonsters.Fields("Name")) & "(" & tabMonsters.Fields("Number") & ")"
            If Len(nTmpList(tabMonsters.Fields("Group"), tabMonsters.Fields("Index"))) >= 30 Then
                nTmpList(tabMonsters.Fields("Group"), tabMonsters.Fields("Index")) = _
                    nTmpList(tabMonsters.Fields("Group"), tabMonsters.Fields("Index")) & " ..."
            End If
        End If
    Else
        If tabRooms.Fields("Min Index") > 0 Or tabRooms.Fields("Max Index") > 0 Then
            If tabRooms.Fields("Mon Type") > 39 Then
                MsgBox "Mon Type Invalid on Room #" & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number"), vbExclamation
                GoTo nextrec:
            End If
            If tabRooms.Fields("Min Index") > 9999 Then
                ReDim Preserve nTmpList(39, tabRooms.Fields("Min Index"))
            End If
            If tabRooms.Fields("Max Index") > 9999 Then
                ReDim Preserve nTmpList(39, tabRooms.Fields("Max Index"))
            End If
            
            For x = tabRooms.Fields("Min Index") To tabRooms.Fields("Max Index")
                If Len(nTmpList(tabRooms.Fields("Mon Type"), x)) < 50 Then
                    If Len(nTmpList(tabRooms.Fields("Mon Type"), x)) > 1 Then
                        nTmpList(tabRooms.Fields("Mon Type"), x) = nTmpList(tabRooms.Fields("Mon Type"), x) & ", "
                    End If
                    nTmpList(tabRooms.Fields("Mon Type"), x) = nTmpList(tabRooms.Fields("Mon Type"), x) & tabRooms.Fields("Map Number") & "/" & tabRooms.Fields("Room Number")
'                Else
'                    If Not Right(nTmpList(tabRooms.Fields("Mon Type"), x), 3) = "..." Then
'                        nTmpList(tabRooms.Fields("Mon Type"), x) = nTmpList(tabRooms.Fields("Mon Type"), x) & " ..."
'                    End If
                End If
            Next x
        End If
    End If
    
nextrec:
    tabWorkingTable.MoveNext
    Call IncreaseProgressBar
    stsStatusBar.Panels(2).Text = Fix(tabWorkingTable.PercentPosition) & "%"
    If Not bUseCPU Then DoEvents
Loop

If bCancelProcess Then GoTo out:

nLastMin = -1
nLastMax = -1
sRooms = ""
For x = 0 To 39
    For y = 0 To UBound(nTmpList(), 2)
        If Len(nTmpList(x, y)) > 1 Then
            If Index = 0 Then 'monsters
                lstRecordList.AddItem GetMonGroupName(x) & "/" & y & ": " & nTmpList(x, y)
            Else
                nLastGroup = x
                If nLastMin = -1 Then
                    nLastMin = y
                    nLastMax = y
                    sRooms = nTmpList(x, y)
                Else
                    If y = nLastMax + 1 Then
                        nLastMax = y
                        If Len(sRooms) < 50 Then
                            sRooms = sRooms & ", " & nTmpList(x, y)
                        End If
                    Else
                        If Len(sRooms) > 30 Then
                            sArr() = Split(sRooms, ", ")
                            sRooms = ""
                            For z = 0 To UBound(sArr())
                                If InStr(1, sRooms, sArr(z)) = 0 Then
                                    If Len(sRooms) > 1 Then sRooms = sRooms & ", "
                                    sRooms = sRooms & sArr(z)
                                End If
                                If Len(sRooms) >= 30 Then
                                    sRooms = sRooms & " ..."
                                    Exit For
                                End If
                            Next z
                        End If
                        
                        If nLastMax > nLastMin Then
                            lstRecordList.AddItem GetMonGroupName(nLastGroup) & "/" & nLastMin & "-" & nLastMax & ": " & sRooms
                        Else
                            lstRecordList.AddItem GetMonGroupName(nLastGroup) & "/" & nLastMin & ": " & sRooms
                        End If
                        nLastMin = y
                        nLastMax = y
                        sRooms = nTmpList(x, y)
                    End If
                End If
            End If
        End If
    Next y
    If Index = 1 Then
        If nLastMin >= 0 And nLastMax >= 0 And nLastGroup >= 0 Then
            If Len(sRooms) > 30 Then
                sArr() = Split(sRooms, ", ")
                sRooms = ""
                For z = 0 To UBound(sArr())
                    If InStr(1, sRooms, sArr(z)) = 0 Then
                        If Len(sRooms) > 1 Then sRooms = sRooms & ", "
                        sRooms = sRooms & sArr(z)
                    End If
                    If Len(sRooms) >= 30 Then
                        sRooms = sRooms & " ..."
                        Exit For
                    End If
                Next z
            End If
            If nLastMax > nLastMin Then
                lstRecordList.AddItem GetMonGroupName(nLastGroup) & "/" & nLastMin & "-" & nLastMax & ": " & sRooms
            Else
                lstRecordList.AddItem GetMonGroupName(nLastGroup) & "/" & nLastMin & ": " & sRooms
            End If
        End If
        nLastMin = -1
        nLastMax = -1
        nLastGroup = -1
        sRooms = ""
    End If
Next x

ProgressBar.Value = ProgressBar.Max
fraRL.Visible = True

GoTo out:

no_records:
MsgBox "There are no records in that table.", vbInformation
fraRL.Visible = False
GoTo out:

out:
On Error Resume Next
Call LockWindowUpdate(0&)
Call CloseAll
Erase sArr()
stsStatusBar.Panels(1).Text = ""
stsStatusBar.Panels(2).Text = ""
ProgressBar.Visible = False
cmdStart.Enabled = True
fraMain.Enabled = True
cmdClose.Caption = "&Close"
frmMain.Enabled = True
Erase ChangeList()
Set tabWorkingTable = Nothing
Me.MousePointer = vbDefault
frmMain.MousePointer = vbDefault
DoEvents

Exit Sub
error:
Call HandleError("cmdListRecords_Click")
Resume out:

End Sub

Private Function CheckVersion() As Boolean
On Error GoTo error:
Dim nYesNo As Integer, sVer As String, sCurrentVer As String, sNMRVer As String

CheckVersion = False

If tabInfo.RecordCount = 0 Then
    nYesNo = MsgBox("Unable to verify export file version information, continue anyway?", vbYesNo + vbQuestion)
    If nYesNo = vbYes Then CheckVersion = True
    Exit Function
End If

tabInfo.MoveLast
sVer = tabInfo.Fields("Dat File Version")
sNMRVer = tabInfo.Fields("NMR Version")
sCurrentVer = FriendlyDatVersion(eDatFileVersion)

If Not sVer = sCurrentVer Or Not sNMRVer = sAppVersion Then
    nYesNo = MsgBox("Warning, current NMR Version/Dat File Version does not match the export file's versions." & vbCrLf _
        & "Current: " & sAppVersion & "/" & sCurrentVer & ", Export file: " & sNMRVer & "/" & sVer & vbCrLf & vbCrLf _
        & "Often the export database is updated and changed between releases as new fields are found." & vbCrLf _
        & "Errors may occur, Continue anyway?", vbYesNo + vbQuestion)
    If nYesNo = vbNo Then Exit Function
End If

CheckVersion = True

Exit Function
error:
Call HandleError
nYesNo = MsgBox("Unable to verify export file version information, continue anyway?", vbYesNo + vbQuestion)
If nYesNo = vbYes Then CheckVersion = True
End Function

Private Sub CloseAll()
On Error Resume Next

tabRooms.Close
tabInfo.Close
tabMonsters.Close

DB.Close

Set tabRooms = Nothing
Set tabMonsters = Nothing
Set tabInfo = Nothing

Set DB = Nothing

End Sub

Private Function OpenTables() As Integer
On Error GoTo error:

OpenTables = -1

Set tabRooms = DB.OpenRecordset("Rooms")
Set tabMonsters = DB.OpenRecordset("Monsters")
Set tabInfo = DB.OpenRecordset("Info")

tabRooms.Index = "idxRooms"
tabMonsters.Index = "pkMonsters"
        
OpenTables = 1

Exit Function
error:
Call HandleError
End Function

Private Sub cmdHideRL_Click()
fraRL.Visible = False
End Sub

Private Sub Form_Load()
Me.Top = ReadINI("Windows", "IdxChgTop")
Me.Left = ReadINI("Windows", "IdxChgLeft")

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FolderExists(ReadINI("Options", "ImportPath")) = True Then
    Dir1.Path = ReadINI("Options", "ImportPath")
Else
    Dir1.Path = App.Path
End If

Erase ChangeList()
Call AutoSizeDropDownWidth(cmbGroup(0))
Call AutoSizeDropDownWidth(cmbGroup(1))
Call ExpandCombo(cmbGroup(0), HeightOnly, TripleWidth, fraMain.hwnd)
Call ExpandCombo(cmbGroup(1), HeightOnly, TripleWidth, fraMain.hwnd)
Me.Show
cmbGroup(0).ListIndex = 0
cmbGroup(1).ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If cmdClose.Caption = "&Cancel" Then
    Cancel = 1
    Call cmdClose_Click
    Exit Sub
End If

Erase ChangeList()
Call WriteINI("Options", "ImportPath", Dir1.Path)

If Not Me.WindowState = vbMinimized Then
    Call WriteINI("Windows", "IdxChgTop", Me.Top)
    Call WriteINI("Windows", "IdxChgLeft", Me.Left)
End If

Set fso = Nothing
Set ts = Nothing
End Sub

Private Sub SetRange(ByVal MaxValue As Double)
Dim nNewMax As Integer

nScale = 0

If MaxValue > MaxInt Then
    If MaxValue / 2 < MaxInt Then
        nScale = 2
        nNewMax = MaxValue / 2
    ElseIf MaxValue / 4 < MaxInt Then
        nScale = 4
        nNewMax = MaxValue / 4
    ElseIf MaxValue / 8 < MaxInt Then
        nScale = 8
        nNewMax = MaxValue / 8
    ElseIf MaxValue / 10 < MaxInt Then
        nScale = 10
        nNewMax = MaxValue / 10
    Else
        MaxValue = MaxInt
    End If
Else
    nNewMax = MaxValue
End If

nNewMax = Fix(nNewMax)

nScaleCount = 1
ProgressBar.Value = 0
ProgressBar.Min = 0
ProgressBar.Max = nNewMax
End Sub

Private Sub IncreaseProgressBar()
On Error Resume Next
'If ProgressBar.Value + 1 < ProgressBar.Max Then ProgressBar.Value = ProgressBar.Value + 1

If nScale > 0 Then
    If nScaleCount = nScale Then
        If ProgressBar.Value + 1 < ProgressBar.Max Then ProgressBar.Value = ProgressBar.Value + 1
        nScaleCount = 1
    Else
        nScaleCount = nScaleCount + 1
    End If
Else
    If ProgressBar.Value + 1 < ProgressBar.Max Then ProgressBar.Value = ProgressBar.Value + 1
End If

End Sub

Private Function CalcTotalRecords() As Long
On Error GoTo error:
Dim nStatus As Integer

CalcTotalRecords = 0

CalcTotalRecords = CalcTotalRecords + tabRooms.RecordCount
CalcTotalRecords = CalcTotalRecords + tabMonsters.RecordCount

If CalcTotalRecords <= 0 Then CalcTotalRecords = 1
'If CalcTotalRecords > 32767 Then CalcTotalRecords = 32767

Exit Function

error:
Call HandleError
End Function

Private Sub CalcRange()
txtEnd(1).Text = Val(txtStart(1).Text) + (Val(txtEnd(0).Text) - Val(txtStart(0).Text))
If Val(txtEnd(1).Text) < Val(txtStart(1).Text) Then txtEnd(1).Text = "INVALID"
End Sub

Private Sub lstRecordList_Click()
Dim sLine As String, x As Integer, sTmp As String
On Error GoTo error:

If lstRecordList.ListCount < 1 Then Exit Sub
If lstRecordList.ListIndex < 0 Then Exit Sub

sLine = lstRecordList.List(lstRecordList.ListIndex)
If InStr(1, sLine, ":") <= 0 Then Exit Sub
If InStr(1, sLine, "/") <= 0 Then Exit Sub

sLine = Left(sLine, InStr(1, sLine, ":") - 1)

sTmp = Left(sLine, InStr(1, sLine, "/") - 1)
For x = 0 To 39
    If LCase(cmbGroup(0).List(x)) = LCase(sTmp) Then
        cmbGroup(0).ListIndex = x
        Exit For
    End If
Next x

sTmp = Right(sLine, Len(sLine) - InStr(1, sLine, "/"))
If InStr(1, sTmp, "-") > 0 Then
    txtStart(0).Text = Val(Left(sTmp, InStr(1, sLine, "-") - 1))
    txtEnd(0).Text = Val(Right(sTmp, Len(sLine) - InStr(1, sLine, "-")))
Else
    txtStart(0).Text = Val(sTmp)
    txtEnd(0).Text = Val(sTmp)
End If

Call CalcRange

out:
Exit Sub
error:
Call HandleError("lstRecordList_Click")
Resume out:
End Sub

Private Sub txtEnd_GotFocus(Index As Integer)
Call SelectAll(txtEnd(Index))
End Sub

Private Sub txtEnd_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Call CalcRange
End Sub

Private Sub txtStart_GotFocus(Index As Integer)
Call SelectAll(txtStart(Index))
End Sub

Private Sub txtStart_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Call CalcRange
End Sub
